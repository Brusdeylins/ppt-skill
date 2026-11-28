/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: every part of a slide the plan does not touch -- the slide,
**  its rels, notes, charts, workbooks, embeddings, SmartArt, tags, comments
**  and media -- keeps its exact bytes across an apply, whatever happens to
**  the other slides (edits, additions, copies, moves, removals).
*/

import { beforeAll, describe, expect, it } from "vitest"
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState, readTemplateInfo } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { expectIntact } from "../util/integrity.js"
import { fixedScenario, writeMedia, type Media } from "../util/scenarios.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const FIXTURE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const DEMO = path.join(here, "..", "..", "demo", "kuehlschrank-funktionsweise.pptx")
const PRIVATE = path.join(here, "..", "..", "private-templates")
const TMP = path.join(here, "..", "tmp", "untouched")

const opts: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  relationship types that leave a slide's own content  */
const FOREIGN = /\/(slideLayout|slideMaster|slide|notesMaster|theme)$/

/**  resolve a relationship target  */
const resolve = (base: string, target: string): string =>
    target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join(base, target))

/**  the parts of a slide's own content, with their bytes  */
const closure = async (zip: JSZip, slide: string): Promise<Map<string, Buffer>> => {
    const out = new Map<string, Buffer>()
    const queue = [ slide ]
    while (queue.length > 0) {
        const part = queue.pop() as string
        const file = zip.file(part)
        if (out.has(part) || file === null)
            continue
        out.set(part, await file.async("nodebuffer"))
        const relsPart = `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`
        const rels = zip.file(relsPart)
        if (rels === null)
            continue
        const text = await rels.async("string")
        out.set(relsPart, Buffer.from(text))
        for (const m of text.matchAll(/<Relationship\b[^>]*>/g))
            if (!m[0].includes("TargetMode=\"External\"") && !FOREIGN.test(/Type="([^"]+)"/.exec(m[0])?.[1] ?? ""))
                queue.push(resolve(path.posix.dirname(part), /Target="([^"]+)"/.exec(m[0])?.[1] ?? ""))
    }
    return out
}

/**  apply ops and assert that every slide except the touched ones kept
     every byte of its own content  */
const applyAndCheck = async (deck: string, ops: unknown[], touched: (ids: number[]) => number[]): Promise<void> => {
    const before = await readDeckState(await DeckArchive.open(deck))
    const zipBefore = await JSZip.loadAsync(readFileSync(deck))
    const keep = new Set(touched(before.slides.map((s) => s.id)))
    await executeOps(deck, { ops } as never, opts)
    await expectIntact(deck)
    const after = await readDeckState(await DeckArchive.open(deck))
    const zipAfter = await JSZip.loadAsync(readFileSync(deck))
    for (const slide of before.slides.filter((s) => !keep.has(s.id))) {
        const now = after.slides.find((s) => s.id === slide.id)
        if (now === undefined)
            continue
        expect(now.part, `slide ${slide.id} keeps its part name`).toBe(slide.part)
        for (const [ part, bytes ] of await closure(zipBefore, slide.part)) {
            const file = zipAfter.file(part)
            expect(file, `${part} of untouched slide ${slide.id} still exists`).not.toBeNull()
            expect((await (file as JSZip.JSZipObject).async("nodebuffer")).equals(bytes),
                `${part} of untouched slide ${slide.id} keeps its bytes (${JSON.stringify(ops).slice(0, 200)})`).toBe(true)
        }
    }
}

/**  one-op changes around kept slides: each touches at most one slide  */
const checkAround = async (deck: string, media: Media): Promise<void> => {
    const ids = (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.id)
    const [ a, b, c ] = [ ids[0] as number, ids[1 % ids.length] as number, ids[2 % ids.length] as number ]
    await applyAndCheck(deck, [ { op: "slide.fill", slide: `id:${b}`, notes: "touched notes" } ], () => [ b ])
    await applyAndCheck(deck, [ { op: "el.add", slide: `id:${a}`, elements: [
        { type: "image", name: "Pic", frame: { x: 1, y: 1, w: 1, h: 1 }, path: media.png2 },
        { type: "chart", name: "Chart", frame: { x: 3, y: 1, w: 3, h: 2 },
            data: { type: "line", categories: [ "a", "b" ], series: [ { name: "L", values: [ 1, 2 ] } ] } } ] } ], () => [ a ])
    await applyAndCheck(deck, [ { op: "slide.copy", slide: `id:${c}` } ], () => [])
    await applyAndCheck(deck, [ { op: "slide.add", layout: 0, at: 1 } ], () => [])
    await applyAndCheck(deck, [ { op: "slide.move", slide: `id:${a}`, to: 2 } ], () => [])
    await applyAndCheck(deck, [ { op: "slide.fill", slide: `id:${c}`, hidden: true } ], () => [ c ])
    await applyAndCheck(deck, [ { op: "slide.rm", slide: "index:1" } ], () => [])
    await applyAndCheck(deck, [ { op: "meta.props", set: { title: "untouched", custom: { k: "v" } } } ], () => [])
}

/**  the parts whose bytes differ between two decks (added and removed
     parts included)  */
const changedParts = async (before: JSZip, after: JSZip): Promise<string[]> => {
    const names = [ ...new Set([ ...Object.keys(before.files), ...Object.keys(after.files) ]) ]
    const changed: string[] = []
    for (const name of names.filter((n) => !n.endsWith("/")).sort()) {
        const a = before.file(name)
        const b = after.file(name)
        if (a === null || b === null || !(await a.async("nodebuffer")).equals(await b.async("nodebuffer")))
            changed.push(name)
    }
    return changed
}

/**  the 1.2 in-place ops, each confined to one slide (or to the section
     list of presentation.xml): the other slides keep every byte  */
const checkInPlaceOps = async (deck: string, media: Media): Promise<void> => {
    await executeOps(deck, { ops: [
        { op: "slide.add", layout: "Title Only", ref: "a", section: "S1", placeholders: { title: { text: "Eins" } } },
        { op: "el.add", slide: "$a", elements: [
            { type: "textbox", name: "Box", frame: { x: 0.5, y: 1.5, w: 4, h: 0.6 }, text: "Alpha einmalig" },
            { type: "shape", name: "Form", shape: "rect", frame: { x: 5, y: 1.5, w: 2, h: 1 }, text: "Form" },
            { type: "image", name: "Bild", frame: { x: 0.5, y: 3, w: 2, h: 2 }, path: media.png },
            { type: "table", name: "Tab", frame: { x: 3, y: 3, w: 4 }, data: { rows: [ [ "a", "b" ], [ "c", "d" ] ] } } ] },
        { op: "slide.add", layout: "CONTENT", ref: "b", section: "S2", placeholders: { title: { text: "Zwei" }, body: { text: "Beta" } } },
        { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Drei" }, body: { text: "Gamma" } } }
    ] } as never, opts)
    const ids = (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.id)
    const a = ids[ids.length - 3] as number
    const b = ids[ids.length - 2] as number
    const c = ids[ids.length - 1] as number
    await applyAndCheck(deck, [ { op: "text.replace", find: "Alpha einmalig", replace: "Omega" } ], () => [ a ])
    await applyAndCheck(deck, [ { op: "el.set", slide: `id:${a}`, name: "Box", text: "neu" } ], () => [ a ])
    await applyAndCheck(deck, [ { op: "el.set", slide: `id:${a}`, name: "Tab", cells: [ { row: 1, col: 1, text: "x" } ] } ], () => [ a ])
    await applyAndCheck(deck, [ { op: "el.set", slide: `id:${a}`, name: "Bild", image: media.png2 } ], () => [ a ])
    await applyAndCheck(deck, [ { op: "el.set", slide: `id:${a}`, name: "Form", frame: { x: 5.5, y: 2 } } ], () => [ a ])
    await applyAndCheck(deck, [ { op: "slide.fill", slide: `id:${b}`, slideNumber: true } ], () => [ b ])
    await applyAndCheck(deck, [ { op: "slide.fill", slide: `id:${b}`, date: true } ], () => [ b ])
    await applyAndCheck(deck, [ { op: "slide.fill", slide: `id:${c}`, section: "S3" } ], () => [])
    await applyAndCheck(deck, [ { op: "el.add", slide: `id:${a}`, elements: [
        { type: "shape", name: "Sprung", shape: "rect", frame: { x: 8, y: 6, w: 1, h: 0.5 }, text: "Drei",
            hyperlink: `slide:id:${c}` } ] } ], () => [ a ])

    /*  section ops change nothing but the section list  */
    for (const op of [ { op: "section.rename", from: "S2", to: "Mitte" }, { op: "section.rm", name: "Mitte" } ]) {
        const before = await JSZip.loadAsync(readFileSync(deck))
        await applyAndCheck(deck, [ op ], () => [])
        const changed = await changedParts(before, await JSZip.loadAsync(readFileSync(deck)))
        expect(changed.filter((p) => !/^ppt\/presentation\.xml$|^docProps\//.test(p)), `${op.op} changes only presentation.xml`).toEqual([])
        expect(changed).toContain("ppt/presentation.xml")
    }
}

describe("untouched slides keep their bytes", () => {
    let media: Media
    beforeAll(() => {
        media = writeMedia(TMP)
    })

    it("on a deck built from the fixture", async () => {
        const deck = path.join(TMP, "fixture.pptx")
        writeFileSync(deck, await buildEmptyDeck(FIXTURE))
        const info = await readTemplateInfo(await DeckArchive.open(FIXTURE))
        for (const step of fixedScenario(info.layouts, media))
            await executeOps(deck, { ops: step(await readDeckState(await DeckArchive.open(deck))) } as never, opts)
        await checkAround(deck, media)
    }, 120000)

    it("around the in-place ops on one slide (text.replace, el.set, number/date, sections, slide links)", async () => {
        mkdirSync(TMP, { recursive: true })
        const deck = path.join(TMP, "inplace.pptx")
        writeFileSync(deck, await buildEmptyDeck(FIXTURE))
        const info = await readTemplateInfo(await DeckArchive.open(FIXTURE))
        for (const step of fixedScenario(info.layouts, media))
            await executeOps(deck, { ops: step(await readDeckState(await DeckArchive.open(deck))) } as never, opts)
        await checkInPlaceOps(deck, media)
    }, 120000)

    it.skipIf(!existsSync(DEMO))("on the demo deck", async () => {
        const deck = path.join(TMP, "demo.pptx")
        copyFileSync(DEMO, deck)
        await checkAround(deck, media)
    }, 120000)

    const privates = existsSync(PRIVATE)
        ? readdirSync(PRIVATE).filter((f) => f.endsWith(".potx")).map((f) => path.join(PRIVATE, f))
        : []
    it.skipIf(privates.length === 0)("on decks built from the private templates", async () => {
        mkdirSync(TMP, { recursive: true })
        for (const [ i, tpl ] of privates.entries()) {
            const deck = path.join(TMP, `private-${i}.pptx`)
            writeFileSync(deck, await buildEmptyDeck(tpl))
            const info = await readTemplateInfo(await DeckArchive.open(tpl))
            for (const step of fixedScenario(info.layouts, media))
                await executeOps(deck, { ops: step(await readDeckState(await DeckArchive.open(deck))) } as never, opts)
            await checkAround(deck, media)
        }
    }, 300000)
})
