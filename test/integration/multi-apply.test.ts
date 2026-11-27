/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: identities that must survive a sequence of applies --
**  speaker notes stay with their slide, sldIds stay stable (so `id:`
**  selectors keep their target) and sections keep their members; plus the
**  content guard that turns any such slip into E_INTEGRITY.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { expectIntact } from "../util/integrity.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { verifyContent, type SlideExpect } from "../../src/engine/guard.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

const apply = async (deck: string, ops: unknown[]): Promise<void> => {
    await executeOps(deck, { ops } as never, opts)
}
const slidesOf = async (deck: string): Promise<{ id: number, title: string | null, notes: string | null }[]> =>
    (await readDeckState(await DeckArchive.open(deck))).slides
const addTitled = (title: string, at?: number, notes?: string): unknown =>
    ({ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: title } },
        ...(at !== undefined && { at }), ...(notes !== undefined && { notes }) })

/**  member sldIds per section name, in file order  */
const sectionsOf = async (deck: string): Promise<Record<string, number[]>> => {
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const out: Record<string, number[]> = {}
    for (const m of pres.matchAll(/<p14:section name="([^"]+)"[^>]*>([\s\S]*?)<\/p14:section>/g))
        out[m[1] as string] = [ ...(m[2] as string).matchAll(/<p14:sldId id="(\d+)"/g) ].map((x) => Number(x[1]))
    return out
}

beforeAll(() => {
    mkdirSync(TMP, { recursive: true })
})

describe("notes across applies", () => {
    it("keeps every slide's notes when slides are inserted in front", async () => {
        const deck = path.join(TMP, "multi-notes.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        for (const [ i, at ] of [ 0, 0, 1, 0 ].entries())
            await apply(deck, [ addTitled(`T${i + 1}`, at, `notes-T${i + 1}`) ])
        await expectIntact(deck)
        for (const s of await slidesOf(deck))
            expect(s.notes).toBe(`notes-${s.title as string}`)
    })

    it("replaces notes of a kept slide without leaving the old part behind", async () => {
        const deck = path.join(TMP, "multi-notes-replace.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await apply(deck, [ addTitled("A", undefined, "a1"), addTitled("B", undefined, "b1") ])
        await apply(deck, [ { op: "slide.fill", slide: "title:A", notes: "a2" } ])
        await apply(deck, [ addTitled("C", 0, "c1") ])
        await expectIntact(deck)
        expect((await slidesOf(deck)).map((s) => s.notes)).toEqual([ "c1", "a2", "b1" ])
        const zip = await JSZip.loadAsync(readFileSync(deck))
        expect(Object.keys(zip.files).filter((f) => /^ppt\/notesSlides\/[^/]+\.xml$/.test(f))).toHaveLength(3)
    })

    it("gives a copied slide its own notes", async () => {
        const deck = path.join(TMP, "multi-notes-copy.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await apply(deck, [ addTitled("A", undefined, "a1") ])
        await apply(deck, [ { op: "slide.copy", slide: "title:A", ref: "c" } ])
        const [ , copy ] = await slidesOf(deck)
        await apply(deck, [ { op: "slide.fill", slide: `id:${(copy as { id: number }).id}`, notes: "copy" } ])
        await expectIntact(deck)
        expect((await slidesOf(deck)).map((s) => s.notes)).toEqual([ "a1", "copy" ])
    })
})

describe("slide ids across applies", () => {
    it("keeps the ids of kept slides and never re-uses a removed id", async () => {
        const deck = path.join(TMP, "multi-ids.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await apply(deck, [ addTitled("A"), addTitled("B"), addTitled("C") ])
        const before = new Map((await slidesOf(deck)).map((s) => [ s.title, s.id ]))
        await apply(deck, [ addTitled("N", 0) ])
        await apply(deck, [ { op: "slide.rm", slide: `id:${before.get("A") as number}` } ])
        await apply(deck, [ { op: "slide.move", slide: `id:${before.get("C") as number}`, to: 0 } ])
        await apply(deck, [ addTitled("M") ])
        await expectIntact(deck)
        const after = await slidesOf(deck)
        expect(after.map((s) => s.title)).toEqual([ "C", "N", "B", "M" ])
        expect(after.find((s) => s.title === "B")?.id).toBe(before.get("B"))
        expect(after.find((s) => s.title === "C")?.id).toBe(before.get("C"))
        const maxOld = Math.max(...before.values())
        for (const t of [ "N", "M" ])
            expect((after.find((s) => s.title === t) as { id: number }).id).toBeGreaterThan(maxOld)
        expect(new Set(after.map((s) => s.id)).size).toBe(after.length)
    })
})

describe("sections across applies", () => {
    it("keeps members on removal and files new slides into the right section", async () => {
        const deck = path.join(TMP, "multi-sections.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await apply(deck, [ addTitled("A"), addTitled("B"), addTitled("C"), addTitled("D") ])
        const ids = (await slidesOf(deck)).map((s) => s.id)
        const ref = (n: number): string => `<p14:sldId id="${ids[n] as number}"/>`
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
        const ext = "<p:ext uri=\"{521415D9-36F7-43E2-AB2F-B90AF26B5E84}\">"
            + "<p14:sectionLst xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\">"
            + `<p14:section name="Intro" id="{11111111-2222-3333-4444-555555555555}"><p14:sldIdLst>${ref(0)}${ref(1)}</p14:sldIdLst></p14:section>`
            + `<p14:section name="Main" id="{66666666-2222-3333-4444-555555555555}"><p14:sldIdLst>${ref(2)}${ref(3)}</p14:sldIdLst></p14:section>`
            + "</p14:sectionLst></p:ext>"
        zip.file("ppt/presentation.xml", pres.includes("<p:extLst>")
            ? pres.replace("<p:extLst>", `<p:extLst>${ext}`)
            : pres.replace("</p:presentation>", `<p:extLst>${ext}</p:extLst></p:presentation>`))
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))

        await apply(deck, [ { op: "slide.rm", slide: "title:A" } ])
        await apply(deck, [ addTitled("E") ])
        await apply(deck, [ addTitled("F", 0) ])
        await expectIntact(deck)
        const byTitle = new Map((await slidesOf(deck)).map((s) => [ s.title, s.id ]))
        const id = (t: string): number => byTitle.get(t) as number
        expect(await sectionsOf(deck)).toEqual({
            Intro: [ id("F"), id("B") ],
            Main:  [ id("C"), id("D"), id("E") ]
        })
    })
})

describe("content guard", () => {
    it("reports notes and media that change on a slide no op touched", async () => {
        const deck = path.join(TMP, "multi-guard.pptx")
        const png = path.join(TMP, "multi-guard.png")
        writeFileSync(png, Buffer.from(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==", "base64"))
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await apply(deck, [ addTitled("A", undefined, "a1"), { op: "el.add", slide: "title:A", elements: [
            { type: "image", name: "Bild", frame: { x: 1, y: 1.5, w: 2, h: 2 }, path: png } ] } ])
        const before = await JSZip.loadAsync(readFileSync(deck))
        const part = ((await readDeckState(await DeckArchive.open(deck))).slides[0] as { part: string }).part
        const expects: SlideExpect[] = [ { source: part, label: "A", notes: false, assets: "same" } ]
        expect(await verifyContent(before, before, expects)).toEqual({ blocked: [], lost: [] })

        /*  tamper: swap the notes text and the image bytes of the untouched slide  */
        const after = await JSZip.loadAsync(readFileSync(deck))
        for (const f of Object.keys(after.files)) {
            if (/^ppt\/notesSlides\/[^/]+\.xml$/.test(f))
                after.file(f, (await (after.file(f) as JSZip.JSZipObject).async("string")).replace("a1", "zz"))
            if (/^ppt\/media\/.+\.png$/.test(f))
                after.file(f, Buffer.from("other"))
        }
        const { blocked } = await verifyContent(before, after, expects)
        expect(blocked.some((f) => f.includes("speaker notes"))).toBe(true)
        expect(blocked.some((f) => f.includes("images, charts"))).toBe(true)
        expect(await verifyContent(before, after, [ { ...expects[0] as SlideExpect, notes: true, assets: "any" } ]))
            .toEqual({ blocked: [], lost: [] })

        /*  content that merely vanishes is a loss (warning), not a block  */
        const gone = await JSZip.loadAsync(readFileSync(deck))
        const slideRels = `ppt/slides/_rels/${path.posix.basename(part)}.rels`
        gone.file(slideRels, (await (gone.file(slideRels) as JSZip.JSZipObject).async("string"))
            .replace(/<Relationship [^>]*\/image"[^>]*\/>/, ""))
        expect(await verifyContent(before, gone, expects)).toEqual({ blocked: [], lost: [ { slide: 0, what: "image" } ] })
    })
})

describe("plan-time checks and rev", () => {
    it("reports a missing image already in --dry-run", async () => {
        const deck = path.join(TMP, "multi-dryrun.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await apply(deck, [ addTitled("A") ])
        await expect(executeOps(deck, { ops: [ { op: "el.add", slide: "title:A", elements: [
            { type: "image", name: "X", frame: { x: 1, y: 1, w: 1, h: 1 }, path: path.join(TMP, "does-not-exist.png") } ] } ] } as never,
        { ...opts, dryRun: true })).rejects.toMatchObject({ code: "E_FILE" })
    })

    it("changes the rev when only the notes change", async () => {
        const deck = path.join(TMP, "multi-rev.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await apply(deck, [ addTitled("A", undefined, "n1") ])
        const before = (await readDeckState(await DeckArchive.open(deck))).rev
        await apply(deck, [ { op: "slide.fill", slide: "title:A", notes: "n2" } ])
        expect((await readDeckState(await DeckArchive.open(deck))).rev).not.toBe(before)
    })
})

describe("seed cache", () => {
    it("rebuilds the deck-derived seed when a layout of the deck changes (and rejects the stale template)", async () => {
        const deck = path.join(TMP, "multi-seed.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        const editLayout = async (edit: (xml: string) => string): Promise<void> => {
            const zip = await JSZip.loadAsync(readFileSync(deck))
            const layout = (await Promise.all(Object.keys(zip.files).filter((f) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(f))
                .map(async (f) => [ f, await (zip.file(f) as JSZip.JSZipObject).async("string") ] as const)))
                .find(([ , xml ]) => xml.includes("<p:cSld name=\"CONTENT\"")) as readonly [ string, string ]
            zip.file(layout[0], edit(layout[1]))
            writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
        }

        /*  extension lists early and late in the layout, as PowerPoint writes them  */
        await editLayout((xml) => xml
            .replace("<p:ph type=\"title\"/></p:nvPr>", "<p:ph type=\"title\"/><p:extLst><p:ext uri=\"{A}\"/></p:extLst></p:nvPr>")
            .replace("</p:sldLayout>", "<p:extLst><p:ext uri=\"{B}\"/></p:extLst></p:sldLayout>"))
        await apply(deck, [ addTitled("A") ])

        /*  a second add caches the seed under the key of the now settled deck  */
        await executeOps(deck, { ops: [ addTitled("A2") ] } as never, { ...opts, templatePath: null })

        /*  change the CONTENT layout's body placeholder idx 1 -> 31 in the deck  */
        await editLayout((xml) => xml.replace("<p:ph idx=\"1\"/>", "<p:ph idx=\"31\"/>"))

        const add = { ops: [ { op: "slide.add", layout: "CONTENT",
            placeholders: { title: { text: "B" }, body: { text: "neu" } } } ] } as never

        /*  the original template no longer matches the deck's layouts  */
        await expect(executeOps(deck, add, opts)).rejects.toMatchObject({ code: "E_TEMPLATE" })
        await executeOps(deck, add, { ...opts, templatePath: null })
        await expectIntact(deck)
        expect((await slidesOf(deck)).map((s) => s.title)).toEqual([ "A", "A2", "B" ])
        const b = (await readDeckState(await DeckArchive.open(deck))).slides[2] as { part: string }
        const xml = await ((await JSZip.loadAsync(readFileSync(deck))).file(b.part) as JSZip.JSZipObject).async("string")
        expect(xml).toMatch(/<p:ph idx="31"\/>[\s\S]*<a:t>neu<\/a:t>/)
    })
})
