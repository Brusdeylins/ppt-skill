/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: seeded random apply sequences checked against an expected
**  model after EVERY apply -- not only "the file is intact", but every
**  slide still has its own title, notes, image, chart and sldId. Runs on
**  a copy of the fixture serialized like PowerPoint does (CRLF after the
**  XML declaration), and on the private corporate templates if present.
*/

import { describe, expect, it } from "vitest"
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { expectIntact } from "../util/integrity.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const FIXTURE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const PRIVATE = path.join(here, "..", "..", "private-templates")
const TMP = path.join(here, "..", "tmp")

/**  smallest valid 1x1 PNG; a trailing tag after IEND makes each image distinct  */
const PIXEL = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64")

/**  deterministic PRNG (mulberry32)  */
const prng = (seed: number): (() => number) => () => {
    seed = (seed + 0x6D2B79F5) | 0
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
}

/**  the fixture re-serialized like PowerPoint: CRLF after every XML declaration  */
const powerPointStyled = async (file: string): Promise<string> => {
    mkdirSync(TMP, { recursive: true })
    const zip = await JSZip.loadAsync(readFileSync(file))
    for (const name of Object.keys(zip.files).filter((f) => /\.(xml|rels)$/.test(f)))
        zip.file(name, (await (zip.file(name) as JSZip.JSZipObject).async("string"))
            .replace(/^(<\?xml[^>]*\?>)\n?/, "$1\r\n"))
    const out = path.join(TMP, "fuzz-semantic-template.pptx")
    writeFileSync(out, await zip.generateAsync({ type: "nodebuffer" }))
    return out
}

/**  expected content of one slide  */
interface Expected { title: string, notes: string, image: string | null, chart: string | null, id: number | null }

/**  what the deck actually shows, per slide in order  */
const actual = async (deck: string): Promise<{ title: string, notes: string | null, id: number, image: string | null, chart: string | null }[]> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const out = []
    for (const s of state.slides as { part: string, title: string | null, notes: string | null, id: number }[]) {
        const rels = await zip.file(`ppt/slides/_rels/${path.posix.basename(s.part)}.rels`)?.async("string") ?? ""
        let image: string | null = null
        let chart: string | null = null
        for (const m of rels.matchAll(/Type="[^"]*\/(image|chart)" Target="([^"]+)"/g)) {
            const part = path.posix.normalize(path.posix.join("ppt/slides", m[2] as string))
            const bytes = await (zip.file(part) as JSZip.JSZipObject).async("nodebuffer")
            if (m[1] === "image" && bytes.subarray(0, PIXEL.length).equals(PIXEL))
                image = bytes.subarray(PIXEL.length).toString()
            if (m[1] === "chart")
                chart = /<c:v>(S-[^<]+)<\/c:v>/.exec(bytes.toString())?.[1] ?? "?"
        }
        out.push({ title: s.title ?? "", notes: s.notes, id: s.id, image, chart })
    }
    return out
}

/**  run a seeded random session and compare after every apply  */
const fuzz = async (template: string, tag: string, seed: number, rounds: number): Promise<void> => {
    mkdirSync(TMP, { recursive: true })
    const opts: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }
    const deck = path.join(TMP, `fuzz-semantic-${tag}.pptx`)
    writeFileSync(deck, await buildEmptyDeck(template))
    const rnd = prng(seed)
    const pick = (n: number): number => Math.floor(rnd() * n)
    const model: Expected[] = []
    let serial = 0
    const title = (): string => `T${++serial}`
    const ensureImage = (t: string): string => {
        const f = path.join(TMP, `fuzz-semantic-${tag}-${t}.png`)
        writeFileSync(f, Buffer.concat([ PIXEL, Buffer.from(t) ]))
        return f
    }

    for (let round = 0; round < rounds; round++) {
        const kind = model.length < 2 ? 0 : pick(10)
        const ops: unknown[] = []
        if (kind === 0 || kind === 6) {
            const t = title()
            const at = pick(model.length + 1)
            ops.push({ op: "slide.add", ref: "n", layout: 1, at, placeholders: { title: { text: t } }, notes: `notes-${t}` })
            const withImage = rnd() < 0.5
            const withChart = rnd() < 0.3
            const elements: unknown[] = []
            if (withImage)
                elements.push({ type: "image", name: "Bild", frame: { x: 1, y: 2, w: 1, h: 1 }, path: ensureImage(t) })
            if (withChart)
                elements.push({ type: "chart", name: "Diagramm", frame: { x: 3, y: 2, w: 4, h: 3 },
                    data: { type: "column", categories: [ "a", "b" ], series: [ { name: `S-${t}`, values: [ 1, 2 ] } ] } })
            if (elements.length > 0)
                ops.push({ op: "el.add", slide: "$n", elements })
            model.splice(at, 0, { title: t, notes: `notes-${t}`, image: withImage ? t : null, chart: withChart ? `S-${t}` : null, id: null })
        }
        else if (kind === 1) {
            const i = pick(model.length)
            ops.push({ op: "slide.rm", slide: `title:${(model[i] as Expected).title}` })
            model.splice(i, 1)
        }
        else if (kind === 2) {
            const i = pick(model.length)
            const to = pick(model.length)
            ops.push({ op: "slide.move", slide: `title:${(model[i] as Expected).title}`, to })
            const [ moved ] = model.splice(i, 1)
            model.splice(to, 0, moved as Expected)
        }
        else if (kind === 3) {
            const i = pick(model.length)
            const src = model[i] as Expected
            const t = title()
            ops.push({ op: "slide.copy", slide: `title:${src.title}`, ref: "c" },
                { op: "slide.fill", slide: "$c", placeholders: { title: { text: t } } })
            model.splice(i + 1, 0, { ...src, title: t, id: null })
        }
        else if (kind === 4) {
            const target = model[pick(model.length)] as Expected
            target.notes = `notes-${target.title}-r${round}`
            ops.push({ op: "slide.fill", slide: `title:${target.title}`, notes: target.notes })
        }
        else if (kind === 7) {
            const target = model[pick(model.length)] as Expected
            ops.push({ op: "slide.fill", slide: `title:${target.title}`, footer: `Fuß ${round}` })
        }
        else if (kind === 8) {
            const target = model[pick(model.length)] as Expected
            ops.push({ op: "slide.fill", slide: `title:${target.title}`, hidden: rnd() < 0.5 })
        }
        else if (kind === 9) {
            const target = model[pick(model.length)] as Expected
            ops.push({ op: "el.add", slide: `title:${target.title}`, elements: [
                { type: "textbox", name: `Box${round}`, frame: { x: 8, y: 5, w: 2, h: 1 }, text: `Box ${round}` } ] })
        }
        else
            ops.push({ op: "meta.props", set: { title: `round ${round}` } })

        await executeOps(deck, { ops } as never, opts)
        const got = await actual(deck)
        const want = model.map((m, i) => ({ title: m.title, notes: m.notes, image: m.image, chart: m.chart,
            id: m.id ?? (got[i] as { id: number }).id }))
        expect(got, `round ${round} (${JSON.stringify(ops).slice(0, 200)})`).toEqual(want)
        for (const [ i, m ] of model.entries())
            m.id = (got[i] as { id: number }).id
        if (round % 10 === 9)
            await expectIntact(deck)
    }
    await expectIntact(deck)
}

describe("semantic fuzz across applies", () => {
    it("keeps every slide's content on a PowerPoint-serialized template (seed 7)", async () => {
        await fuzz(await powerPointStyled(FIXTURE), "neutral", 7, 40)
    }, 120000)

    const privates = existsSync(PRIVATE)
        ? readdirSync(PRIVATE).filter((f) => f.endsWith(".potx")).map((f) => path.join(PRIVATE, f))
        : []
    it.skipIf(privates.length === 0)("keeps every slide's content on the private templates", async () => {
        for (const [ i, tpl ] of privates.entries())
            await fuzz(tpl, `private-${i}`, 11 + i, 20)
    }, 300000)
})
