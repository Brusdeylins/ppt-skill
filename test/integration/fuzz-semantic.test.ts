/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: seeded random apply sequences checked against an expected
**  model after EVERY apply -- not only "the file is intact", but every
**  slide still has its own title, notes, image, chart, embedded OLE object,
**  section and sldId. Runs on a copy of the fixture serialized like
**  PowerPoint does (CRLF after the XML declaration), and on the private
**  corporate templates if present. A long run is configured by
**  PPTC_FUZZ_ROUNDS (rounds per seed) and PPTC_FUZZ_SEEDS (comma-separated
**  seeds), e.g. `PPTC_FUZZ_ROUNDS=150 PPTC_FUZZ_SEEDS=1,2,3 npx vitest run
**  test/integration/fuzz-semantic.test.ts`.
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

/**  long-run configuration (defaults: the regular test run)  */
const ROUNDS = process.env["PPTC_FUZZ_ROUNDS"] === undefined ? null : Number(process.env["PPTC_FUZZ_ROUNDS"])
const SEEDS = process.env["PPTC_FUZZ_SEEDS"] === undefined ? null
    : process.env["PPTC_FUZZ_SEEDS"].split(",").map((x) => Number(x.trim()))
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

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
interface Expected { title: string, notes: string, image: string | null, chart: string | null, ole: string | null, id: number | null }

/**  what one slide actually shows  */
interface Actual { title: string, notes: string | null, id: number, image: string | null, chart: string | null, ole: string | null, section: number | null }

/**  what the deck actually shows, per slide in order  */
const actual = async (deck: string): Promise<Actual[]> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const sectionOf = new Map<number, number>()
    for (const [ k, m ] of [ ...pres.matchAll(/<p14:section\b[^>]*>([\s\S]*?)<\/p14:section>/g) ].entries())
        for (const id of (m[1] as string).matchAll(/<p14:sldId id="(\d+)"/g))
            sectionOf.set(Number(id[1]), k)
    const out: Actual[] = []
    for (const s of state.slides as { part: string, title: string | null, notes: string | null, id: number }[]) {
        const rels = await zip.file(`ppt/slides/_rels/${path.posix.basename(s.part)}.rels`)?.async("string") ?? ""
        let image: string | null = null
        let chart: string | null = null
        let ole: string | null = null
        for (const m of rels.matchAll(/Type="[^"]*\/(image|chart|oleObject)" Target="([^"]+)"/g)) {
            const part = path.posix.normalize(path.posix.join("ppt/slides", m[2] as string))
            const bytes = await (zip.file(part) as JSZip.JSZipObject).async("nodebuffer")
            if (m[1] === "image" && bytes.subarray(0, PIXEL.length).equals(PIXEL))
                image = bytes.subarray(PIXEL.length).toString()
            if (m[1] === "chart")
                chart = /<c:v>(S-[^<]+)<\/c:v>/.exec(bytes.toString())?.[1] ?? "?"
            if (m[1] === "oleObject")
                ole = bytes.toString()
        }
        out.push({ title: s.title ?? "", notes: s.notes, id: s.id, image, chart, ole, section: sectionOf.get(s.id) ?? null })
    }
    return out
}

/**  give the slide with a title an embedded OLE object, as PowerPoint writes it  */
const injectOle = async (deck: string, title: string, payload: string): Promise<void> => {
    const slide = (await readDeckState(await DeckArchive.open(deck))).slides.find((s) => s.title === title) as { part: string }
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const text = async (part: string): Promise<string> => await (zip.file(part) as JSZip.JSZipObject).async("string")
    const bin = `ppt/embeddings/fuzzOle-${title}.bin`
    const rels = `ppt/slides/_rels/${path.posix.basename(slide.part)}.rels`
    zip.file(bin, Buffer.from(payload))
    zip.file("[Content_Types].xml", (await text("[Content_Types].xml")).replace("</Types>",
        `<Override PartName="/${bin}" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>`))
    zip.file(rels, (await text(rels)).replace("</Relationships>",
        `<Relationship Id="rIdFuzzOle" Type="${REL}/oleObject" Target="../embeddings/fuzzOle-${title}.bin"/></Relationships>`))
    zip.file(slide.part, (await text(slide.part)).replace("</p:spTree>",
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"900\" name=\"Objekt\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        + "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></p:xfrm>"
        + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\">"
        + "<p:oleObj progId=\"Package\" r:id=\"rIdFuzzOle\"><p:embed/></p:oleObj>"
        + "</a:graphicData></a:graphic></p:graphicFrame></p:spTree>"))
    writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
}

/**  split the deck into two PowerPoint sections: the first slide, the rest  */
const injectSections = async (deck: string, ids: number[]): Promise<void> => {
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const refs = (list: number[]): string => list.map((id) => `<p14:sldId id="${id}"/>`).join("")
    const ext = "<p:ext uri=\"{521415D9-36F7-43E2-AB2F-B90AF26B5E84}\">"
        + "<p14:sectionLst xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\">"
        + `<p14:section name="Erster" id="{11111111-1111-1111-1111-111111111111}"><p14:sldIdLst>${refs(ids.slice(0, 1))}</p14:sldIdLst></p14:section>`
        + `<p14:section name="Rest" id="{22222222-2222-2222-2222-222222222222}"><p14:sldIdLst>${refs(ids.slice(1))}</p14:sldIdLst></p14:section>`
        + "</p14:sectionLst></p:ext>"
    zip.file("ppt/presentation.xml", pres.includes("<p:extLst>")
        ? pres.replace("<p:extLst>", `<p:extLst>${ext}`)
        : pres.replace("</p:presentation>", `<p:extLst>${ext}</p:extLst></p:presentation>`))
    writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
}

/**  expected section of every slide after an apply: slides that kept their
     relative order (the longest increasing run of their old positions, first
     run on ties) keep their section; new, copied and moved slides join the
     section of their predecessor, the first one when they lead the deck  */
const nextSections = (before: Actual[], order: number[]): Map<number, number> => {
    const oldPos = new Map(before.filter((s) => s.section !== null).map((s, i) => [ s.id, i ]))
    const kept = order.filter((id) => oldPos.has(id))
    const len = kept.map(() => 1)
    const prev = kept.map(() => -1)
    for (let i = 0; i < kept.length; i++)
        for (let j = 0; j < i; j++)
            if ((oldPos.get(kept[j] as number) as number) < (oldPos.get(kept[i] as number) as number) && (len[j] as number) + 1 > (len[i] as number)) {
                len[i] = (len[j] as number) + 1
                prev[i] = j
            }
    const stable = new Set<number>()
    for (let i = len.indexOf(Math.max(0, ...len)); i >= 0; i = prev[i] as number)
        stable.add(kept[i] as number)
    const sectionOf = new Map(before.map((s) => [ s.id, s.section ]))
    const out = new Map<number, number>()
    let current = 0
    for (const id of order) {
        if (stable.has(id))
            current = sectionOf.get(id) as number
        out.set(id, current)
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

    let prev: Actual[] | null = null
    for (let round = 0; round < rounds; round++) {
        const kind = model.length < 2 ? 0 : pick(10)
        const ops: unknown[] = []
        let oleFor: string | null = null
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
            oleFor = rnd() < 0.25 ? t : null
            model.splice(at, 0, { title: t, notes: `notes-${t}`, image: withImage ? t : null, chart: withChart ? `S-${t}` : null,
                ole: oleFor === null ? null : `ole-${t}`, id: null })
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
        if (oleFor !== null)
            await injectOle(deck, oleFor, `ole-${oleFor}`)
        const got = await actual(deck)
        const sections = prev !== null && prev.some((p) => p.section !== null) ? nextSections(prev, got.map((g) => g.id)) : null
        const want = model.map((m, i) => ({ title: m.title, notes: m.notes, image: m.image, chart: m.chart, ole: m.ole,
            id: m.id ?? (got[i] as Actual).id, section: sections?.get((got[i] as Actual).id) ?? null }))
        expect(got, `round ${round} (${JSON.stringify(ops).slice(0, 200)})`).toEqual(want)
        for (const [ i, m ] of model.entries())
            m.id = (got[i] as Actual).id
        if (sections === null && model.length >= 2)
            await injectSections(deck, got.map((g) => g.id))
        prev = await actual(deck)
        if (round % 10 === 9)
            await expectIntact(deck)
    }
    await expectIntact(deck)
}

describe("semantic fuzz across applies", () => {
    const long = (base: number): number => Math.max(base, (ROUNDS ?? 0) * (SEEDS?.length ?? 1) * 4000)

    it("keeps every slide's content on a PowerPoint-serialized template", async () => {
        const template = await powerPointStyled(FIXTURE)
        for (const seed of SEEDS ?? [ 7 ])
            await fuzz(template, `neutral-${seed}`, seed, ROUNDS ?? 40)
    }, long(120000))

    const privates = existsSync(PRIVATE)
        ? readdirSync(PRIVATE).filter((f) => f.endsWith(".potx")).map((f) => path.join(PRIVATE, f))
        : []
    it.skipIf(privates.length === 0)("keeps every slide's content on the private templates", async () => {
        for (const [ i, tpl ] of privates.entries())
            for (const seed of SEEDS ?? [ 11 ])
                await fuzz(tpl, `private-${i}-${seed}`, seed + i, ROUNDS ?? 20)
    }, long(300000) * Math.max(1, privates.length))
})
