/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: seeded random apply sequences checked against an expected
**  model after EVERY apply -- not only "the file is intact", but every
**  slide still has its own title, notes, image, chart, embedded OLE object,
**  section and sldId, and its textboxes and table carry what el.set (text,
**  frame, fill, alt text, image, table cells and rows) and text.replace
**  made of them. Runs on a copy of the fixture serialized like
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
import type { ShapeInfo } from "../../src/core/model.js"

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

/**  expected content of one slide; `boxes` maps a textbox name to
     "text|x|fill|altText" and the table to its cells  */
interface Expected { title: string, notes: string, image: string | null, chart: string | null, ole: string | null, id: number | null,
    section: string | null, boxes: Record<string, string> }

/**  what one slide actually shows  */
interface Actual { title: string, notes: string | null, id: number, image: string | null, chart: string | null, ole: string | null, section: string | null,
    boxes: Record<string, string> }

/**  the textboxes and the table of a slide as the fuzz model sees them  */
const boxesOf = (shapes: ShapeInfo[]): Record<string, string> => {
    const out: Record<string, string> = {}
    for (const sh of shapes) {
        const name = /^((?:Box|Tab)\d*)-[0-9a-f-]{36}$/.exec(sh.name)?.[1]
        if (name !== undefined)
            out[name] = name.startsWith("Tab") ? JSON.stringify(sh.table)
                : [ sh.text ?? "", sh.frame?.x ?? "", sh.fill ?? "", sh.altText ?? "" ].join("|")
    }
    return out
}

/**  the section names of the deck in order, as the file lists them  */
const sectionNames = async (deck: string): Promise<string[]> => {
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    return [ ...pres.matchAll(/<p14:section name="([^"]*)"/g) ].map((m) => m[1] as string)
}

/**  what the deck actually shows, per slide in order  */
const actual = async (deck: string): Promise<Actual[]> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const sectionOf = new Map<number, string>()
    for (const m of pres.matchAll(/<p14:section name="([^"]*)"[^>]*>([\s\S]*?)<\/p14:section>/g))
        for (const id of (m[2] as string).matchAll(/<p14:sldId id="(\d+)"/g))
            if (!sectionOf.has(Number(id[1])))
                sectionOf.set(Number(id[1]), m[1] as string)
    const out: Actual[] = []
    for (const s of state.slides) {
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
        out.push({ title: s.title ?? "", notes: s.notes, id: s.id, image, chart, ole, section: sectionOf.get(s.id) ?? null,
            boxes: boxesOf(s.shapes) })
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
const nextSections = (before: Actual[], order: number[], first: string | null): Map<number, string | null> => {
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
    const out = new Map<number, string | null>()
    let current = first
    for (const id of order) {
        if (stable.has(id))
            current = sectionOf.get(id) as string
        out.set(id, current)
    }
    return out
}

/**  the expected sections after a slide starts the section `name` (a
     fresh name): it runs up to the next section start, a section left
     without slides goes, a deck without sections gets "Default Section"
     for the slides before it  */
const startSection = (model: Expected[], order: string[], idx: number, name: string, added: boolean): void => {
    if (order.length === 0) {
        if (idx > 0)
            order.push("Default Section")
        order.push(name)
        model.forEach((m, i) => { m.section = i < idx ? "Default Section" : name })
        return
    }
    const from = added ? (idx > 0 ? (model[idx - 1] as Expected).section : null) : (model[idx] as Expected).section
    order.splice(from === null ? 0 : order.indexOf(from) + 1, 0, name)
    ;(model[idx] as Expected).section = name
    for (let j = idx + 1; from !== null && j < model.length && (model[j] as Expected).section === from; j++)
        (model[j] as Expected).section = name
    if (from !== null && !model.some((m) => m.section === from))
        order.splice(order.indexOf(from), 1)
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
    let order: string[] = []
    for (let round = 0; round < rounds; round++) {
        const kind = model.length < 2 ? 0 : pick(17)
        const ops: unknown[] = []
        let oleFor: string | null = null

        /*  whether the round sets the sections of the model itself  */
        let sectioned = false
        if (kind === 0 || kind === 6) {
            const t = title()
            const at = pick(model.length + 1)
            const section = model.length >= 2 && rnd() < 0.2 ? `Neu${round}` : null
            ops.push({ op: "slide.add", ref: "n", layout: 1, at, placeholders: { title: { text: t } }, notes: `notes-${t}`,
                ...(section !== null && { section }) })
            const withImage = rnd() < 0.5
            const withChart = rnd() < 0.3
            const withTable = rnd() < 0.3
            const elements: unknown[] = []
            const boxes: Record<string, string> = {}
            if (withImage)
                elements.push({ type: "image", name: "Bild", frame: { x: 1, y: 2, w: 1, h: 1 }, path: ensureImage(t) })
            if (withChart)
                elements.push({ type: "chart", name: "Diagramm", frame: { x: 3, y: 2, w: 4, h: 3 },
                    data: { type: "column", categories: [ "a", "b" ], series: [ { name: `S-${t}`, values: [ 1, 2 ] } ] } })
            if (withTable) {
                const rows = [ [ `c${t}-0`, `c${t}-1` ], [ `c${t}-2`, `c${t}-3` ] ]
                elements.push({ type: "table", name: "Tab", frame: { x: 8, y: 2, w: 3 }, data: { rows } })
                boxes["Tab"] = JSON.stringify(rows)
            }
            if (elements.length > 0)
                ops.push({ op: "el.add", slide: "$n", elements })
            oleFor = rnd() < 0.25 ? t : null
            model.splice(at, 0, { title: t, notes: `notes-${t}`, image: withImage ? t : null, chart: withChart ? `S-${t}` : null,
                ole: oleFor === null ? null : `ole-${t}`, id: null, section: null, boxes })

            /*  the others keep their sections, the new slide starts its own  */
            if (section !== null) {
                startSection(model, order, at, section, true)
                sectioned = true
            }
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
            model.splice(i + 1, 0, { ...src, title: t, id: null, section: null, boxes: { ...src.boxes } })
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
            let text = `Box ${round}`

            /*  sequence: a replace in the same document sees the new box  */
            if (rnd() < 0.3) {
                ops.push({ op: "text.replace", find: "box", replace: "Kiste" })
                text = `Kiste ${round}`
                for (const m of model)
                    for (const [ name, box ] of Object.entries(m.boxes))
                        if (name.startsWith("Box"))
                            m.boxes[name] = box.replace(/^Box/i, "Kiste")
            }
            target.boxes[`Box${round}`] = `${text}|8||`
        }
        else if (kind === 10 && model.some((m) => Object.keys(m.boxes).some((b) => b.startsWith("Box")))) {
            /*  el.set text, frame, fill and alt text of a textbox  */
            const target = model.filter((m) => Object.keys(m.boxes).some((b) => b.startsWith("Box")))[0] as Expected
            const name = Object.keys(target.boxes).filter((b) => b.startsWith("Box"))[0] as string
            const [ text, x, fill, alt ] = (target.boxes[name] as string).split("|")
            const set: Record<string, unknown> = {}
            const next = { text, x, fill, alt }
            if (rnd() < 0.5)
                next.text = set["text"] = `neu ${round}`
            if (rnd() < 0.5)
                set["frame"] = { x: Number(next.x = String(7 + pick(4) / 4)) }
            if (rnd() < 0.5)
                next.fill = set["fill"] = [ "FFC000", "00B050", "7030A0" ][pick(3)] as string
            if (rnd() < 0.5 || Object.keys(set).length === 0)
                next.alt = set["altText"] = `alt ${round}`
            ops.push({ op: "el.set", slide: `title:${target.title}`, name, ...set })
            target.boxes[name] = [ next.text, next.x, next.fill, next.alt ].join("|")
        }
        else if (kind === 11 && model.some((m) => m.image !== null)) {
            /*  el.set image: another picture in place  */
            const target = model.filter((m) => m.image !== null)[pick(model.filter((m) => m.image !== null).length)] as Expected
            const tag = `${target.title}-r${round}`
            ops.push({ op: "el.set", slide: `title:${target.title}`, name: "Bild", image: ensureImage(tag) })
            target.image = tag
        }
        else if (kind === 12 && model.some((m) => m.boxes["Tab"] !== undefined)) {
            /*  el.set table cells and rows  */
            const target = model.filter((m) => m.boxes["Tab"] !== undefined)[0] as Expected
            const rows = JSON.parse(target.boxes["Tab"] as string) as string[][]
            const row = pick(rows.length)
            const set: Record<string, unknown> = { cells: [ { row, col: 1, text: `x${round}` } ] }
            ;(rows[row] as string[])[1] = `x${round}`
            if (rnd() < 0.5) {
                set["addRows"] = { rows: [ [ `a${round}`, `b${round}` ] ] }
                rows.push([ `a${round}`, `b${round}` ])
            }
            else if (rows.length > 2) {
                set["removeRows"] = [ 0 ]
                rows.splice(0, 1)
                set["cells"] = [ { row: row === 0 ? 0 : row - 1, col: 1, text: `x${round}` } ]
                if (row === 0)
                    (rows[0] as string[])[1] = `x${round}`
            }
            ops.push({ op: "el.set", slide: `title:${target.title}`, name: "Tab", ...set })
            target.boxes["Tab"] = JSON.stringify(rows)
        }
        else if (kind === 14) {
            /*  a slide starts a section: a fresh name, or its own (no-op)  */
            const i = pick(model.length)
            const target = model[i] as Expected
            const own = target.section !== null && rnd() < 0.2
            const name = own ? target.section as string : `Teil${round}`
            ops.push({ op: "slide.fill", slide: `title:${target.title}`, section: name })
            if (!own)
                startSection(model, order, i, name, false)
            sectioned = true
        }
        else if (kind === 15 && order.length > 0) {
            const from = order[pick(order.length)] as string
            ops.push({ op: "section.rename", from, to: `Umbenannt${round}` })
            for (const m of model)
                if (m.section === from)
                    m.section = `Umbenannt${round}`
            order[order.indexOf(from)] = `Umbenannt${round}`
            sectioned = true
        }
        else if (kind === 16 && order.length > 0) {
            /*  its slides join the previous section, the next for the first  */
            const k = pick(order.length)
            const name = order[k] as string
            const into = (k > 0 ? order[k - 1] : order[k + 1]) ?? null
            ops.push({ op: "section.rm", name })
            for (const m of model)
                if (m.section === name)
                    m.section = into
            order.splice(k, 1)
            sectioned = true
        }
        else if (kind === 13) {
            /*  text.replace in the notes of one slide  */
            const target = model[pick(model.length)] as Expected
            ops.push({ op: "text.replace", find: "notes-T", replace: "notiz-T", slide: `title:${target.title}`, notes: true })
            target.notes = target.notes.replace(/notes-T/g, "notiz-T")
        }
        else
            ops.push({ op: "meta.props", set: { title: `round ${round}` } })

        await executeOps(deck, { ops } as never, opts)
        if (oleFor !== null)
            await injectOle(deck, oleFor, `ole-${oleFor}`)
        const got = await actual(deck)
        if (!sectioned) {
            const sections = prev !== null && order.length > 0 ? nextSections(prev, got.map((g) => g.id), order[0] ?? null) : null
            for (const [ i, m ] of model.entries())
                m.section = sections?.get((got[i] as Actual).id) ?? null
        }
        const want = model.map((m, i) => ({ title: m.title, notes: m.notes, image: m.image, chart: m.chart, ole: m.ole,
            id: m.id ?? (got[i] as Actual).id, section: m.section, boxes: m.boxes }))
        const where = `round ${round} (${JSON.stringify(ops).slice(0, 200)})`
        expect(got, where).toEqual(want)
        expect(await sectionNames(deck), where).toEqual(order)
        for (const [ i, m ] of model.entries())
            m.id = (got[i] as Actual).id
        if (order.length === 0 && model.length >= 2) {
            await injectSections(deck, got.map((g) => g.id))
            order = [ "Erster", "Rest" ]
            model.forEach((m, i) => { m.section = i === 0 ? "Erster" : "Rest" })
        }
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
