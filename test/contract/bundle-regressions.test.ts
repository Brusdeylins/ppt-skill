/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Contract: regressions found in 1.0.x hold in the BUILT bundle, not only
**  in the sources -- images added in separate applies overwriting each
**  other, OLE objects of two slides collapsing into one file, and a BOM
**  template's deck with an embedded Excel object, replayed through
**  dst/pptc.mjs. `npm test` builds first.
*/

import { describe, expect, it } from "vitest"
import { spawnSync } from "node:child_process"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"

const here = path.dirname(fileURLToPath(import.meta.url))
const BIN = path.join(here, "..", "..", "dst", "pptc.mjs")
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

/**  run the CLI bundle and parse its envelope  */
const run = (...argv: string[]): Record<string, unknown> =>
    JSON.parse(spawnSync("node", [ BIN, ...argv ], { encoding: "utf8" }).stdout) as Record<string, unknown>

describe("1.0.x regressions in the built bundle", () => {
    it("keeps three SVG images added in three applies apart", async () => {
        mkdirSync(TMP, { recursive: true })
        const deck = path.join(TMP, "bundle-regressions.pptx")
        expect(run("new", deck, "--template", TEMPLATE, "--force").ok).toBe(true)
        const colors = [ "red", "green", "blue" ]
        for (const c of colors) {
            const svg = path.join(TMP, `bundle-regressions-${c}.svg`)
            writeFileSync(svg, `<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="${c}"/></svg>`)
            const op = { op: "el.add", slide: "$s", elements: [ { type: "image", name: "B", frame: { x: 1, y: 1, w: 2, h: 2 }, path: svg } ] }
            const res = spawnSync("node", [ BIN, "apply", deck, "--ops", "-" ], { encoding: "utf8",
                input: JSON.stringify({ ops: [ { op: "slide.add", ref: "s", layout: "CONTENT" }, op ] }) })
            expect((JSON.parse(res.stdout) as { ok: boolean }).ok).toBe(true)
        }
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const fills = await Promise.all(Object.keys(zip.files).filter((f) => /^ppt\/media\/.+\.svg$/.test(f))
            .map(async (f) => /fill="(\w+)"/.exec(await (zip.file(f) as JSZip.JSZipObject).async("string"))?.[1]))
        expect(fills.sort()).toEqual([ ...colors ].sort())
    })

    it("keeps the embedded OLE objects of two slides apart", async () => {
        mkdirSync(TMP, { recursive: true })
        const deck = path.join(TMP, "bundle-regressions-ole.pptx")
        expect(run("new", deck, "--template", TEMPLATE, "--force").ok).toBe(true)
        const add = spawnSync("node", [ BIN, "apply", deck, "--ops", "-" ], { encoding: "utf8", input: JSON.stringify({ ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "A" } } },
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "B" } } } ] }) })
        expect((JSON.parse(add.stdout) as { ok: boolean }).ok).toBe(true)

        /*  one OLE object per slide, each with its own payload  */
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const slides = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort()
        for (const [ i, slide ] of slides.entries()) {
            zip.file(`ppt/embeddings/oleObject${i + 1}.bin`, Buffer.from(`payload-${i}`))
            const ct = await (zip.file("[Content_Types].xml") as JSZip.JSZipObject).async("string")
            zip.file("[Content_Types].xml", ct.replace("</Types>",
                `<Override PartName="/ppt/embeddings/oleObject${i + 1}.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>`))
            const relsPart = `ppt/slides/_rels/${path.posix.basename(slide)}.rels`
            const rels = await (zip.file(relsPart) as JSZip.JSZipObject).async("string")
            zip.file(relsPart, rels.replace("</Relationships>", "<Relationship Id=\"rId90\" "
                + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject\" "
                + `Target="../embeddings/oleObject${i + 1}.bin"/></Relationships>`))
            const xml = await (zip.file(slide) as JSZip.JSZipObject).async("string")
            zip.file(slide, xml.replace("</p:spTree>", "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"90\" name=\"Objekt\"/>"
                + "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></p:xfrm>"
                + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\">"
                + "<p:oleObj progId=\"Package\" r:id=\"rId90\"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame></p:spTree>"))
        }
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))

        expect(run("apply", deck, "-e", JSON.stringify({ op: "meta.props", set: { title: "x" } })).ok).toBe(true)
        const out = await JSZip.loadAsync(readFileSync(deck))
        const payloads: string[] = []
        for (const slide of Object.keys(out.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort()) {
            const rels = await (out.file(`ppt/slides/_rels/${path.posix.basename(slide)}.rels`) as JSZip.JSZipObject).async("string")
            const target = /\/oleObject" Target="\.\.\/embeddings\/([^"]+)"/.exec(rels)?.[1] as string
            payloads.push(await (out.file(`ppt/embeddings/${target}`) as JSZip.JSZipObject).async("string"))
        }
        expect(payloads.sort()).toEqual([ "payload-0", "payload-1" ])
    })

    it("edits a BOM template's deck with an embedded Excel object", async () => {
        mkdirSync(TMP, { recursive: true })
        const template = path.join(TMP, "bundle-regressions-bom.pptx")
        const tpl = await JSZip.loadAsync(readFileSync(TEMPLATE))
        for (const n of Object.keys(tpl.files).filter((f) => /\.(xml|rels)$/.test(f)))
            tpl.file(n, Buffer.concat([ Buffer.from([ 0xEF, 0xBB, 0xBF ]), await (tpl.file(n) as JSZip.JSZipObject).async("nodebuffer") ]))
        writeFileSync(template, await tpl.generateAsync({ type: "nodebuffer" }))
        const deck = path.join(TMP, "bundle-regressions-package.pptx")
        expect(run("new", deck, "--template", template, "--force").ok).toBe(true)
        expect(run("apply", deck, "-e", JSON.stringify({ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "A" } } })).ok).toBe(true)

        /*  an embedded Excel object: a package rel on the slide  */
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const slide = Object.keys(zip.files).find((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)) as string
        const relsPart = `ppt/slides/_rels/${path.posix.basename(slide)}.rels`
        zip.file("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx", Buffer.from("xlsx-payload"))
        const ct = await (zip.file("[Content_Types].xml") as JSZip.JSZipObject).async("string")
        zip.file("[Content_Types].xml", ct.replace("<Default ",
            "<Default Extension=\"xlsx\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\"/><Default "))
        zip.file(relsPart, (await (zip.file(relsPart) as JSZip.JSZipObject).async("string")).replace("</Relationships>",
            "<Relationship Id=\"rId90\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\" "
            + "Target=\"../embeddings/Microsoft_Excel_Worksheet.xlsx\"/></Relationships>"))
        zip.file(slide, (await (zip.file(slide) as JSZip.JSZipObject).async("string")).replace("</p:spTree>",
            "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"90\" name=\"Tabelle\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
            + "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></p:xfrm>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\">"
            + "<p:oleObj r:id=\"rId90\" progId=\"Excel.Sheet.12\"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame></p:spTree>"))
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))

        expect(run("apply", deck, "-e", JSON.stringify({ op: "meta.props", set: { title: "x" } })).ok).toBe(true)
        expect(run("apply", deck, "-e", JSON.stringify({ op: "el.rm", slide: "title:A", name: "Tabelle" })).ok).toBe(true)
        expect(run("verify", deck).ok).toBe(true)
    })
})
