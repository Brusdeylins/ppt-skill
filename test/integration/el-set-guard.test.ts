/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: el.set takes text only on shapes with a text body of their
**  own -- placeholders, textboxes and autoshapes. On a table, chart,
**  picture, connector or group it fails with E_SCHEMA at plan time instead
**  of silently doing nothing (or retexting a group's first child), and the
**  deck stays byte-identical.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { PptcError } from "../../src/infra/errors.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "el-set-guard.pptx")
const PNG = path.join(TMP, "el-set-guard.png")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  smallest valid 1x1 PNG  */
const PIXEL = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64")

/**  a group of two textboxes, crafted into the slide (el.add makes none)  */
const GROUP = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"900\" name=\"Gruppe\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
    + "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/>"
    + "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"914400\" cy=\"914400\"/></a:xfrm></p:grpSpPr>"
    + [ 901, 902 ].map((id) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Kind ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
        + "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm></p:spPr>"
        + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>child ${id}</a:t></a:r></a:p></p:txBody></p:sp>`).join("")
    + "</p:grpSp>"

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(PNG, PIXEL)
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    await executeOps(DECK, { ops: [
        { op: "slide.add", layout: "CONTENT", ref: "s", placeholders: { title: { text: "Guard" } } },
        { op: "el.add", slide: "$s", elements: [
            { type: "textbox", name: "Box", frame: { x: 0.5, y: 5.5, w: 3, h: 0.6 }, text: "box" },
            { type: "shape", name: "Form", shape: "roundRect", frame: { x: 4, y: 5.5, w: 2, h: 0.6 }, text: "form" },
            { type: "table", name: "Tab", frame: { x: 6.5, y: 5.5, w: 3 }, data: { rows: [ [ "a", "b" ] ] } },
            { type: "chart", name: "Diagramm", frame: { x: 10, y: 4, w: 3, h: 2 },
                data: { type: "column", categories: [ "x" ], series: [ { name: "s", values: [ 1 ] } ] } },
            { type: "image", name: "Bild", frame: { x: 0.5, y: 6.5, w: 0.5, h: 0.5 }, path: PNG },
            { type: "connector", name: "Linie", from: [ 2, 6.6 ], to: [ 4, 6.9 ] }
        ] }
    ] } as never, opts)

    /*  craft a group onto the slide  */
    const slide = (await readDeckState(await DeckArchive.open(DECK))).slides[0] as { part: string }
    const zip = await JSZip.loadAsync(readFileSync(DECK))
    const xml = await (zip.file(slide.part) as JSZip.JSZipObject).async("string")
    zip.file(slide.part, xml.replace("</p:spTree>", `${GROUP}</p:spTree>`))
    writeFileSync(DECK, await zip.generateAsync({ type: "nodebuffer" }))
})

describe("el.set type guard", () => {
    for (const [ name, type ] of [ [ "Tab", "table" ], [ "Diagramm", "chart" ], [ "Bild", "picture" ],
        [ "Linie", "connector" ], [ "Gruppe", "group" ] ]) {
        it(`rejects text on a ${type}`, async () => {
            const before = readFileSync(DECK)
            await expect(executeOps(DECK, { ops: [ { op: "el.set", slide: "0", name, text: "neu" } ] } as never, opts))
                .rejects.toSatisfy((err: unknown) => err instanceof PptcError && err.code === "E_SCHEMA"
                    && err.message.includes(`is a ${type}`))
            expect(readFileSync(DECK).equals(before)).toBe(true)
        })
    }

    it("keeps a group's children untouched", async () => {
        const shapes = (await readDeckState(await DeckArchive.open(DECK))).slides[0]?.shapes ?? []
        expect(shapes.find((sh) => sh.name === "Gruppe")?.type).toBe("group")
        const zip = await JSZip.loadAsync(readFileSync(DECK))
        const slide = (await readDeckState(await DeckArchive.open(DECK))).slides[0] as { part: string }
        expect(await (zip.file(slide.part) as JSZip.JSZipObject).async("string")).toContain("child 901")
    })

    it("still sets text on placeholders, textboxes and shapes", async () => {
        const title = (await readDeckState(await DeckArchive.open(DECK))).slides[0]?.shapes
            .find((sh) => sh.placeholderKind === "title")?.name as string
        await executeOps(DECK, { ops: [
            { op: "el.set", slide: "0", name: title, text: "Neuer Titel" },
            { op: "el.set", slide: "0", name: "Box", text: "neue box" },
            { op: "el.set", slide: "0", name: "Form", text: "neue form" }
        ] } as never, opts)
        await expectIntact(DECK)
        const state = (await readDeckState(await DeckArchive.open(DECK))).slides[0]
        expect(state?.title).toBe("Neuer Titel")
        expect(state?.shapes.find((sh) => sh.name.startsWith("Box-"))?.text).toBe("neue box")
        expect(state?.shapes.find((sh) => sh.name.startsWith("Form-"))?.text).toBe("neue form")
    })
})
