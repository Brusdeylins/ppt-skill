/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: text layout. Textboxes and shapes take `valign` and
**  `margin`; paragraphs `spaceBefore`/`spaceAfter`, `lineSpacing` and
**  numbered bullets; runs `strike`, `superscript`, `subscript` and
**  `highlight` -- written the same by el.add (PptxGenJS) and by slide.fill
**  / el.set (DOM), and read back by `state`. A paragraph of several runs
**  keeps one paragraph property set; shape text takes rich text.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import JSZip from "jszip"
import type { Element } from "@xmldom/xmldom"
import { elements, firstElement, parseXml } from "../../src/engine/xml.js"
import type { ParaInfo } from "../../src/core/model.js"
import { apply, attrs, failure, freshDeck, shapeOf, stateOf } from "../util/elements.js"

let deck: string

beforeAll(async () => {
    deck = await freshDeck("text-layout.pptx")
})

/**  rich text with every paragraph and run option  */
const RICH = [
    { runs: [ { text: "H" }, { text: "2", subscript: true }, { text: "O", strike: true, highlight: "FFFF00" },
        { text: "n", superscript: true } ], bullet: "number", spaceBefore: 6, spaceAfter: 3, lineSpacing: 1.5 },
    { text: "plain", bullet: false }
]

/**  the paragraph facts of a text body: pPr children and run attributes  */
const facts = (scope: Element): unknown[] => elements(scope, "a:p").map((p) => ({
    pPr: elements(p, "a:pPr").length,
    spacing: [ "a:lnSpc", "a:spcBef", "a:spcAft" ].map((tag) => {
        const el = firstElement(p, tag)
        return el === null ? null : (firstElement(el, "a:spcPct") ?? firstElement(el, "a:spcPts"))?.getAttribute("val")
    }),
    number: firstElement(p, "a:buAutoNum") !== null,
    runs: elements(p, "a:rPr").map((rPr) => [ rPr.getAttribute("strike"), rPr.getAttribute("baseline"),
        firstElement(rPr, "a:highlight") === null ? null : "highlight" ])
}))

/**  the expected facts of RICH (PptxGenJS writes no strike/baseline
     attribute where none is set)  */
const EXPECTED = [
    { pPr: 1, spacing: [ "150000", "600", "300" ], number: true,
        runs: [ [ null, null, null ], [ null, "-25000", null ], [ "sngStrike", null, "highlight" ], [ null, "30000", null ] ] },
    { pPr: 1, spacing: [ null, null, null ], number: false, runs: [ [ null, null, null ] ] }
]

/**  the paragraphs state reads back  */
const READ: Partial<ParaInfo>[] = [
    { bullet: "number", spaceBefore: 6, spaceAfter: 3, lineSpacing: 1.5 },
    { bullet: false }
]

describe("text layout in el.add", () => {
    it("writes anchoring, margins, paragraph and run options of a textbox", async () => {
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "textbox", name: "Box", frame: { x: 1, y: 1, w: 4, h: 2 }, valign: "bottom",
                margin: [ 0.1, 0.2, 0.3, 0.4 ], text: RICH } ] } ])
        const { shape } = await shapeOf(deck, "Box")
        const bodyPr = firstElement(shape, "a:bodyPr")
        expect([ "anchor", "lIns", "tIns", "rIns", "bIns" ].map((a) => bodyPr?.getAttribute(a)))
            .toEqual([ "b", "91440", "182880", "274320", "365760" ])
        expect(facts(shape)).toEqual(EXPECTED)
        const read = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Box-"))
        expect(read?.paragraphs?.map((p) => ({ ...p, runs: undefined }))).toEqual(READ.map((p) => ({ ...p, runs: undefined })))
        expect(read?.paragraphs?.[0]?.runs.map((r) => [ r.strike, r.subscript, r.superscript, r.highlight ]))
            .toEqual([ [ undefined, undefined, undefined, undefined ], [ undefined, true, undefined, undefined ],
                [ true, undefined, undefined, "FFFF00" ], [ undefined, undefined, true, undefined ] ])
    })

    it("takes rich text and a uniform margin on a shape", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "shape", name: "Shape", shape: "roundRect", frame: { x: 5, y: 1, w: 3, h: 2 }, valign: "top", margin: 0,
                text: [ { runs: [ { text: "bold", bold: true }, { text: " normal" } ] }, { text: "second", alignment: "left" } ] } ] } ])
        const { shape } = await shapeOf(deck, "Shape")
        const bodyPr = firstElement(shape, "a:bodyPr")
        expect([ "anchor", "lIns", "tIns", "rIns", "bIns" ].map((a) => bodyPr?.getAttribute(a))).toEqual([ "t", "0", "0", "0", "0" ])
        expect(elements(shape, "a:p").map((p) => [ elements(p, "a:pPr").length, firstElement(p, "a:pPr")?.getAttribute("algn"),
            elements(p, "a:t").map((t) => t.textContent).join("") ])).toEqual([ [ 1, "ctr", "bold normal" ], [ 1, "l", "second" ] ])
    })

    it("retexts a shape planned in the same document with rich text", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "shape", name: "Later", shape: "rect", frame: { x: 1, y: 4, w: 3, h: 1 }, text: "old" } ] },
        { op: "el.set", slide: "0", name: "Later", text: [ { runs: [ { text: "new", strike: true } ] } ] } ])
        const { shape } = await shapeOf(deck, "Later")
        expect(attrs(shape, "a:rPr", "strike")).toEqual([ "sngStrike" ])
    })

    it("rejects superscript together with subscript", async () => {
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [ { type: "textbox", frame: { x: 1, y: 1, w: 1, h: 1 },
            text: [ { runs: [ { text: "x", superscript: true, subscript: true } ] } ] } ] } ])).toBe("E_SCHEMA")
    })
})

describe("text layout in placeholders", () => {
    it("writes the same paragraph and run options through slide.fill and el.set", async () => {
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { body: { text: RICH } } } ])
        const placeholderFacts = async (): Promise<unknown[]> => {
            const state = await stateOf(deck)
            const part = (state.slides.at(-1) as { part: string }).part
            const zip = await JSZip.loadAsync(readFileSync(deck))
            const doc = parseXml(await (zip.file(part) as JSZip.JSZipObject).async("string"))
            const body = elements(doc, "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("idx") === "1")
            return facts(body as Element)
        }
        expect(await placeholderFacts()).toEqual(EXPECTED)
        await apply(deck, [ { op: "el.set", slide: "1", name: "PptcPh-1", text: RICH } ])
        expect(await placeholderFacts()).toEqual(EXPECTED)
        const read = (await stateOf(deck)).slides[1]?.shapes.find((sh) => sh.name === "PptcPh-1")
        expect(read?.paragraphs?.map((p) => ({ ...p, runs: undefined }))).toEqual(READ.map((p) => ({ ...p, runs: undefined })))
    })
})
