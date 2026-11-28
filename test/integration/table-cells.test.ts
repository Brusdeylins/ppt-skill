/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: table columns and cells. `colWidths` (inches, scaled to an
**  explicit frame width) round-trip with `state`; `colAlign` aligns a
**  column unless a cell sets its own `align`; styled cells carry rich text,
**  fill, color, bold, alignment and text direction; `colspan`/`rowspan`
**  merge cells HTML-like (covered cells are omitted), `state` reads the
**  merges back; a ragged grid or a per-column list of the wrong length
**  fails at plan time.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { elements, firstElement } from "../../src/engine/xml.js"
import { apply, attrs, failure, freshDeck, shapeOf, stateOf } from "../util/elements.js"

let deck: string

beforeAll(async () => {
    deck = await freshDeck("table-cells.pptx")
    await apply(deck, [ { op: "slide.add", layout: "DEFAULT" } ])
})

/**  the table shape of the last slide in the read model  */
const readTable = async (name: string): Promise<Record<string, unknown>> =>
    (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith(`${name}-`)) as unknown as Record<string, unknown>

describe("table columns", () => {
    it("writes column widths that state reads back", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [ { type: "table", name: "Widths",
            frame: { x: 1, y: 1 }, colWidths: [ 1, 2.5, 1.5 ],
            data: { headers: [ "A", "B", "C" ], rows: [ [ "1", "2", "3" ] ] } } ] } ])
        const read = await readTable("Widths")
        expect(read["colWidths"]).toEqual([ 1, 2.5, 1.5 ])
        expect((read["frame"] as { w: number }).w).toBe(5)
    })

    it("scales the widths to an explicit frame width", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [ { type: "table", name: "Scaled",
            frame: { x: 1, y: 3, w: 8 }, colWidths: [ 1, 3 ], data: { rows: [ [ "a", "b" ] ] } } ] } ])
        expect((await readTable("Scaled"))["colWidths"]).toEqual([ 2, 6 ])
    })

    it("aligns columns, a cell's own alignment wins", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [ { type: "table", name: "Aligned",
            frame: { x: 1, y: 4.5, w: 6 }, colAlign: [ "left", "center", "right" ],
            data: { headers: [ "L", "C", "R" ], rows: [ [ "1", "2", { text: "3", align: "left" } ] ] } } ] } ])
        const { shape } = await shapeOf(deck, "Aligned")
        const rows = elements(shape, "a:tr").map((tr) => elements(tr, "a:pPr").map((p) => p.getAttribute("algn")))
        expect(rows).toEqual([ [ "l", "ctr", "r" ], [ "l", "ctr", "l" ] ])
    })

    it("rejects per-column lists of the wrong length", async () => {
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [ { type: "table",
            frame: { x: 1, y: 1 }, colWidths: [ 1, 2 ], data: { rows: [ [ "a", "b", "c" ] ] } } ] } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [ { type: "table",
            frame: { x: 1, y: 1 }, colAlign: [ "left" ], data: { rows: [ [ "a", "b" ] ] } } ] } ])).toBe("E_SCHEMA")
    })
})

describe("table cells", () => {
    it("merges cells HTML-like and reads the merges back", async () => {
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "1", elements: [ { type: "table",
            name: "Merged", frame: { x: 1, y: 1, w: 6 },
            data: { headers: [ { text: "Both", colspan: 2 }, "C" ],
                rows: [ [ { text: "Tall", rowspan: 2 }, "b1", "c1" ], [ "b2", "c2" ] ] } } ] } ])
        const { shape } = await shapeOf(deck, "Merged")
        expect(attrs(shape, "a:tc", "gridSpan")).toContain("2")
        expect(attrs(shape, "a:tc", "hMerge")).toContain("1")
        expect(attrs(shape, "a:tc", "rowSpan")).toContain("2")
        expect(attrs(shape, "a:tc", "vMerge")).toContain("1")
        const read = (await stateOf(deck)).slides[1]?.shapes.find((sh) => sh.name.startsWith("Merged-"))
        expect(read?.table).toEqual([ [ "Both", "", "C" ], [ "Tall", "b1", "c1" ], [ "", "b2", "c2" ] ])
        expect(read?.merges).toEqual([ { row: 0, col: 0, rowspan: 1, colspan: 2 }, { row: 1, col: 0, rowspan: 2, colspan: 1 } ])
    })

    it("styles a cell: rich text, fill, color, bold, valign, text direction", async () => {
        await apply(deck, [ { op: "el.add", slide: "1", elements: [ { type: "table", name: "Styled",
            frame: { x: 1, y: 4, w: 6 }, textDirection: "rotate270",
            data: { rows: [ [ { text: [ { runs: [ { text: "rich", italic: true }, { text: " text", bold: true } ] } ],
                fill: "C00000", color: "FFFFFF", bold: false, valign: "bottom", textDirection: "stacked" }, "plain" ] ] } } ] } ])
        const { shape } = await shapeOf(deck, "Styled")
        const [ first, second ] = elements(shape, "a:tc")
        const tcPr = first === undefined ? null : firstElement(first, "a:tcPr")
        expect([ tcPr?.getAttribute("anchor"), tcPr?.getAttribute("vert") ]).toEqual([ "b", "wordArtVert" ])
        const fill = elements(tcPr ?? shape, "a:solidFill").find((f) => f.parentNode === tcPr)
        expect(fill === undefined ? null : firstElement(fill, "a:srgbClr")?.getAttribute("val")).toBe("C00000")
        expect(elements(first ?? shape, "a:t").map((t) => t.textContent)).toEqual([ "rich", " text" ])
        expect(second === undefined ? null : firstElement(second, "a:tcPr")?.getAttribute("vert")).toBe("vert270")
        const read = (await stateOf(deck)).slides[1]?.shapes.find((sh) => sh.name.startsWith("Styled-"))
        expect(read?.cellTextDirections).toEqual([ { row: 0, col: 0, textDirection: "stacked" }, { row: 0, col: 1, textDirection: "rotate270" } ])
    })

    it("rejects a ragged grid and spans beyond the table", async () => {
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [ { type: "table", frame: { x: 1, y: 1 },
            data: { rows: [ [ "a", "b" ], [ "c" ] ] } } ] } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [ { type: "table", frame: { x: 1, y: 1 },
            data: { rows: [ [ { text: "a", rowspan: 3 }, "b" ], [ "c" ] ] } } ] } ])).toBe("E_SCHEMA")
    })
})
