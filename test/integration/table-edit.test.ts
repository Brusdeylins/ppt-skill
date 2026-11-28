/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: el.set on a table in place -- `cells` sets texts by grid
**  position (header row 0), `addRows` inserts rows in the format of the
**  neighbouring body row, `removeRows` takes rows out; the table style and the cell
**  formatting stay, the frame's height follows the rows, merged cells are
**  respected (a covered cell, a row of a vertical merge or a position
**  inside one fail at plan time), no row goes above a header row (as on a
**  planned table), and state reads the result.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import type { Element } from "@xmldom/xmldom"
import { elements, firstElement } from "../../src/engine/xml.js"
import { apply, failure, freshDeck, shapeNamed, slideDoc, stateOf } from "../util/elements.js"

let deck: string

/**  the table shape on the first slide as state reads it  */
const table = async (): Promise<{ table?: string[][], frame: { h: number } | null, merges?: unknown[] }> =>
    (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Tab-")) as never

beforeAll(async () => {
    deck = await freshDeck("table-edit.pptx")
    await apply(deck, [
        { op: "slide.add", layout: "Title Only", ref: "s", placeholders: { title: { text: "Tabelle" } } },
        { op: "el.add", slide: "$s", elements: [ { type: "table", name: "Tab", frame: { x: 0.5, y: 1.5, w: 8 },
            data: { headers: [ "Name", "Q1", "Q2" ], style: { rowHeight: 0.4 },
                rows: [ [ { text: "Nord", rowspan: 2, fill: "DDEEFF" }, { text: "1", bold: true }, "2" ],
                    [ "3", "4" ],
                    [ { text: "Summe", colspan: 2, fill: "FFEEDD" }, "9" ] ] } } ] }
    ])
})

describe("el.set table cells and rows", () => {
    it("sets cells by grid position, keeping their format and the table style", async () => {
        const style = firstElement(shapeNamed(await slideDoc(deck, 0), "Tab"), "a:tableStyleId")?.textContent ?? null
        await apply(deck, [ { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 1, col: 1, text: "10" }, { row: 3, col: 0, text: "Gesamt" } ] } ])
        expect((await table()).table).toEqual([ [ "Name", "Q1", "Q2" ], [ "Nord", "10", "2" ], [ "", "3", "4" ], [ "Gesamt", "", "9" ] ])
        const tab = shapeNamed(await slideDoc(deck, 0), "Tab")
        expect(firstElement(tab, "a:tableStyleId")?.textContent ?? null).toBe(style)
        const cell = elements(elements(tab, "a:tr")[1] as never, "a:tc")[1] as never
        expect(firstElement(cell, "a:rPr")?.getAttribute("b")).toBe("1")
    })

    it("adds rows in the format of their neighbour; the frame takes the rows' height", async () => {
        const before = await table()
        await apply(deck, [ { op: "el.set", slide: "0", name: "Tab", addRows: { at: 4, rows: [ [ "Süd", 5, true ] ] } } ])
        const after = await table()
        expect(after.table?.[4]).toEqual([ "Süd", "5", "true" ])
        expect(after.frame?.h).toBeCloseTo(5 * 0.4, 1)
        const row = elements(shapeNamed(await slideDoc(deck, 0), "Tab"), "a:tr")[4] as never
        expect(elements(row, "a:tc").map((tc) => tc.getAttribute("gridSpan") ?? tc.getAttribute("hMerge"))).toEqual([ null, null, null ])
        const tcPr = firstElement(row, "a:tcPr") as never
        const fill = elements(tcPr, "a:solidFill").find((f) => f.parentNode === tcPr) as never
        expect(firstElement(fill, "a:srgbClr")?.getAttribute("val")).toBe("FFEEDD")
        expect(after.merges).toEqual(before.merges)
    })

    it("adds rows at 1 in the format of the first body row, not of the header row", async () => {
        const own = await freshDeck("table-edit-header.pptx")
        await apply(own, [
            { op: "slide.add", layout: "Title Only", ref: "s", placeholders: { title: { text: "Kopf" } } },
            { op: "el.add", slide: "$s", elements: [ { type: "table", name: "Tab", frame: { x: 0.5, y: 1.5, w: 6 },
                data: { headers: [ "A", "B" ], style: { headerBg: "1F4E79", headerFg: "FFFFFF", headerHeight: 0.6, rowHeight: 0.3 },
                    rows: [ [ "1", "2" ], [ "3", "4" ] ] } } ] }
        ])

        /*  a row's format: its height, cell and run properties  */
        const format = (row: Element): string => [ row.getAttribute("h"),
            ...[ ...elements(row, "a:tcPr"), ...elements(row, "a:rPr") ].map((el) => el.toString()) ].join("|")
        const rowsOf = async (): Promise<Element[]> => elements(shapeNamed(await slideDoc(own, 0), "Tab"), "a:tr")
        const [ header, body ] = (await rowsOf()).map(format)
        expect(body).not.toBe(header)
        await apply(own, [ { op: "el.set", slide: "0", name: "Tab", addRows: { at: 1, rows: [ [ "x", "y" ] ] } } ])
        const rows = await rowsOf()
        expect(rows.map((r) => elements(r, "a:t").map((t) => t.textContent).join(","))).toEqual([ "A,B", "x,y", "1,2", "3,4" ])
        expect(format(rows[1] as Element)).toBe(body)
    })

    it("rejects rows above the header row of an existing table, as of a planned one", async () => {
        const own = await freshDeck("table-edit-at0.pptx")
        await apply(own, [
            { op: "slide.add", layout: "Title Only", ref: "s", placeholders: { title: { text: "Kopf" } } },
            { op: "el.add", slide: "$s", elements: [
                { type: "table", name: "Head", frame: { x: 0.5, y: 1.5, w: 6 }, data: { headers: [ "A", "B" ], rows: [ [ "1", "2" ] ] } },
                { type: "table", name: "Body", frame: { x: 0.5, y: 4, w: 6 }, data: { rows: [ [ "1", "2" ] ] } } ] }
        ])
        const shapes = (await stateOf(own)).slides[0]?.shapes ?? []
        expect(shapes.find((sh) => sh.name.startsWith("Head-"))?.headerRow).toBe(true)
        expect(shapes.find((sh) => sh.name.startsWith("Body-"))?.headerRow).toBeUndefined()
        const at0 = { addRows: { at: 0, rows: [ [ "x", "y" ] ] } }
        const before = readFileSync(own)
        expect(await failure(own, [ { op: "el.set", slide: "0", name: "Head", ...at0 } ])).toBe("E_SCHEMA")
        expect(await failure(own, [ { op: "el.add", slide: "0", elements: [ { type: "table", name: "Plan", frame: { x: 7, y: 1.5, w: 3 },
            data: { headers: [ "A", "B" ], rows: [ [ "1", "2" ] ] } } ] }, { op: "el.set", slide: "0", name: "Plan", ...at0 } ])).toBe("E_SCHEMA")
        expect(readFileSync(own).equals(before)).toBe(true)

        /*  a table without a header row, or one whose header row went, takes rows at 0  */
        await apply(own, [ { op: "el.set", slide: "0", name: "Body", ...at0 },
            { op: "el.set", slide: "0", name: "Head", removeRows: [ 0 ] }, { op: "el.set", slide: "0", name: "Head", ...at0 } ])
        const after = (await stateOf(own)).slides[0]?.shapes ?? []
        expect(after.find((sh) => sh.name.startsWith("Body-"))?.table).toEqual([ [ "x", "y" ], [ "1", "2" ] ])
        const head = after.find((sh) => sh.name.startsWith("Head-"))
        expect([ head?.table, head?.headerRow ]).toEqual([ [ [ "x", "y" ], [ "1", "2" ] ], undefined ])
    })

    it("removes rows and shrinks the frame; the sequence of a document sees every step", async () => {
        const before = await table()
        await apply(deck, [
            { op: "el.set", slide: "0", name: "Tab", removeRows: [ 4 ], addRows: { rows: [ [ "a", "b", "c" ], [ "d", "e", "f" ] ] } },
            { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 5, col: 2, text: "F" } ] },
            { op: "el.set", slide: "0", name: "Tab", removeRows: [ 4 ] }
        ])
        const after = await table()
        expect(after.table?.slice(4)).toEqual([ [ "d", "e", "F" ] ])
        expect(after.frame?.h).toBeCloseTo(before.frame?.h ?? 0, 1)
    })

    it("rejects covered cells, merged rows, positions inside merges and wrong addresses", async () => {
        const before = readFileSync(deck)
        const cases: Record<string, unknown>[] = [
            { cells: [ { row: 2, col: 0, text: "x" } ] },
            { cells: [ { row: 3, col: 1, text: "x" } ] },
            { cells: [ { row: 9, col: 0, text: "x" } ] },
            { removeRows: [ 1 ] },
            { removeRows: [ 2 ] },
            { removeRows: [ 0, 1, 2, 3, 4 ] },
            { addRows: { at: 2, rows: [ [ "a", "b", "c" ] ] } },
            { addRows: { rows: [ [ "a", "b" ] ] } },
            { addRows: { at: 9, rows: [ [ "a", "b", "c" ] ] } }
        ]
        for (const set of cases)
            expect(await failure(deck, [ { op: "el.set", slide: "0", name: "Tab", ...set } ]), JSON.stringify(set)).toBe("E_SCHEMA")
        expect(readFileSync(deck).equals(before)).toBe(true)
    })

    it("checks addresses against the rows of earlier ops", async () => {
        expect(await failure(deck, [
            { op: "el.set", slide: "0", name: "Tab", removeRows: [ 4 ] },
            { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 4, col: 0, text: "x" } ] } ])).toBe("E_SCHEMA")
    })

    it("edits a table planned in the same document", async () => {
        await apply(deck, [
            { op: "el.add", slide: "0", elements: [ { type: "table", name: "Neu", frame: { x: 9, y: 1.5, w: 3 },
                data: { headers: [ "A", "B" ], rows: [ [ { text: "x", colspan: 2 } ], [ "1", "2" ] ] } } ] },
            { op: "el.set", slide: "0", name: "Neu", cells: [ { row: 2, col: 1, text: "zwei" } ],
                addRows: { at: 3, rows: [ [ "3", [ { runs: [ { text: "vier", bold: true } ] } ] ] ] } },
            { op: "el.set", slide: "0", name: "Neu", removeRows: [ 1 ] }
        ])
        const neu = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Neu-"))
        expect(neu?.table).toEqual([ [ "A", "B" ], [ "1", "zwei" ], [ "3", "vier" ] ])
        expect(await failure(deck, [
            { op: "el.add", slide: "0", elements: [ { type: "table", name: "N2", frame: { x: 9, y: 4, w: 3 },
                data: { headers: [ "A", "B" ], rows: [ [ { text: "x", colspan: 2 } ] ] } } ] },
            { op: "el.set", slide: "0", name: "N2", cells: [ { row: 1, col: 1, text: "covered" } ] } ])).toBe("E_SCHEMA")
    })
})
