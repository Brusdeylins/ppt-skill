/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the HTML-like grid layout of table rows with merged cells.
*/

import { describe, expect, it } from "vitest"
import { layoutTable } from "../../src/core/table-grid.js"

describe("layoutTable", () => {
    it("places plain rows column by column", () => {
        expect(layoutTable([ [ "a", "b" ], [ 1, true ] ])).toEqual({ cols: 2, cells: [
            [ { row: 0, col: 0, colspan: 1, rowspan: 1 }, { row: 0, col: 1, colspan: 1, rowspan: 1 } ],
            [ { row: 1, col: 0, colspan: 1, rowspan: 1 }, { row: 1, col: 1, colspan: 1, rowspan: 1 } ] ] })
    })

    it("skips the columns covered by spans from above and to the left", () => {
        const grid = layoutTable([
            [ { colspan: 2 }, "c" ],
            [ { rowspan: 2 }, "b", "c" ],
            [ { colspan: 2 } ]
        ])
        expect(typeof grid === "string" ? grid : grid.cells.map((row) => row.map((c) => c.col))).toEqual([ [ 0, 2 ], [ 0, 1, 2 ], [ 1 ] ])
    })

    it("reports ragged rows, overflowing spans and overlaps", () => {
        expect(layoutTable([ [ "a", "b" ], [ "c" ] ])).toMatch(/row 1 fills 1 of the 2/)
        expect(layoutTable([ [ "a" ], [ "b", "c" ] ])).toMatch(/row 1 has more cells/)
        expect(layoutTable([ [ { rowspan: 3 }, "b" ], [ "c" ] ])).toMatch(/spans 3 rows/)
        expect(layoutTable([ [ "a", { rowspan: 2 } ], [ { colspan: 2 } ] ])).toMatch(/more cells|overlaps/)
    })
})
