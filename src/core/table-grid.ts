/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/table-grid: the grid layout of a table payload with merged cells.
**  Cells covered by a `colspan`/`rowspan` are omitted from the following
**  cells of their row and from the rows below (as in HTML); the layout
**  assigns every given cell its grid position and checks that all rows
**  fill the same number of columns.
*/

/**  the span fields of a table cell (plain values span nothing)  */
export interface SpanCell {
    colspan?: number | undefined
    rowspan?: number | undefined
}

/**  the grid position of one given cell  */
export interface GridCell {
    /**  zero-based grid row  */
    row: number
    /**  zero-based grid column of the cell's first column  */
    col: number
    /**  number of columns the cell spans  */
    colspan: number
    /**  number of rows the cell spans  */
    rowspan: number
}

/**  the laid-out grid: its column count and the position of every given
     cell, in the shape of the input rows  */
export interface TableGrid {
    cols: number
    cells: GridCell[][]
}

/**  the spans of a cell value, 1 for plain values  */
const spansOf = (cell: unknown): { colspan: number, rowspan: number } => {
    const c = typeof cell === "object" && cell !== null ? cell as SpanCell : {}
    return { colspan: c.colspan ?? 1, rowspan: c.rowspan ?? 1 }
}

/**
 *  Lay out table rows on their grid, HTML-like.
 *
 *  @param rows - the table rows (header row first, if any); a cell is a
 *                plain value or an object with optional spans
 *  @returns the grid, or the reason why the rows do not form one
 */
export const layoutTable = (rows: unknown[][]): TableGrid | string => {
    const cols = (rows[0] ?? []).reduce<number>((sum, cell) => sum + spansOf(cell).colspan, 0)
    const taken: boolean[][] = rows.map(() => new Array<boolean>(cols).fill(false))
    const cells: GridCell[][] = []
    for (const [ r, row ] of rows.entries()) {
        const placed: GridCell[] = []
        let col = 0
        for (const cell of row) {
            while (col < cols && taken[r]?.[col] === true)
                col++
            const { colspan, rowspan } = spansOf(cell)
            if (col + colspan > cols)
                return `row ${r} has more cells than the ${cols} column(s) of the first row`
            if (r + rowspan > rows.length)
                return `the cell at row ${r}, column ${col} spans ${rowspan} rows, but the table ends at row ${rows.length - 1}`
            for (let dr = 0; dr < rowspan; dr++) {
                const line = taken[r + dr] as boolean[]
                for (let dc = 0; dc < colspan; dc++) {
                    if (line[col + dc] === true)
                        return `the cell at row ${r}, column ${col} overlaps a merged cell at row ${r + dr}, column ${col + dc}`
                    line[col + dc] = true
                }
            }
            placed.push({ row: r, col, colspan, rowspan })
            col += colspan
        }
        const filled = (taken[r] as boolean[]).filter((t) => t).length
        if (filled < cols)
            return `row ${r} fills ${filled} of the ${cols} column(s) of the first row`
        cells.push(placed)
    }
    return { cols, cells }
}
