/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/replace: plan transformer of `text.replace`. The op sees the
**  deck after every earlier op of the document: the content planned so far
**  (placeholder fills, el.set texts and table cells, el.add elements,
**  footer, notes) is replaced right here and counted per slide; the
**  content the slide carries before the apply is replaced by the engine,
**  which leaves out what an earlier op replaced or removed (the planned
**  replace lists it), and counts there.
*/

import type { TextReplaceOp } from "../../schema/ops.js"
import type { ElementSpec, RichText } from "../../schema/payloads.js"
import { replacePlain, replaceRichText, type ReplaceSpec } from "../text-replace.js"
import { resolveEntry } from "./fill-common.js"
import { existingShapes } from "./elements.js"
import { tableModel } from "./el-set.js"
import type { OpHandler, PlanContext, PlannedReplace, SlidePlanEntry } from "./registry.js"

/**  a table cell value with its text replaced  */
const replaceCell = (cell: unknown, rich: (t: RichText) => RichText): unknown => {
    if (typeof cell === "object" && cell !== null)
        return { ...cell, text: rich((cell as { text: RichText }).text) }
    const text = String(cell)
    const done = rich(text)
    return done === text ? cell : done
}

/**  replace in the text of a planned element (textbox, shape, table)  */
const replaceSpec = (spec: ElementSpec, rich: (t: RichText) => RichText): void => {
    if (spec.type === "textbox" || (spec.type === "shape" && spec.text !== undefined))
        spec.text = rich(spec.text as RichText)
    else if (spec.type === "table")
        spec.data = {
            ...spec.data,
            ...(spec.data.headers !== undefined && { headers: spec.data.headers.map((c) => replaceCell(c, rich)) as never }),
            rows: spec.data.rows.map((row) => row.map((c) => replaceCell(c, rich))) as never
        }
}

/**  what the engine must leave out of the slide's existing content: text
     an earlier op of the document replaced or removed  */
const skipped = (ctx: PlanContext, entry: SlidePlanEntry): Pick<PlannedReplace, "skip" | "skipCells"> => {
    const skip = [
        ...entry.removeNames,
        ...entry.fills.filter((f) => (f.text !== undefined && !f.append) || f.image !== undefined).map((f) => f.phName),
        ...entry.edits.filter((e) => e.set.text !== undefined).map((e) => e.name),
        ...entry.footer !== null ? [ "@ftr" ] : [],
        ...entry.slideNumber !== null ? [ "@sldNum" ] : [],
        ...entry.date !== null ? [ "@dt" ] : []
    ]
    const skipCells: PlannedReplace["skipCells"] = []
    const tables = new Set(entry.edits.filter((e) => e.set.cells !== undefined || e.set.addRows !== undefined
        || e.set.removeRows !== undefined).map((e) => e.name))
    for (const shape of existingShapes(ctx, entry).filter((sh) => tables.has(sh.name))) {
        const model = tableModel(shape, entry.edits)
        for (const cell of model.overwritten) {
            const [ row, col ] = cell.split(",").map(Number) as [ number, number ]
            skipCells.push({ name: shape.name, row, col })
        }
        for (let row = 0; row < (shape.table ?? []).length; row++)
            if (!model.origin.includes(row))
                skipCells.push({ name: shape.name, row, col: -1 })
    }
    return { skip: [ ...new Set(skip) ], skipCells }
}

/**  op handler: text.replace  */
export const textReplace: OpHandler<TextReplaceOp> = {
    name: "text.replace",
    plan(ctx, op): void {
        const spec: ReplaceSpec = { find: op.find, replace: op.replace,
            matchCase: op.matchCase ?? false, wholeWord: op.wholeWord ?? false }
        const entries = op.slide === undefined ? [ ...ctx.plan.entries ] : [ resolveEntry(ctx, op.slide) ]
        for (const entry of entries) {
            let count = 0
            const rich = (text: RichText): RichText => {
                const done = replaceRichText(text, spec)
                count += done.count
                return done.text
            }
            const plain = (text: string): string => {
                const done = replacePlain(text, spec)
                count += done.count
                return done.text
            }

            /*  the content planned so far  */
            for (const fill of entry.fills)
                if (fill.text !== undefined)
                    fill.text = rich(fill.text)
            for (const { set } of entry.edits) {
                if (set.text !== undefined)
                    set.text = rich(set.text)
                if (set.cells !== undefined)
                    set.cells = set.cells.map((cell) => ({ ...cell, text: rich(cell.text) }))
                if (set.addRows !== undefined)
                    set.addRows = { ...set.addRows, rows: set.addRows.rows.map((row) =>
                        row.map((cell) => replaceCell(cell, rich) as RichText | number | boolean)) }
            }
            for (const planned of entry.elements)
                replaceSpec(planned.spec, rich)
            if (entry.footer !== null)
                entry.footer = plain(entry.footer)
            if (op.notes === true && entry.notes !== null)
                entry.notes = plain(entry.notes)
            if (count > 0)
                entry.replaced.push({ op: ctx.opIndex, count })

            /*  the title as later selectors of the document see it  */
            if (entry.title !== null)
                entry.title = replacePlain(entry.title, spec).text

            /*  the content the slide carries before the apply  */
            entry.replaces.push({ op: ctx.opIndex, ...spec, notes: op.notes === true && entry.notes === null,
                ...skipped(ctx, entry) })
        }
    }
}
