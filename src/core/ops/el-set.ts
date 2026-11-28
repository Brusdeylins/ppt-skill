/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/el-set: plan transformer of `el.set` -- the in-place change of
**  one named element: text, frame, fill (with transparency)/border/font
**  color, alt text and title, the click target (URL or slide link), the
**  image of a picture, the cells and rows of a table, the stacking order
**  (in op order, after the document's new elements are placed). Every
**  field is checked against the element's type. An element planned earlier
**  in the same document has its spec patched; an existing one gets a
**  planned edit the engine applies after the slide's fills. A table's rows
**  are tracked through the document, so cell addresses, merges and row
**  counts are checked at plan time.
*/

import { PptcError } from "../../infra/errors.js"
import type { ElSetOp } from "../../schema/ops.js"
import type { ElementSpec, Paragraph, RichText, TableCell } from "../../schema/payloads.js"
import { seedPlaceholderName, type Frame, type ShapeInfo } from "../model.js"
import { lintElementOverlap, lintFontSize, richTextSizes, richTextToPlain } from "../lint.js"
import { layoutTable } from "../table-grid.js"
import { entryLayout, resolveEntry } from "./fill-common.js"
import { PROMPT_BOX_PREFIX, checkTable, existingShapes, nameMatches, overlapObstacles, plannedFrame,
    rotatedBox, specFrame } from "./elements.js"
import { slideAddrOf, type ElementPatch, type OpHandler, type PlanContext, type PlannedEdit,
    type SlidePlanEntry } from "./registry.js"

/**  an element type as el.set distinguishes it  */
type EditType = ShapeInfo["type"]

/**  every element type  */
const ALL: EditType[] = [ "placeholder", "textbox", "shape", "picture", "table", "chart", "connector", "group", "other" ]

/**  the existing element types each el.set field applies to  */
const EXISTING_TYPES: Record<keyof ElementPatch, EditType[]> = {
    text: [ "placeholder", "textbox", "shape" ],
    frame: ALL,
    fill: [ "placeholder", "textbox", "shape" ],
    transparency: [ "placeholder", "textbox", "shape" ],
    border: [ "placeholder", "textbox", "shape", "picture", "connector" ],
    borderPt: [ "placeholder", "textbox", "shape", "picture", "connector" ],
    fontColor: [ "placeholder", "textbox", "shape" ],
    altText: ALL,
    title: ALL,
    hyperlink: [ "placeholder", "textbox", "shape", "picture" ],
    image: [ "picture" ],
    cells: [ "table" ],
    addRows: [ "table" ],
    removeRows: [ "table" ],
    zOrder: ALL
}

/**  the planned element types each el.set field applies to  */
const PLANNED_TYPES: Record<keyof ElementPatch, ElementSpec["type"][]> = {
    text: [ "textbox", "shape" ],
    frame: [ "textbox", "shape", "table", "chart", "image" ],
    fill: [ "textbox", "shape" ],
    transparency: [ "textbox", "shape" ],
    border: [ "textbox", "shape", "connector" ],
    borderPt: [ "textbox", "shape", "connector" ],
    fontColor: [ "textbox", "shape" ],
    altText: [ "textbox", "shape", "table", "chart", "image", "connector" ],
    title: [ "textbox", "shape", "table", "chart", "image", "connector" ],
    hyperlink: [ "textbox", "shape", "image" ],
    image: [ "image" ],
    cells: [ "table" ],
    addRows: [ "table" ],
    removeRows: [ "table" ],
    zOrder: [ "textbox", "shape", "table", "chart", "image", "connector" ]
}

/**  preset geometries of lines and connectors (el.add draws a connector
     as a line autoshape)  */
const LINE_GEOMETRY = /^(line|straightConnector\d|bentConnector\d|curvedConnector\d)$/

/**  the given fields of an el.set  */
const fieldsOf = (patch: ElementPatch): (keyof ElementPatch)[] =>
    (Object.keys(patch) as (keyof ElementPatch)[]).filter((k) => patch[k] !== undefined)

/**  reject a field the element's type does not take  */
const checkFields = <T extends string>(name: string, type: T, fields: (keyof ElementPatch)[],
    allowed: Record<keyof ElementPatch, T[]>): void => {
    for (const field of fields)
        if (!allowed[field].includes(type))
            throw new PptcError("E_SCHEMA", `element '${name}' is a ${type} -- '${field}' applies to `
                + `${allowed[field].join(", ")} only`)
}

/**  a merged cell of a table: its top-left grid position and spans  */
interface Merge { row: number, col: number, rowspan: number, colspan: number }

/**  an existing table as the document has changed it so far: for every
     current row the row it was before the apply (null: added), the merged
     cells, and the original cells an el.set has overwritten  */
export interface TableModel {
    origin: (number | null)[]
    cols: number
    merges: Merge[]
    overwritten: Set<string>
    /**  whether row 0 is the table's header row  */
    header: boolean
}

/**  the merge a grid cell lies in, null outside every merge  */
const mergeAt = (merges: Merge[], row: number, col: number): Merge | null =>
    merges.find((m) => m.row <= row && row < m.row + m.rowspan && m.col <= col && col < m.col + m.colspan) ?? null

/**  a merge spanning rows across the boundary above `row`  */
const spansInto = (merges: Merge[], row: number): Merge | null =>
    merges.find((m) => m.rowspan > 1 && m.row < row && row < m.row + m.rowspan) ?? null

/**
 *  Apply the table fields of one el.set to the model of a table (E_SCHEMA
 *  when an address is out of range, a row of a vertical merge is removed,
 *  a row is inserted into one, or a cell covered by a merge is set).
 *
 *  @param model - the table model (changed in place)
 *  @param set - the el.set fields
 *  @param name - the element name for messages
 */
const applyTableEdit = (model: TableModel, set: ElementPatch, name: string): void => {
    const fail = (message: string): never => {
        throw new PptcError("E_SCHEMA", `table '${name}': ${message}`)
    }
    if (set.removeRows !== undefined) {
        const rows = [ ...new Set(set.removeRows) ].sort((a, b) => b - a)
        for (const row of rows) {
            if (row >= model.origin.length)
                fail(`row ${row} is out of range (0-${model.origin.length - 1})`)
            const merge = model.merges.find((m) => m.rowspan > 1 && m.row <= row && row < m.row + m.rowspan)
            if (merge !== undefined)
                fail(`row ${row} is part of the merged cell at row ${merge.row}, column ${merge.col}`)
        }
        if (rows.length >= model.origin.length)
            fail("a table keeps at least one row")
        if (rows.includes(0))
            model.header = false
        for (const row of rows) {
            model.origin.splice(row, 1)
            model.merges = model.merges.filter((m) => m.row !== row)
                .map((m) => m.row > row ? { ...m, row: m.row - 1 } : m)
        }
    }
    if (set.addRows !== undefined) {
        const at = set.addRows.at ?? model.origin.length
        if (at > model.origin.length)
            fail(`row position ${at} is out of range (0-${model.origin.length})`)
        if (at === 0 && model.header)
            fail("row 0 is the header row -- insert at 1 or later")
        const merge = spansInto(model.merges, at)
        if (merge !== null)
            fail(`row position ${at} lies inside the merged cell at row ${merge.row}, column ${merge.col}`)
        for (const [ i, row ] of set.addRows.rows.entries())
            if (row.length !== model.cols)
                fail(`added row ${i} has ${row.length} cell(s), the table has ${model.cols} column(s)`)
        const n = set.addRows.rows.length
        model.origin.splice(at, 0, ...new Array<null>(n).fill(null))
        model.merges = model.merges.map((m) => m.row >= at ? { ...m, row: m.row + n } : m)
    }
    for (const cell of set.cells ?? []) {
        if (cell.row >= model.origin.length || cell.col >= model.cols)
            fail(`cell row ${cell.row}, column ${cell.col} is out of range `
                + `(${model.origin.length} row(s), ${model.cols} column(s))`)
        const merge = mergeAt(model.merges, cell.row, cell.col)
        if (merge !== null && (merge.row !== cell.row || merge.col !== cell.col))
            fail(`cell row ${cell.row}, column ${cell.col} is covered by the merged cell at row ${merge.row}, column ${merge.col}`)
        const origin = model.origin[cell.row]
        if (origin !== null && origin !== undefined)
            model.overwritten.add(`${origin},${cell.col}`)
    }
}

/**
 *  The model of an existing table after the el.set edits planned for it
 *  so far.
 *
 *  @param shape - the table as read from the deck
 *  @param edits - the planned edits of the slide
 *  @returns the table model
 */
export const tableModel = (shape: ShapeInfo, edits: PlannedEdit[]): TableModel => {
    const rows = shape.table ?? []
    const model: TableModel = {
        origin: rows.map((_, i) => i),
        cols: shape.colWidths?.length ?? rows[0]?.length ?? 0,
        merges: (shape.merges ?? []).map((m) => ({ ...m })),
        overwritten: new Set(),
        header: shape.headerRow === true
    }
    for (const edit of edits)
        if (edit.name === shape.name)
            applyTableEdit(model, edit.set, shape.name)
    return model
}

/**  a row cell of el.set as a table payload cell  */
const payloadCell = (value: RichText | number | boolean): unknown =>
    typeof value === "object" ? { text: value } : value

/**  patch the table payload of a planned table: rows go and come by grid
     row (header row 0), cells are set by grid position  */
const patchTable = (spec: Extract<ElementSpec, { type: "table" }>, set: ElementPatch, name: string): void => {
    const fail = (message: string): never => {
        throw new PptcError("E_SCHEMA", `table '${name}': ${message}`)
    }
    const headers = spec.data.headers !== undefined
    let rows: unknown[][] = [ ...(headers ? [ spec.data.headers as unknown[] ] : []), ...spec.data.rows ]
    let headerKept = headers
    const grid = (): { cols: number, merges: Merge[], cells: { row: number, col: number }[][] } => {
        const g = layoutTable(rows)
        if (typeof g === "string")
            return fail(g)
        return { cols: g.cols, cells: g.cells, merges: g.cells.flat().filter((c) => c.colspan > 1 || c.rowspan > 1) }
    }
    if (set.removeRows !== undefined) {
        const { merges } = grid()
        const remove = [ ...new Set(set.removeRows) ].sort((a, b) => b - a)
        for (const row of remove) {
            if (row >= rows.length)
                fail(`row ${row} is out of range (0-${rows.length - 1})`)
            const merge = merges.find((m) => m.rowspan > 1 && m.row <= row && row < m.row + m.rowspan)
            if (merge !== undefined)
                fail(`row ${row} is part of the merged cell at row ${merge.row}, column ${merge.col}`)
        }
        for (const row of remove) {
            rows.splice(row, 1)
            if (row === 0)
                headerKept = false
        }
    }
    if (set.addRows !== undefined) {
        const { cols, merges } = grid()
        const at = set.addRows.at ?? rows.length
        if (at > rows.length)
            fail(`row position ${at} is out of range (0-${rows.length})`)
        if (at === 0 && headerKept)
            fail("row 0 is the header row -- insert at 1 or later")
        const merge = spansInto(merges, at)
        if (merge !== null)
            fail(`row position ${at} lies inside the merged cell at row ${merge.row}, column ${merge.col}`)
        for (const [ i, row ] of set.addRows.rows.entries())
            if (row.length !== cols)
                fail(`added row ${i} has ${row.length} cell(s), the table has ${cols} column(s)`)
        rows.splice(at, 0, ...set.addRows.rows.map((row) => row.map(payloadCell)))
    }
    if (set.cells !== undefined) {
        const { cols, cells, merges } = grid()
        rows = rows.map((row) => [ ...row ])
        for (const cell of set.cells) {
            if (cell.row >= rows.length || cell.col >= cols)
                fail(`cell row ${cell.row}, column ${cell.col} is out of range (${rows.length} row(s), ${cols} column(s))`)
            const k = (cells[cell.row] ?? []).findIndex((c) => c.col === cell.col)
            if (k < 0) {
                const merge = mergeAt(merges, cell.row, cell.col) as Merge
                fail(`cell row ${cell.row}, column ${cell.col} is covered by the merged cell at row ${merge.row}, column ${merge.col}`)
            }
            const row = rows[cell.row] as unknown[]
            const old = row[k]
            row[k] = typeof old === "object" && old !== null ? { ...old as TableCell, text: cell.text }
                : typeof cell.text === "string" ? cell.text : { text: cell.text }
        }
    }
    const body = headerKept ? rows.slice(1) : rows
    if (body.length === 0)
        fail("a table keeps at least one body row")
    spec.data = { ...spec.data, ...(headerKept ? { headers: rows[0] as never } : {}), rows: body as never }
    if (!headerKept)
        delete spec.data.headers
    checkTable(spec, name)
}

/**  rich text with every run in one color  */
const colored = (text: RichText, color: string): RichText =>
    typeof text === "string"
        ? text.split("\n").map((line) => ({ text: line, color }))
        : text.map((p): Paragraph => p.runs === undefined ? { ...p, color } : { ...p, runs: p.runs.map((r) => ({ ...r, color })) })

/**  patch the spec of an element planned earlier in the same document  */
const patchSpec = (spec: ElementSpec, set: ElementPatch, name: string): void => {
    checkFields(name, spec.type, fieldsOf(set), PLANNED_TYPES)
    if (set.frame?.w === 0 || set.frame?.h === 0)
        throw new PptcError("E_SCHEMA", `element '${name}' is a ${spec.type} -- a zero 'w' or 'h' `
            + "applies to existing lines and connectors only")
    if (spec.type !== "connector" && set.frame !== undefined)
        for (const [ key, value ] of Object.entries(set.frame))
            if (value !== undefined)
                (spec.frame as Record<string, number>)[key] = value
    if (set.altText !== undefined)
        spec.altText = set.altText
    if (set.title !== undefined)
        spec.title = set.title
    if (set.hyperlink !== undefined && (spec.type === "textbox" || spec.type === "shape" || spec.type === "image")) {
        if (set.hyperlink === "")
            delete spec.hyperlink
        else
            spec.hyperlink = set.hyperlink
    }
    switch (spec.type) {
        case "textbox":
        case "shape":
            if (set.text !== undefined)
                spec.text = set.text
            if (set.fill !== undefined)
                spec.fill = set.fill
            if (set.transparency !== undefined) {
                if (spec.fill === undefined)
                    throw new PptcError("E_SCHEMA", `element '${name}' has no fill -- give 'fill' with 'transparency'`)
                spec.transparency = set.transparency
            }
            if (set.border !== undefined)
                spec.border = set.border
            if (set.borderPt !== undefined)
                spec.borderPt = set.borderPt
            if (set.fontColor !== undefined && spec.type === "shape")
                spec.fontColor = set.fontColor
            else if (set.fontColor !== undefined && spec.type === "textbox")
                spec.text = colored(spec.text, set.fontColor)
            break
        case "connector":
            if (set.border !== undefined)
                spec.color = set.border
            if (set.borderPt !== undefined)
                spec.widthPt = set.borderPt
            break
        case "image":
            if (set.image !== undefined)
                spec.path = set.image
            break
        case "table":
            if (set.cells !== undefined || set.addRows !== undefined || set.removeRows !== undefined)
                patchTable(spec, set, name)
            break
        default:
            break
    }
}

/**  the frame patch of an existing element, complete (x, y, w, h) when
     the element inherits its frame -- from its layout placeholder, which
     the engine then writes into the slide. An element with a transform
     of its own (even a zero-height line) or one an earlier el.set gave
     takes the patch as it is.  */
const completeFrame = (ctx: PlanContext, entry: SlidePlanEntry, shape: ShapeInfo,
    patch: NonNullable<ElementPatch["frame"]>): NonNullable<ElementPatch["frame"]> => {
    if (shape.frame !== null || entry.edits.some((e) => e.name === shape.name && e.set.frame !== undefined))
        return patch
    const inherited = shape.placeholderIdx === null ? null
        : entryLayout(ctx, entry)?.placeholders.find((p) => p.idx === shape.placeholderIdx)?.frame ?? null
    const base: Partial<Frame> = inherited ?? {}
    const full = { x: patch.x ?? base.x, y: patch.y ?? base.y, w: patch.w ?? base.w, h: patch.h ?? base.h }
    if (full.x === undefined || full.y === undefined || full.w === undefined || full.h === undefined)
        throw new PptcError("E_SCHEMA", `element '${shape.name}' inherits its frame and the layout gives none -- `
            + "set x, y, w and h")
    return { ...patch, x: full.x, y: full.y, w: full.w, h: full.h }
}

/**  the shapes el.set can address on a slide: those of a kept slide, the
     placeholders of a slide added in this document (`PptcPh-<idx>`)  */
const editableShapes = (ctx: PlanContext, entry: SlidePlanEntry): ShapeInfo[] =>
    entry.source.kind === "self" ? existingShapes(ctx, entry)
        : (entryLayout(ctx, entry)?.placeholders ?? []).map((ph) => ({
            name: seedPlaceholderName(ph.idx), id: 0, type: "placeholder", placeholderIdx: ph.idx,
            placeholderKind: ph.kind, frame: null, text: null
        }))

/**  op handler: el.set  */
export const elSet: OpHandler<ElSetOp> = {
    name: "el.set",
    plan(ctx, op): void {
        const entry = resolveEntry(ctx, op.slide)
        const set = Object.fromEntries(Object.entries(op)
            .filter(([ key ]) => ![ "op", "slide", "name" ].includes(key))) as ElementPatch
        const slideAddr = slideAddrOf(ctx, entry)
        const prompt = op.name.startsWith(PROMPT_BOX_PREFIX)

        /*  the new text is rendered verbatim, so it gets the same
            readability lint as el.add (prompt boxes stay exempt)  */
        if (!prompt && op.text !== undefined) {
            const tooSmall = lintFontSize({ element: op.name }, richTextSizes(op.text), ctx.minFontPt, slideAddr)
            if (tooSmall !== null)
                ctx.plan.warnings.push(tooSmall)
        }

        /*  an element generated earlier in the same run: patch its spec  */
        const planned = entry.elements.find((e) => e.name !== null && nameMatches(e.name, op.name))
        if (planned !== undefined) {
            patchSpec(planned.spec, set, op.name)
            if (set.zOrder !== undefined)
                entry.zOrder.push({ target: planned, zOrder: set.zOrder, add: false })
            const frame = specFrame(planned.spec)
            if (set.frame !== undefined && !prompt && frame !== null && set.zOrder !== "back") {
                const overlap = lintElementOverlap(op.name, frame, overlapObstacles(ctx, entry, planned.name), slideAddr)
                if (overlap !== null)
                    ctx.plan.warnings.push(overlap)
            }
            return
        }

        /*  otherwise the element must already exist on the slide  */
        const shapes = editableShapes(ctx, entry).filter((sh) => !entry.removeNames.includes(sh.name))
        const matches = shapes.filter((sh) => nameMatches(sh.name, op.name))
        if (matches.length === 0)
            throw new PptcError("E_ADDR_NOTFOUND",
                `slide has no element named '${op.name}'`,
                { available: shapes.map((sh) => sh.name) })
        if (matches.length > 1)
            throw new PptcError("E_ADDR_AMBIGUOUS",
                `'${op.name}' matches ${matches.length} elements`,
                { candidates: matches.map((sh) => sh.name) })
        const target = matches[0] as ShapeInfo
        /*  a slide number or date an earlier op turned off is gone  */
        if ((target.placeholderKind === "slideNumber" && entry.slideNumber === false)
            || (target.placeholderKind === "date" && entry.date === false))
            throw new PptcError("E_ADDR_NOTFOUND", `element '${target.name}' was removed by an earlier op `
                + `(slide.fill ${target.placeholderKind}: false)`)

        /*  a picture placeholder filled in this document is a picture  */
        const imageFill = entry.fills.find((f) => f.image !== undefined && f.phName === target.name)
        const type: EditType = imageFill !== undefined ? "picture"
            : target.type === "shape" && LINE_GEOMETRY.test(target.shape ?? "") ? "connector" : target.type
        if (set.image !== undefined && type === "placeholder" && target.placeholderKind === "picture")
            throw new PptcError("E_SCHEMA",
                `element '${target.name}' is an empty picture placeholder -- fill it with slide.fill 'image'`)
        checkFields(target.name, type, fieldsOf(set), EXISTING_TYPES)
        if (set.frame?.rotation !== undefined && [ "table", "chart", "other" ].includes(type))
            throw new PptcError("E_SCHEMA", `element '${target.name}' is a ${type} -- it cannot be rotated`)
        /*  only a line or connector may be flat (a horizontal or vertical one)  */
        if ((set.frame?.w === 0 || set.frame?.h === 0) && type !== "connector")
            throw new PptcError("E_SCHEMA", `element '${target.name}' is a ${type} -- a zero 'w' or 'h' `
                + "applies to lines and connectors only")

        /*  the image of a placeholder filled in this document is the fill's  */
        if (imageFill !== undefined && set.image !== undefined) {
            imageFill.image = set.image
            delete set.image
        }
        if (set.frame !== undefined) {
            set.frame = completeFrame(ctx, entry, target, set.frame)
            const before = plannedFrame(entry, target.name, target.frame)
            const f = { ...before, ...Object.fromEntries(Object.entries(set.frame).filter(([ , v ]) => v !== undefined)) }
            const frame = rotatedBox(f as Frame & { rotation?: number })
            if (!prompt && frame !== null && set.zOrder !== "back") {
                const overlap = lintElementOverlap(target.name, frame, overlapObstacles(ctx, entry, target.name), slideAddr)
                if (overlap !== null)
                    ctx.plan.warnings.push(overlap)
            }
        }
        /*  a transparency applies to a solid fill: the new one, one an
            earlier el.set gave, or the element's own  */
        if (set.transparency !== undefined && set.fill === undefined && target.fill === undefined
            && !entry.edits.some((e) => e.name === target.name && e.set.fill !== undefined))
            throw new PptcError("E_SCHEMA", `element '${target.name}' has no solid fill of its own -- `
                + "give 'fill' with 'transparency'")
        if (type === "table")
            applyTableEdit(tableModel(target, entry.edits), set, target.name)
        if (set.text !== undefined && (target.placeholderKind === "title" || target.placeholderIdx === 0))
            entry.title = richTextToPlain(set.text).join(" ")
        /*  the stacking order changes in op order, after the elements
            the document adds are in place  */
        if (set.zOrder !== undefined) {
            entry.zOrder.push({ target: target.name, zOrder: set.zOrder, add: false })
            delete set.zOrder
        }
        if (fieldsOf(set).length > 0)
            entry.edits.push({ name: target.name, set })
    }
}
