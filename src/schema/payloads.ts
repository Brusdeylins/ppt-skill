/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  schema/payloads: Zod schemas for the JSON payload building blocks shared
**  by several ops (rich text, tables, charts, free elements). These schemas
**  are the single source of truth: they validate at runtime, generate the
**  TypeScript types via z.infer, and feed the `pptc schema` command.
*/

import { z } from "zod"

/**  theme color slots a color field may name instead of a hex value  */
export const SCHEME_COLORS = [
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "tx1", "tx2", "bg1", "bg2", "dk1", "lt1", "dk2", "lt2"
] as const

/**  a theme color slot name  */
export type SchemeColor = typeof SCHEME_COLORS[number]

/**
 *  Whether a (validated) color names a theme slot rather than a hex value.
 *
 *  @param color - a color as validated by ColorSchema
 *  @returns true for a theme color slot
 */
export const isSchemeColor = (color: string): color is SchemeColor =>
    (SCHEME_COLORS as readonly string[]).includes(color)

/**  6-digit hex color, with or without leading `#`, normalized to RRGGBB --
     or a bare theme color slot (`accent1`, `tx1`, `dk2`, ...), written as
     a theme reference that follows the template's color scheme  */
export const ColorSchema = z
    .string()
    .regex(/^(#?[0-9a-fA-F]{6}|accent[1-6]|tx[12]|bg[12]|dk[12]|lt[12])$/,
        "expected a 6-digit hex color like '#1F4E79' or a theme color "
        + "(accent1-accent6, tx1, tx2, bg1, bg2, dk1, lt1, dk2, lt2)")
    .transform((s) => isSchemeColor(s) ? s : s.replace("#", "").toUpperCase())

/**  Geometry of a box in inches; `x`/`y` are required, size may be omitted.  */
export const FrameSchema = z.object({
    x: z.number().min(0),
    y: z.number().min(0),
    w: z.number().positive().optional(),
    h: z.number().positive().optional(),
    rotation: z.number().min(-360).max(360).optional()
}).strict()

/**  Text language as a BCP 47 tag ("de-DE", "en-US") -- PowerPoint's
     proofing and hyphenation language of the text.  */
export const LangSchema = z
    .string()
    .regex(/^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$/, "expected a language tag like 'de-DE' or 'en-US'")

/**  Direction of the text in a box or table cell: `rotate90` runs top to
     bottom (`vert`), `rotate270` bottom to top (`vert270`), `stacked`
     stacks the letters (`wordArtVert`).  */
export const TextDirectionSchema = z.enum(["horizontal", "rotate90", "rotate270", "stacked"])

/**  Vertical anchoring of text in its box or cell.  */
export const ValignSchema = z.enum(["top", "middle", "bottom"])

/**  Inner text margin in inches: one value for all sides or `[l, t, r, b]`.  */
export const MarginSchema = z.union([
    z.number().min(0),
    z.tuple([z.number().min(0), z.number().min(0), z.number().min(0), z.number().min(0)])
])

/**  Hyperlink target of a run or picture (a URL).  */
export const HyperlinkSchema = z.string().url()

/**  Character-level formatting shared by paragraphs and runs.  */
const fontProps = {
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    strike: z.boolean().optional(),
    superscript: z.boolean().optional(),
    subscript: z.boolean().optional(),
    size: z.number().positive().optional(),
    font: z.string().min(1).optional(),
    color: ColorSchema.optional(),
    /**  text highlight (marker) color  */
    highlight: ColorSchema.optional(),
    /**  text language; a paragraph's applies to its runs  */
    lang: LangSchema.optional()
}

/**  superscript and subscript exclude each other  */
const oneBaseline = (f: { superscript?: boolean | undefined, subscript?: boolean | undefined }): boolean =>
    f.superscript !== true || f.subscript !== true

/**  One formatted text run inside a paragraph.  */
export const RunSchema = z.object({
    text: z.string(),
    hyperlink: HyperlinkSchema.optional(),
    ...fontProps
}).strict().refine(oneBaseline, "a run is either 'superscript' or 'subscript' (not both)")

/**  One paragraph of rich text: plain `text` or formatted `runs`.  */
export const ParagraphSchema = z.object({
    text: z.string().optional(),
    runs: z.array(RunSchema).optional(),
    level: z.number().int().min(0).max(8).optional(),
    alignment: z.enum(["left", "center", "right", "justify"]).optional(),
    /**  `true` a bullet (placeholders: their own bullet style), `"number"`
         a numbered list (1. 2. 3.), `false` none  */
    bullet: z.union([z.boolean(), z.literal("number")]).optional(),
    /**  space before / after the paragraph in pt  */
    spaceBefore: z.number().min(0).optional(),
    spaceAfter: z.number().min(0).optional(),
    /**  line spacing as a multiple of single spacing (1.5 = 150%)  */
    lineSpacing: z.number().positive().max(10).optional(),
    ...fontProps
}).strict().refine((p) => (p.text !== undefined) !== (p.runs !== undefined),
    "a paragraph needs either 'text' or 'runs' (not both)")
    .refine(oneBaseline, "a paragraph is either 'superscript' or 'subscript' (not both)")

/**  Rich text content: a single plain string or a list of paragraphs.  */
export const RichTextSchema = z.union([z.string(), z.array(ParagraphSchema).min(1)])

/**  Horizontal alignment of a table column or cell.  */
export const CellAlignSchema = z.enum(["left", "center", "right"])

/**  A styled table cell; `colspan`/`rowspan` merge it with its neighbours
     to the right / below, whose cells are then omitted from their rows
     (as in HTML).  */
export const TableCellSchema = z.object({
    text: RichTextSchema,
    bold: z.boolean().optional(),
    fill: ColorSchema.optional(),
    color: ColorSchema.optional(),
    align: CellAlignSchema.optional(),
    valign: ValignSchema.optional(),
    colspan: z.number().int().min(1).optional(),
    rowspan: z.number().int().min(1).optional(),
    textDirection: TextDirectionSchema.optional()
}).strict()

/**  One table cell: a plain value or a styled cell.  */
export const CellSchema = z.union([z.string(), z.number(), z.boolean(), TableCellSchema])

/**  Table payload: headers, rows and optional styling.  */
export const TableSchema = z.object({
    headers: z.array(z.union([z.string(), TableCellSchema])).optional(),
    rows: z.array(z.array(CellSchema)).min(1),
    style: z.object({
        headerBg: ColorSchema.optional(),
        headerFg: ColorSchema.optional(),
        altRowBg: ColorSchema.optional(),
        fontSize: z.number().positive().optional(),
        border: ColorSchema.optional(),
        /**  fixed header-row height in inches (uniform headers across tables)  */
        headerHeight: z.number().positive().optional(),
        /**  fixed body-row height in inches (uniform rows)  */
        rowHeight: z.number().positive().optional()
    }).strict().optional()
}).strict()

/**  Chart types supported by the chart element (PptxGenJS naming).  */
export const ChartTypeSchema = z.enum([
    "bar", "barStacked", "column", "columnStacked", "line", "lineMarkers",
    "pie", "doughnut", "area", "areaStacked", "radar", "scatter", "bubble"
])

/**  Where data labels sit; which positions a chart type allows is checked
     against the type (PowerPoint rejects the others).  */
export const LabelPositionSchema = z.enum([
    "center", "insideEnd", "insideBase", "outsideEnd", "bestFit", "left", "right", "above", "below"
])

/**  One chart axis: title, value range and number format.  */
const AxisSchema = z.object({
    title: z.string().optional(),
    min: z.number().optional(),
    max: z.number().optional(),
    /**  number format code, e.g. "0%" or "#,##0"  */
    format: z.string().min(1).optional()
}).strict()

/**  the data-label positions per chart type (bar/column stacked or not)  */
const LABEL_POSITIONS: Record<string, readonly string[]> = {
    clustered: [ "center", "insideEnd", "insideBase", "outsideEnd" ],
    stacked: [ "center", "insideEnd", "insideBase" ],
    line: [ "center", "left", "right", "above", "below" ],
    pie: [ "center", "insideEnd", "outsideEnd", "bestFit" ]
}

/**  the label-position family of a chart, null when it takes none  */
const labelFamily = (type: string, stacked: boolean): string | null => {
    if (["bar", "column"].includes(type))
        return stacked ? "stacked" : "clustered"
    if (["barStacked", "columnStacked"].includes(type))
        return "stacked"
    if (["line", "lineMarkers", "scatter", "bubble"].includes(type))
        return "line"
    return type === "pie" ? "pie" : null
}

/**  Chart payload: type, categories and one or more data series.  */
export const ChartSchema = z.object({
    type: ChartTypeSchema,
    categories: z.array(z.union([z.string(), z.number()])).optional(),
    series: z.array(z.object({
        name: z.string(),
        values: z.array(z.number()).optional(),
        x: z.array(z.number()).optional(),
        y: z.array(z.number()).optional(),
        size: z.array(z.number()).optional()
    }).strict()).min(1),
    title: z.string().optional(),
    colors: z.array(ColorSchema).optional(),
    /**  show the legend (`true`: at the bottom) or place it  */
    legend: z.union([z.boolean(), z.enum(["top", "bottom", "left", "right"])]).optional(),
    fontSize: z.number().positive().optional(),
    /**  data labels: values (default when `labels` is given) and/or
         percentages (pie/doughnut), number format, position  */
    labels: z.object({
        value: z.boolean().optional(),
        percent: z.boolean().optional(),
        format: z.string().min(1).optional(),
        position: LabelPositionSchema.optional()
    }).strict().optional(),
    /**  x = category axis (the x value axis of scatter/bubble), y = value axis  */
    axes: z.object({ x: AxisSchema.optional(), y: AxisSchema.optional() }).strict().optional(),
    /**  bar/column/area: stack the series to 100%  */
    stacked: z.literal("percent").optional(),
    /**  line: smoothed lines  */
    smooth: z.boolean().optional(),
    /**  doughnut: hole size in percent of the diameter  */
    holeSize: z.number().int().min(10).max(90).optional()
}).strict().superRefine((c, ctx) => {
    const issue = (path: (string | number)[], message: string): void => {
        ctx.addIssue({ code: "custom", path, message })
    }
    const xy = c.type === "scatter" || c.type === "bubble"
    if (c.axes !== undefined && (c.type === "pie" || c.type === "doughnut"))
        issue([ "axes" ], `a ${c.type} chart has no axes`)
    const x = c.axes?.x
    if (x !== undefined && !xy && (x.min !== undefined || x.max !== undefined || x.format !== undefined))
        issue([ "axes", "x" ], "only the x axis of a scatter or bubble chart takes min, max and format")
    for (const [ key, axis ] of Object.entries(c.axes ?? {}))
        if (axis?.min !== undefined && axis.max !== undefined && axis.min >= axis.max)
            issue([ "axes", key ], "'min' must be below 'max'")
    if (c.labels?.percent === true && c.type !== "pie" && c.type !== "doughnut")
        issue([ "labels", "percent" ], "percent labels need a pie or doughnut chart")
    const position = c.labels?.position
    if (position !== undefined) {
        const family = labelFamily(c.type, c.stacked !== undefined)
        const allowed = family === null ? [] : LABEL_POSITIONS[family] ?? []
        if (!allowed.includes(position))
            issue([ "labels", "position" ], allowed.length === 0
                ? `a ${c.type} chart takes no label position`
                : `a ${c.stacked !== undefined ? "percent-stacked " : ""}${c.type} chart takes the label positions ${allowed.join(", ")}`)
    }
    if (c.stacked !== undefined && !["bar", "barStacked", "column", "columnStacked", "area", "areaStacked"].includes(c.type))
        issue([ "stacked" ], "percent stacking needs a bar, column or area chart")
    if (c.smooth !== undefined && !["line", "lineMarkers"].includes(c.type))
        issue([ "smooth" ], "smoothing needs a line chart")
    /*  scatter and bubble series share one x column (PptxGenJS)  */
    if (xy)
        c.series.forEach((s, i) => {
            if (s.x === undefined || s.y === undefined || s.x.length !== s.y.length)
                issue([ "series", i ], `a ${c.type} series needs 'x' and 'y' of the same length`)
            else if (JSON.stringify(s.x) !== JSON.stringify(c.series[0]?.x))
                issue([ "series", i, "x" ], `all series of a ${c.type} chart share the same 'x' values`)
            if (c.type === "bubble" && s.size?.length !== s.x?.length)
                issue([ "series", i, "size" ], "a bubble series needs one 'size' per point")
        })
    if (c.holeSize !== undefined && c.type !== "doughnut")
        issue([ "holeSize" ], "a hole size needs a doughnut chart")
})

/**  Shape types supported by the shape element (PptxGenJS naming subset).  */
export const ShapeTypeSchema = z.enum([
    "rect", "roundRect", "ellipse", "diamond", "triangle", "rightArrow",
    "leftArrow", "upArrow", "downArrow", "pentagon", "hexagon", "chevron",
    "star5", "cloud", "heart", "flowChartProcess", "flowChartDecision",
    "flowChartTerminator", "line"
])

/**  Line end decoration of a connector.  */
export const ArrowSchema = z.enum(["none", "arrow", "triangle", "stealth", "diamond", "oval"])

/**  Text box layout shared by textbox and shape: anchoring, inner margin
     and text direction.  */
const boxTextProps = {
    valign: ValignSchema.optional(),
    margin: MarginSchema.optional(),
    textDirection: TextDirectionSchema.optional()
}

/**  A free element placed by `el.add`, discriminated by `type`.  */
export const ElementSchema = z.discriminatedUnion("type", [
    z.object({
        type: z.literal("textbox"),
        frame: FrameSchema,
        text: RichTextSchema,
        fill: ColorSchema.optional(),
        border: ColorSchema.optional(),
        borderPt: z.number().positive().optional(),
        ...boxTextProps,
        /**  text language of the element (paragraphs and runs may override)  */
        lang: LangSchema.optional(),
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("table"),
        frame: FrameSchema,
        data: TableSchema,
        /**  column widths in inches (as `state` reports them); scaled to
             `frame.w` when that is given  */
        colWidths: z.array(z.number().positive()).min(1).optional(),
        /**  horizontal alignment per column (a cell's own `align` wins)  */
        colAlign: z.array(CellAlignSchema).min(1).optional(),
        /**  text direction of every cell (a cell's own wins)  */
        textDirection: TextDirectionSchema.optional(),
        /**  text language of every cell  */
        lang: LangSchema.optional(),
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("chart"),
        frame: FrameSchema,
        data: ChartSchema,
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("shape"),
        frame: FrameSchema,
        shape: ShapeTypeSchema,
        text: RichTextSchema.optional(),
        fill: ColorSchema.optional(),
        border: ColorSchema.optional(),
        borderPt: z.number().positive().optional(),
        fontColor: ColorSchema.optional(),
        fontSize: z.number().positive().optional(),
        ...boxTextProps,
        /**  text language of the shape's text  */
        lang: LangSchema.optional(),
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("image"),
        frame: FrameSchema,
        path: z.string().min(1),
        /**  how the image takes the frame: `stretch` (default) distorts it,
             `crop` covers the frame and cuts the overhang off centered,
             `contain` shows the whole image in a centered, shrunk frame  */
        fit: z.enum(["crop", "contain", "stretch"]).optional(),
        /**  click target of the picture  */
        hyperlink: HyperlinkSchema.optional(),
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("connector"),
        from: z.tuple([z.number(), z.number()]),
        to: z.tuple([z.number(), z.number()]),
        color: ColorSchema.optional(),
        widthPt: z.number().positive().optional(),
        dash: z.enum(["solid", "dash", "dot", "dashDot", "longDash"]).optional(),
        /**  decoration at the `from` / `to` end  */
        startArrow: ArrowSchema.optional(),
        endArrow: ArrowSchema.optional(),
        name: z.string().optional()
    }).strict()
])

/**  Inferred payload types (single source of truth: the schemas above).  */
export type RichText = z.infer<typeof RichTextSchema>
export type Paragraph = z.infer<typeof ParagraphSchema>
export type Run = z.infer<typeof RunSchema>
export type TableData = z.infer<typeof TableSchema>
export type TableCell = z.infer<typeof TableCellSchema>
export type TextDirection = z.infer<typeof TextDirectionSchema>
export type ChartData = z.infer<typeof ChartSchema>
export type ElementSpec = z.infer<typeof ElementSchema>
