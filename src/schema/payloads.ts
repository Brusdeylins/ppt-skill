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

/**  6-digit hex color, with or without leading `#`, normalized to RRGGBB.  */
export const ColorSchema = z
    .string()
    .regex(/^#?[0-9a-fA-F]{6}$/, "expected 6-digit hex color like '#1F4E79'")
    .transform((s) => s.replace("#", "").toUpperCase())

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

/**  Character-level formatting shared by paragraphs and runs.  */
const fontProps = {
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    size: z.number().positive().optional(),
    font: z.string().min(1).optional(),
    color: ColorSchema.optional(),
    /**  text language; a paragraph's applies to its runs  */
    lang: LangSchema.optional()
}

/**  One formatted text run inside a paragraph.  */
export const RunSchema = z.object({
    text: z.string(),
    hyperlink: z.string().url().optional(),
    ...fontProps
}).strict()

/**  One paragraph of rich text: plain `text` or formatted `runs`.  */
export const ParagraphSchema = z.object({
    text: z.string().optional(),
    runs: z.array(RunSchema).optional(),
    level: z.number().int().min(0).max(8).optional(),
    alignment: z.enum(["left", "center", "right", "justify"]).optional(),
    bullet: z.boolean().optional(),
    ...fontProps
}).strict().refine((p) => (p.text !== undefined) !== (p.runs !== undefined),
    "a paragraph needs either 'text' or 'runs' (not both)")

/**  Rich text content: a single plain string or a list of paragraphs.  */
export const RichTextSchema = z.union([z.string(), z.array(ParagraphSchema).min(1)])

/**  Table payload: headers, rows and optional styling.  */
export const TableSchema = z.object({
    headers: z.array(z.string()).optional(),
    rows: z.array(z.array(z.union([z.string(), z.number(), z.boolean()]))).min(1),
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
    legend: z.boolean().optional(),
    fontSize: z.number().positive().optional()
}).strict()

/**  Shape types supported by the shape element (PptxGenJS naming subset).  */
export const ShapeTypeSchema = z.enum([
    "rect", "roundRect", "ellipse", "diamond", "triangle", "rightArrow",
    "leftArrow", "upArrow", "downArrow", "pentagon", "hexagon", "chevron",
    "star5", "cloud", "heart", "flowChartProcess", "flowChartDecision",
    "flowChartTerminator", "line"
])

/**  A free element placed by `el.add`, discriminated by `type`.  */
export const ElementSchema = z.discriminatedUnion("type", [
    z.object({
        type: z.literal("textbox"),
        frame: FrameSchema,
        text: RichTextSchema,
        fill: ColorSchema.optional(),
        border: ColorSchema.optional(),
        borderPt: z.number().positive().optional(),
        /**  text language of the element (paragraphs and runs may override)  */
        lang: LangSchema.optional(),
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("table"),
        frame: FrameSchema,
        data: TableSchema,
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
        text: z.string().optional(),
        fill: ColorSchema.optional(),
        border: ColorSchema.optional(),
        borderPt: z.number().positive().optional(),
        fontColor: ColorSchema.optional(),
        fontSize: z.number().positive().optional(),
        /**  text language of the shape's text  */
        lang: LangSchema.optional(),
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("image"),
        frame: FrameSchema,
        path: z.string().min(1),
        name: z.string().optional()
    }).strict(),
    z.object({
        type: z.literal("connector"),
        from: z.tuple([z.number(), z.number()]),
        to: z.tuple([z.number(), z.number()]),
        color: ColorSchema.optional(),
        widthPt: z.number().positive().optional(),
        dash: z.enum(["solid", "dash", "dot"]).optional(),
        name: z.string().optional()
    }).strict()
])

/**  Inferred payload types (single source of truth: the schemas above).  */
export type RichText = z.infer<typeof RichTextSchema>
export type Paragraph = z.infer<typeof ParagraphSchema>
export type Run = z.infer<typeof RunSchema>
export type TableData = z.infer<typeof TableSchema>
export type ChartData = z.infer<typeof ChartSchema>
export type ElementSpec = z.infer<typeof ElementSchema>
