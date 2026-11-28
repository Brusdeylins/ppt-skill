/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  schema/ops: Zod schemas for the ops document, the single write interface
**  of pptc. Every mutation -- from a one-word title fix to building a whole
**  deck -- is a list of ops validated against these schemas before the first
**  byte of the deck is touched.
*/

import { z } from "zod"
import { ColorSchema, ElementSchema, FramePatchSchema, RichTextSchema } from "./payloads.js"

/**
 *  Slide selector grammar:
 *  `id:N` (canonical sldId) | `title:...` (exact, must be unique) |
 *  `index:N` (zero-based position) | `$ref` (slide created earlier in the
 *  same ops document) | bare digits (shorthand for `index:N`).
 */
export const SlideSelectorSchema = z
    .string()
    .regex(/^(id:\d+|title:.+|index:\d+|\$[A-Za-z][A-Za-z0-9_-]*|\d+)$/,
        "expected 'id:N', 'title:...', 'index:N', '$ref' or a bare index")

/**
 *  Placeholder address inside `slide.fill`: the numeric OOXML `idx`, or a
 *  semantic key as reported by `tpl describe` ("title", "subtitle"), or a
 *  kind:idx combination like "image:14".
 */
export const PlaceholderKeySchema = z
    .string()
    .regex(/^(\d+|title|subtitle|body(:\d+)?|image(:\d+)?|text:\d+)$/,
        "expected a placeholder idx, 'title', 'subtitle', 'body', 'image:N' or 'text:N'")

/**  Content assigned to a single placeholder by `slide.fill`.  */
export const PlaceholderFillSchema = z.object({
    /**  rich or plain text content (text placeholders)  */
    text: RichTextSchema.optional(),
    /**  path of an image file to insert (picture placeholders)  */
    image: z.string().min(1).optional(),
    /**  how the image takes the placeholder frame: `crop` (default) covers
         it and cuts the overhang off centered, `contain` shows the whole
         image in a centered, shrunk frame, `stretch` distorts it to fit  */
    fit: z.enum(["crop", "contain", "stretch"]).optional(),
    /**  append instead of replacing existing content  */
    append: z.boolean().optional()
}).strict().refine((f) => f.text !== undefined || f.image !== undefined,
    "a placeholder fill needs 'text' or 'image'")
    .refine((f) => f.fit === undefined || f.image !== undefined, "'fit' needs an 'image'")

/**  Shared shape of the `fill` payload used by slide.add and slide.fill.  */
const fillProps = {
    placeholders: z.record(PlaceholderKeySchema, PlaceholderFillSchema).optional(),
    notes: z.string().optional(),
    footer: z.string().optional(),
    background: z.object({ color: ColorSchema }).strict().optional(),
    /**  hide ("Hide Slide", `show="0"`) or show the slide; omitted leaves the
         current visibility untouched (kept slides keep their hidden state)  */
    hidden: z.boolean().optional(),
    /**  show the slide number (`true`: the layout's slide-number placeholder
         with its field) or remove it (`false`); omitted leaves it as it is  */
    slideNumber: z.boolean().optional(),
    /**  show the date (`true`: the layout's date placeholder with its field)
         or remove it (`false`); omitted leaves it as it is  */
    date: z.boolean().optional()
}

/**  Op: add a new slide based on a template layout.  */
export const SlideAddSchema = z.object({
    op: z.literal("slide.add"),
    /**  layout address: zero-based index or exact layout name  */
    layout: z.union([z.number().int().min(0), z.string().min(1)]),
    /**  document-local name to address this slide in later ops via `$ref`  */
    ref: z.string().regex(/^[A-Za-z][A-Za-z0-9_-]*$/).optional(),
    /**  zero-based insert position; appends when omitted  */
    at: z.number().int().min(0).optional(),
    ...fillProps
}).strict()

/**  Op: fill placeholders, notes, footer or background of an existing slide.  */
export const SlideFillSchema = z.object({
    op: z.literal("slide.fill"),
    slide: SlideSelectorSchema,
    ...fillProps
}).strict()

/**  Op: remove a slide.  */
export const SlideRmSchema = z.object({
    op: z.literal("slide.rm"),
    slide: SlideSelectorSchema
}).strict()

/**  Op: move a slide to a new zero-based position.  */
export const SlideMoveSchema = z.object({
    op: z.literal("slide.move"),
    slide: SlideSelectorSchema,
    to: z.number().int().min(0)
}).strict()

/**  Op: duplicate a slide; the copy is inserted after the source.  */
export const SlideCopySchema = z.object({
    op: z.literal("slide.copy"),
    slide: SlideSelectorSchema,
    ref: z.string().regex(/^[A-Za-z][A-Za-z0-9_-]*$/).optional()
}).strict()

/**  Op: add free elements (table, chart, textbox, shape, image, connector).  */
export const ElAddSchema = z.object({
    op: z.literal("el.add"),
    slide: SlideSelectorSchema,
    elements: z.array(ElementSchema).min(1)
}).strict()

/**  Op: remove a named element from a slide.  */
export const ElRmSchema = z.object({
    op: z.literal("el.rm"),
    slide: SlideSelectorSchema,
    name: z.string().min(1)
}).strict()

/**  A table cell value of `el.set` rows: rich text or a plain value.  */
const RowCellSchema = z.union([RichTextSchema, z.number(), z.boolean()])

/**  the fields of `el.set` that change something  */
const EL_SET_FIELDS = [ "text", "frame", "fill", "border", "borderPt", "fontColor", "altText", "title",
    "image", "cells", "addRows", "removeRows" ] as const

/**  Op: change a named element on a slide in place -- its text, frame,
     style, alt text, picture or table cells/rows; each field is checked
     against the element's type.  */
export const ElSetSchema = z.object({
    op: z.literal("el.set"),
    slide: SlideSelectorSchema,
    name: z.string().min(1),
    /**  new text (placeholders, textboxes, shapes); keeps the first
         paragraph's and run's format where the new text sets none  */
    text: RichTextSchema.optional(),
    /**  move, resize or rotate: the given keys change, the others stay  */
    frame: FramePatchSchema.optional(),
    /**  solid fill (placeholders, textboxes, shapes)  */
    fill: ColorSchema.optional(),
    /**  outline color and width in pt (also pictures and connectors)  */
    border: ColorSchema.optional(),
    borderPt: z.number().positive().optional(),
    /**  color of every text run (placeholders, textboxes, shapes)  */
    fontColor: ColorSchema.optional(),
    /**  alt text (`cNvPr@descr`) and title (`cNvPr@title`) of any element;
         an empty string removes them  */
    altText: z.string().optional(),
    title: z.string().optional(),
    /**  replace the image of a picture or a filled picture placeholder; the
         frame, a crop (re-cropped to the new aspect) and animations stay  */
    image: z.string().min(1).optional(),
    /**  table: set cell texts by 0-based grid position (header row 0)  */
    cells: z.array(z.object({
        row: z.number().int().min(0),
        col: z.number().int().min(0),
        text: RichTextSchema
    }).strict()).min(1).optional(),
    /**  table: insert rows at a 0-based row position (default: append);
         they take the format of the row above (the first row at 0)  */
    addRows: z.object({
        at: z.number().int().min(0).optional(),
        rows: z.array(z.array(RowCellSchema).min(1)).min(1)
    }).strict().optional(),
    /**  table: remove rows by 0-based index (before addRows and cells)  */
    removeRows: z.array(z.number().int().min(0)).min(1).optional()
}).strict().refine((op) => EL_SET_FIELDS.some((f) => op[f] !== undefined),
    `el.set needs at least one of ${EL_SET_FIELDS.join(", ")}`)

/**  Op: find and replace plain text across the deck or on one slide --
     shapes, placeholders and table cells, the speaker notes on request;
     fields (slide number, date) never.  */
export const TextReplaceSchema = z.object({
    op: z.literal("text.replace"),
    find: z.string().min(1),
    replace: z.string(),
    /**  only this slide; the whole deck when omitted  */
    slide: SlideSelectorSchema.optional(),
    /**  match upper/lower case exactly (default false)  */
    matchCase: z.boolean().optional(),
    /**  match whole words only (default false)  */
    wholeWord: z.boolean().optional(),
    /**  also replace in the speaker notes (default false)  */
    notes: z.boolean().optional()
}).strict()

/**  Op: overlay picture placeholders with visible image-prompt boxes.  */
export const ImgPromptsSchema = z.object({
    op: z.literal("img.prompts"),
    slide: SlideSelectorSchema,
    /**  prompt text per picture-placeholder idx, or one prompt for all  */
    prompts: z.union([z.string(), z.record(z.string().regex(/^\d+$/), z.string())])
}).strict()

/**  Op: set document properties -- standard core fields and/or arbitrary
     custom name/value pairs (stored in docProps/custom.xml, so they travel
     inside the .pptx and survive PowerPoint round-trips).  */
export const MetaPropsSchema = z.object({
    op: z.literal("meta.props"),
    set: z.object({
        title: z.string().optional(),
        author: z.string().optional(),
        subject: z.string().optional(),
        keywords: z.string().optional(),
        category: z.string().optional(),
        comments: z.string().optional(),
        /**  arbitrary custom document properties (name→value); a value
             may be empty to clear the visible content of a property  */
        custom: z.record(z.string().min(1), z.string()).optional()
    }).strict()
}).strict()

/**  Any single op, discriminated by the `op` field.  */
export const OpSchema = z.discriminatedUnion("op", [
    SlideAddSchema, SlideFillSchema, SlideRmSchema, SlideMoveSchema,
    SlideCopySchema, ElAddSchema, ElRmSchema, ElSetSchema,
    ImgPromptsSchema, MetaPropsSchema, TextReplaceSchema
])

/**  The ops document: optional revision lock plus the ordered op list.  */
export const OpsDocumentSchema = z.object({
    /**  optimistic lock: must equal the deck's current `rev` when given  */
    expectRev: z.string().optional(),
    ops: z.array(OpSchema).min(1)
}).strict()

/**  Inferred op types (single source of truth: the schemas above).  */
export type Op = z.infer<typeof OpSchema>
export type OpsDocument = z.infer<typeof OpsDocumentSchema>
export type SlideAddOp = z.infer<typeof SlideAddSchema>
export type SlideFillOp = z.infer<typeof SlideFillSchema>
export type PlaceholderFill = z.infer<typeof PlaceholderFillSchema>
export type ElSetOp = z.infer<typeof ElSetSchema>
export type TextReplaceOp = z.infer<typeof TextReplaceSchema>

/**  All op names, used by `pptc schema` for enumeration.  */
export const OP_NAMES = [
    "slide.add", "slide.fill", "slide.rm", "slide.move", "slide.copy",
    "el.add", "el.rm", "el.set", "img.prompts", "meta.props", "text.replace"
] as const
