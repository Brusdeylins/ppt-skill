/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/elements: plan transformers for the element-level ops --
**  el.add (generate new elements), el.set (retext) and el.rm (remove),
**  plus img.prompts (prompt-box overlays on picture placeholders).
*/

import { PptcError } from "../../infra/errors.js"
import type { Op } from "../../schema/ops.js"
import type { ElementSpec } from "../../schema/payloads.js"
import type { Frame, ShapeInfo } from "../model.js"
import { slideAddrOf, type OpHandler, type PlanContext, type SlidePlanEntry } from "./registry.js"
import { entryLayout, resolveEntry } from "./fill-common.js"
import { lintElementOverlap, lintFontSize, richTextSizes, type Obstacle } from "../lint.js"
import { nearestAspect } from "../describe/position.js"
import { layoutTable } from "../table-grid.js"

/**  shape-name prefix of generated prompt boxes (removable via el.rm)  */
export const PROMPT_BOX_PREFIX = "PptcPromptBox"

/**  the frame of an element spec -- for a rotated element the axis-aligned
     box it covers (90/270 degrees swap width and height); null for
     connectors (line geometry) and for boxes without explicit extent
     (auto-height tables)  */
const specFrame = (spec: ElementSpec): Frame | null => {
    if (spec.type === "connector")
        return null
    const f = spec.frame
    if (f.w === undefined || f.h === undefined)
        return null
    const rad = (f.rotation ?? 0) * Math.PI / 180
    const cos = Math.abs(Math.cos(rad))
    const sin = Math.abs(Math.sin(rad))
    const w = f.w * cos + f.h * sin
    const h = f.w * sin + f.h * cos
    return { x: f.x + (f.w - w) / 2, y: f.y + (f.h - h) / 2, w, h }
}

/**  check a table's merged cells and its per-column fields against its
     grid (E_SCHEMA): every row fills the columns of the first row, spans
     stay inside the table, `colWidths`/`colAlign` give one value per column  */
const checkTable = (spec: Extract<ElementSpec, { type: "table" }>, name: string): void => {
    const rows: unknown[][] = [ ...(spec.data.headers !== undefined ? [ spec.data.headers ] : []), ...spec.data.rows ]
    const grid = layoutTable(rows)
    if (typeof grid === "string")
        throw new PptcError("E_SCHEMA", `table '${name}': ${grid}`)
    for (const field of [ "colWidths", "colAlign" ] as const) {
        const values = spec[field]
        if (values !== undefined && values.length !== grid.cols)
            throw new PptcError("E_SCHEMA",
                `table '${name}': '${field}' has ${values.length} value(s), the table has ${grid.cols} column(s)`)
    }
}

/**  every explicitly set font size (pt) an element carries, for the readability
     lint: textbox runs/paragraphs, shape/table/chart font sizes; empty when the
     element sets none (image/connector, or a spec that inherits template sizes)  */
const collectElementSizes = (spec: ElementSpec): number[] => {
    switch (spec.type) {
        case "textbox": return richTextSizes(spec.text)
        case "shape": return [ ...(spec.fontSize !== undefined ? [spec.fontSize] : []),
            ...(spec.text !== undefined ? richTextSizes(spec.text) : []) ]
        case "table": return [ ...(spec.data.style?.fontSize !== undefined ? [spec.data.style.fontSize] : []),
            ...[ ...(spec.data.headers ?? []), ...spec.data.rows.flat() ]
                .flatMap((cell) => typeof cell === "object" ? richTextSizes(cell.text) : []) ]
        case "chart": return spec.data.fontSize !== undefined ? [spec.data.fontSize] : []
        default: return []
    }
}

/**  text-bearing obstacles a new element must not cover: text placeholders
     (incl. footer/slide-number/date), existing text shapes and elements
     planned earlier -- picture placeholders and prompt boxes are exempt  */
const overlapObstacles = (ctx: PlanContext, entry: SlidePlanEntry): Obstacle[] => {
    const obstacles: Obstacle[] = []
    const layout = entryLayout(ctx, entry)
    if (layout !== null) {
        for (const ph of layout.placeholders)
            if (ph.kind !== "picture" && ph.frame !== null)
                obstacles.push({ name: ph.name, frame: ph.frame })
        for (const frame of layout.reserved)
            obstacles.push({ name: "footer/slide-number/date area", frame })
    }
    if (entry.source.kind === "self") {
        const part = entry.source.part
        for (const sh of ctx.deck.slides.find((s) => s.part === part)?.shapes ?? [])
            if (sh.frame !== null && !sh.name.startsWith(PROMPT_BOX_PREFIX)
                && sh.placeholderKind !== "picture" && sh.type !== "picture"
                && sh.type !== "connector"
                /*  shapes removed earlier in the same document are gone  */
                && !entry.removeNames.some((rm) => nameMatches(sh.name, rm)))
                obstacles.push({ name: sh.name, frame: sh.frame })
    }
    for (const planned of entry.elements) {
        const frame = specFrame(planned.spec)
        if (frame !== null && (planned.name === null || !planned.name.startsWith(PROMPT_BOX_PREFIX)))
            obstacles.push({ name: planned.name ?? planned.spec.type, frame })
    }
    return obstacles
}

/**  op handler: el.add  */
export const elAdd: OpHandler<Extract<Op, { op: "el.add" }>> = {
    name: "el.add",
    plan(ctx, op): void {
        const entry = resolveEntry(ctx, op.slide)
        for (const spec of op.elements) {
            const frame = specFrame(spec)
            const name = spec.name ?? spec.type
            if (spec.type === "table")
                checkTable(spec, name)
            /*  prompt boxes are exempt from both checks -- they overlay picture
                placeholders by design and their footer-scale caption is intentional  */
            if (!name.startsWith(PROMPT_BOX_PREFIX)) {
                const slideAddr = slideAddrOf(ctx, entry)
                /*  warn when the new element covers a text-bearing shape  */
                if (frame !== null) {
                    const overlap = lintElementOverlap(name, frame, overlapObstacles(ctx, entry), slideAddr)
                    if (overlap !== null)
                        ctx.plan.warnings.push(overlap)
                }
                /*  warn when any run/element font is below the readability floor  */
                const tooSmall = lintFontSize({ element: name }, collectElementSizes(spec), ctx.minFontPt, slideAddr)
                if (tooSmall !== null)
                    ctx.plan.warnings.push(tooSmall)
            }
            entry.elements.push({ name: spec.name ?? null, spec })
        }
    }
}

/**
 *  Match a shape name against an addressed element name. The engine appends
 *  a UUID to generated shape names ("LinkBox-1d22c8b0-..."), so an address
 *  matches exactly or as the prefix before such a suffix.
 *
 *  @param shapeName - actual shape name on the slide
 *  @param address - element name given in the op
 *  @returns true when the shape is addressed by the name
 */
const nameMatches = (shapeName: string, address: string): boolean =>
    shapeName === address || shapeName.startsWith(`${address}-`)

/**  the existing shape types el.set can retext  */
const TEXT_TYPES: ShapeInfo["type"][] = [ "placeholder", "textbox", "shape" ]

/**  preset geometries of lines and connectors  */
const LINE_GEOMETRY = /^(line|straightConnector\d|bentConnector\d|curvedConnector\d)$/

/**  existing shapes of a kept slide, empty for new (seed) slides  */
const existingShapes = (ctx: PlanContext, entry: SlidePlanEntry): ShapeInfo[] => {
    if (entry.source.kind !== "self")
        return []
    const part = entry.source.part
    return ctx.deck.slides.find((s) => s.part === part)?.shapes ?? []
}

/**  op handler: el.set  */
export const elSet: OpHandler<Extract<Op, { op: "el.set" }>> = {
    name: "el.set",
    plan(ctx, op): void {
        const entry = resolveEntry(ctx, op.slide)

        /*  the new text is rendered verbatim on both paths below, so it gets
            the same readability lint as el.add (prompt boxes stay exempt)  */
        if (!op.name.startsWith(PROMPT_BOX_PREFIX)) {
            const tooSmall = lintFontSize({ element: op.name }, richTextSizes(op.text),
                ctx.minFontPt, slideAddrOf(ctx, entry))
            if (tooSmall !== null)
                ctx.plan.warnings.push(tooSmall)
        }

        /*  an element generated earlier in the same run: patch its spec  */
        const planned = entry.elements.find((e) => e.name !== null && nameMatches(e.name, op.name))
        if (planned !== undefined) {
            if (planned.spec.type === "textbox" || planned.spec.type === "shape")
                planned.spec.text = op.text
            else
                throw new PptcError("E_SCHEMA",
                    `element '${op.name}' is a ${planned.spec.type} -- only textbox and shape text can be set`)
            return
        }
        /*  otherwise the element must already exist on the slide  */
        const matches = existingShapes(ctx, entry).filter((sh) => nameMatches(sh.name, op.name))
        if (matches.length === 0)
            throw new PptcError("E_ADDR_NOTFOUND",
                `slide has no element named '${op.name}'`,
                { available: existingShapes(ctx, entry).map((sh) => sh.name) })
        if (matches.length > 1)
            throw new PptcError("E_ADDR_AMBIGUOUS",
                `'${op.name}' matches ${matches.length} elements`,
                { candidates: matches.map((sh) => sh.name) })
        /*  only shapes with a text body of their own take text: a table,
            chart, picture, connector or group would be left unchanged (or a
            group's first child retexted); el.add draws a connector as a line
            autoshape  */
        const target = matches[0] as ShapeInfo
        const type = target.type === "shape" && LINE_GEOMETRY.test(target.shape ?? "") ? "connector" : target.type
        if (!TEXT_TYPES.includes(type))
            throw new PptcError("E_SCHEMA",
                `element '${target.name}' is a ${type} -- only placeholder, textbox and shape text can be set`)
        entry.setTexts.push({ name: target.name, text: op.text })
    }
}

/**  op handler: el.rm  */
export const elRm: OpHandler<Extract<Op, { op: "el.rm" }>> = {
    name: "el.rm",
    plan(ctx, op): void {
        const entry = resolveEntry(ctx, op.slide)
        /*  cancel elements generated earlier in the same run  */
        const plannedBefore = entry.elements.length
        entry.elements = entry.elements.filter((e) =>
            e.name === null || !nameMatches(e.name, op.name))
        const cancelled = plannedBefore - entry.elements.length
        /*  remove all matching existing elements (full names, incl. UUID suffix)  */
        const matches = existingShapes(ctx, entry).filter((sh) => nameMatches(sh.name, op.name))
        for (const shape of matches)
            entry.removeNames.push(shape.name)
        if (cancelled === 0 && matches.length === 0)
            throw new PptcError("E_ADDR_NOTFOUND",
                `slide has no element named '${op.name}'`,
                { available: existingShapes(ctx, entry).map((sh) => sh.name) })
    }
}

/**  visual style of prompt boxes (light yellow, accent border)  */
const PROMPT_FILL = "FFFBE6"
const PROMPT_BORDER = "C00000"

/**  op handler: img.prompts  */
export const imgPrompts: OpHandler<Extract<Op, { op: "img.prompts" }>> = {
    name: "img.prompts",
    plan(ctx, op): void {
        const entry = resolveEntry(ctx, op.slide)
        const layout = entryLayout(ctx, entry)
        if (layout === null)
            throw new PptcError("E_ADDR_NOTFOUND",
                "cannot place prompt boxes: the slide's layout is unknown")
        const pics = layout.placeholders.filter((p) => p.kind === "picture" && p.frame !== null)
        if (pics.length === 0)
            throw new PptcError("E_ADDR_NOTFOUND",
                `layout '${layout.name}' has no picture placeholders`)
        const promptFor = (idx: number): string | null =>
            typeof op.prompts === "string" ? op.prompts : op.prompts[String(idx)] ?? null
        for (const pic of pics) {
            const text = promptFor(pic.idx)
            if (text === null)
                continue
            const frame = pic.frame as { x: number, y: number, w: number, h: number }
            const aspect = nearestAspect(frame)
            const spec: ElementSpec = {
                type: "textbox",
                frame: { x: frame.x, y: frame.y, w: Math.max(frame.w * 0.5, 2), h: Math.max(frame.h * 0.4, 0.8) },
                text: [
                    { text: `IMAGE PROMPT · ${aspect}`, bold: true, size: 9, color: PROMPT_BORDER },
                    { text, size: 9, color: "4A4A4A" }
                ],
                fill: PROMPT_FILL,
                border: PROMPT_BORDER,
                borderPt: 1
            }
            entry.elements.push({ name: `${PROMPT_BOX_PREFIX}-${pic.idx}`, spec })
        }
    }
}

/**  op handler: meta.props -- core fields patch `plan.props`, the `custom`
     map patches `plan.customProps` (written to docProps/custom.xml)  */
export const metaProps: OpHandler<Extract<Op, { op: "meta.props" }>> = {
    name: "meta.props",
    plan(ctx, op): void {
        const { custom, ...core } = op.set
        const coreEntries = Object.entries(core).filter(([, v]) => v !== undefined)
        if (coreEntries.length > 0) {
            const patch: Record<string, string> = { ...(ctx.plan.props ?? {}) }
            for (const [key, value] of coreEntries)
                patch[key] = value as string
            ctx.plan.props = patch
        }
        if (custom !== undefined) {
            const patch: Record<string, string> = { ...(ctx.plan.customProps ?? {}) }
            for (const [key, value] of Object.entries(custom))
                patch[key] = value
            ctx.plan.customProps = patch
        }
    }
}
