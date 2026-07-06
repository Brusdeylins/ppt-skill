/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/fill-common: shared planning logic for everything that fills
**  slides -- placeholder key resolution, fill planning and lint, selector
**  resolution against the evolving plan. Used by slide.add, slide.fill and
**  img.prompts.
*/

import { PptcError } from "../../infra/errors.js"
import { resolveSlide } from "../selector.js"
import { lintFontSize, lintPlaceholderText, richTextSizes, richTextToPlain } from "../lint.js"
import { seedPlaceholderName, type Layout, type Placeholder } from "../model.js"
import type { PlaceholderFill } from "../../schema/ops.js"
import { selectableEntries, type PlanContext, type SlidePlanEntry } from "./registry.js"

/**
 *  Resolve a slide selector against the current plan entries.
 *
 *  @param ctx - planning context
 *  @param selector - raw selector string (id:, title:, index:, $ref, digits)
 *  @returns the matching plan entry
 *  @throws PptcError E_ADDR_* when unmatched or ambiguous
 */
export const resolveEntry = (ctx: PlanContext, selector: string): SlidePlanEntry => {
    const view = selectableEntries(ctx)
    const refs = new Map([...ctx.plan.refs].map(([name, entry]) => [name, entry.virtualId]))
    const hit = resolveSlide(selector, view, refs)
    return (view.find((v) => v.id === hit.id) as { entry: SlidePlanEntry }).entry
}

/**
 *  The layout governing a plan entry (template layout for seed slides, deck
 *  layout for kept slides).
 *
 *  @param ctx - planning context
 *  @param entry - the plan entry
 *  @returns the layout, or null when unresolvable
 */
export const entryLayout = (ctx: PlanContext, entry: SlidePlanEntry): Layout | null => {
    /*  seed slides take the template's layouts when a template was given,
        else the deck's OWN layouts (seed-from-deck); kept slides always use
        the deck's layouts  */
    const layouts = entry.source.kind === "seed"
        ? (ctx.template?.layouts ?? ctx.deckLayouts)
        : ctx.deckLayouts
    return layouts[entry.layoutIndex] ?? null
}

/**
 *  Resolve a placeholder key ("0", "title", "subtitle", "body", "image:N",
 *  "text:N") to a concrete placeholder of a layout.
 *
 *  @param layout - the governing layout
 *  @param key - placeholder key from the ops document
 *  @returns the resolved placeholder
 *  @throws PptcError E_ADDR_NOTFOUND / E_ADDR_AMBIGUOUS
 */
export const resolvePlaceholderKey = (layout: Layout, key: string): Placeholder => {
    const all = layout.placeholders
    const byIdx = (idx: number): Placeholder => {
        const ph = all.find((p) => p.idx === idx)
        if (ph === undefined)
            throw new PptcError("E_ADDR_NOTFOUND",
                `layout '${layout.name}' has no placeholder idx ${idx}`,
                { available: all.map((p) => ({ idx: p.idx, kind: p.kind, name: p.name })) })
        return ph
    }
    if (/^\d+$/.test(key))
        return byIdx(Number(key))
    const m = /^(body|image|text):(\d+)$/.exec(key)
    if (m !== null)
        return byIdx(Number(m[2]))
    if (key === "title") {
        const ph = all.find((p) => p.kind === "title")
        if (ph === undefined)
            throw new PptcError("E_ADDR_NOTFOUND", `layout '${layout.name}' has no title placeholder`)
        return ph
    }
    if (key === "subtitle") {
        const ph = all.find((p) => p.kind === "subtitle")
            ?? all.find((p) => p.kind === "body" && p.frame !== null && p.frame.h < 0.8)
        if (ph === undefined)
            throw new PptcError("E_ADDR_NOTFOUND", `layout '${layout.name}' has no subtitle placeholder`)
        return ph
    }
    /*  "body" / "image": unique-kind shorthand  */
    const kind = key === "image" ? "picture" : "body"
    const matches = all.filter((p) => p.kind === kind)
    if (matches.length === 0)
        throw new PptcError("E_ADDR_NOTFOUND", `layout '${layout.name}' has no ${key} placeholder`)
    if (matches.length > 1)
        throw new PptcError("E_ADDR_AMBIGUOUS",
            `layout '${layout.name}' has ${matches.length} ${key} placeholders, use '${key}:N'`,
            { candidates: matches.map((p) => p.idx) })
    return matches[0] as Placeholder
}

/**  the fill-capable fields shared by slide.add and slide.fill payloads  */
export interface FillProps {
    placeholders?: Record<string, PlaceholderFill> | undefined
    notes?: string | undefined
    footer?: string | undefined
    background?: { color: string } | undefined
    hidden?: boolean | undefined
}

/**
 *  The engine-addressable shape name of a placeholder on a planned slide.
 *  Seed slides carry pptc's normalized names; for kept slides the actual
 *  shape name is taken from the deck's read model.
 *
 *  @param ctx - planning context
 *  @param entry - the plan entry the placeholder lives on
 *  @param ph - the resolved layout placeholder
 *  @returns the shape name to address in the engine
 *  @throws PptcError E_ADDR_NOTFOUND when a kept slide does not instantiate
 *          the placeholder
 */
const placeholderShapeName = (ctx: PlanContext, entry: SlidePlanEntry, ph: Placeholder): string => {
    if (entry.source.kind === "seed")
        return seedPlaceholderName(ph.idx)
    const part = entry.source.part
    const slide = ctx.deck.slides.find((s) => s.part === part)
    const shape = slide?.shapes.find((sh) => sh.placeholderIdx === ph.idx)
    if (shape === undefined)
        throw new PptcError("E_ADDR_NOTFOUND",
            `slide has no placeholder idx ${ph.idx} -- it may have been deleted in PowerPoint`,
            { available: slide?.shapes.filter((sh) => sh.placeholderIdx !== null)
                .map((sh) => ({ idx: sh.placeholderIdx, name: sh.name })) ?? [] })
    return shape.name
}

/**
 *  Plan the fill portion of an op onto an entry: resolve placeholder keys,
 *  lint text capacity, track the slide title.
 *
 *  @param ctx - planning context
 *  @param entry - target plan entry
 *  @param fill - placeholders/notes/footer/background payload
 *  @throws PptcError E_ADDR_* on unresolvable placeholders,
 *          E_SCHEMA when text/image is applied to an incompatible kind
 */
export const planFill = (ctx: PlanContext, entry: SlidePlanEntry, fill: FillProps): void => {
    const layout = entryLayout(ctx, entry)
    for (const [key, content] of Object.entries(fill.placeholders ?? {}) as [string, PlaceholderFill][]) {
        if (layout === null)
            throw new PptcError("E_ADDR_NOTFOUND",
                "cannot resolve placeholders: the slide's layout is unknown")
        const ph = resolvePlaceholderKey(layout, key)
        if (content.image !== undefined && ph.kind !== "picture")
            throw new PptcError("E_SCHEMA",
                `placeholder ${ph.idx} ('${ph.name}') is a ${ph.kind} placeholder, not a picture`)
        if (content.text !== undefined && ph.kind === "picture")
            throw new PptcError("E_SCHEMA",
                `placeholder ${ph.idx} ('${ph.name}') is a picture placeholder, cannot take text`)
        if (content.text !== undefined) {
            const slideAddr = { id: entry.virtualId > 0 ? entry.virtualId : null,
                index: ctx.plan.entries.indexOf(entry), title: entry.title }
            const warning = lintPlaceholderText(content.text, ph, slideAddr)
            if (warning !== null)
                ctx.plan.warnings.push(warning)
            /*  warn on runs that override the template size below the readability
                floor; footer/slide-number/date placeholders use footer scale by
                design and are exempt (they reach this path only defensively)  */
            if (ph.kind !== "footer" && ph.kind !== "slideNumber" && ph.kind !== "date") {
                const tooSmall = lintFontSize({ placeholder: ph.idx }, richTextSizes(content.text), ctx.minFontPt, slideAddr)
                if (tooSmall !== null)
                    ctx.plan.warnings.push(tooSmall)
            }
            if (ph.kind === "title")
                entry.title = richTextToPlain(content.text).join(" ")
        }
        entry.fills.push({
            phIdx: ph.idx,
            phName: placeholderShapeName(ctx, entry, ph),
            append: content.append ?? false,
            ...(content.text !== undefined && { text: content.text }),
            ...(content.image !== undefined && { image: content.image }),
            ...(ph.frame !== null && { frame: ph.frame })
        })
    }
    if (fill.notes !== undefined)
        entry.notes = fill.notes
    if (fill.footer !== undefined)
        entry.footer = fill.footer
    if (fill.background !== undefined)
        entry.background = fill.background.color
    if (fill.hidden !== undefined)
        entry.hidden = fill.hidden
}
