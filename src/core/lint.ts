/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/lint: capacity linting of planned text content. Warnings are computed
**  during the plan phase against the same geometry model `tpl describe` uses;
**  --strict escalates them to a hard failure (exit 7).
*/

import type { Frame, Placeholder } from "./model.js"
import type { RichText } from "../schema/payloads.js"
import { estimateUsedLines } from "./describe/capacity.js"

/**  One lint finding, rendered into the envelope's `warnings` array.  */
export interface LintWarning {
    /**  stable warning code  */
    code: "W_TEXT_OVERFLOW" | "W_ELEMENT_OVERLAP" | "W_FONT_TOO_SMALL"
    /**  slide address the finding refers to  */
    slide: { id: number | null, index: number, title: string | null }
    /**  placeholder idx the finding refers to (overflow, font-too-small)  */
    placeholder?: number
    /**  estimated display lines of the planned content (overflow)  */
    estimatedLines?: number
    /**  estimated line capacity of the box (overflow)  */
    maxLines?: number
    /**  name of the added element (overlap, font-too-small)  */
    element?: string
    /**  name of the covered shape or placeholder (overlap)  */
    covers?: string
    /**  smallest explicit run/element font size found, in pt (font-too-small)  */
    fontPt?: number
    /**  the configured font-size floor in pt (font-too-small)  */
    minPt?: number
    /**  human-readable summary  */
    message: string
}

/**  a slide address as it appears on every lint finding  */
type SlideAddr = { id: number | null, index: number, title: string | null }

/**  a geometry obstacle a new element must not cover  */
export interface Obstacle {
    /**  display name of the existing shape/placeholder  */
    name: string
    /**  geometry in inches  */
    frame: Frame
}

/**  intersection area of two frames, 0 when disjoint  */
const intersectionArea = (a: Frame, b: Frame): number => {
    const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)
    const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)
    return w > 0 && h > 0 ? w * h : 0
}

/**
 *  Check a new element's frame against existing text-bearing shapes.
 *  Picture placeholders are NOT obstacles (prompt boxes overlay them by
 *  design); the caller filters them out and exempts prompt boxes.
 *
 *  @param name - display name of the element about to be added
 *  @param frame - the element's geometry in inches
 *  @param obstacles - text-bearing shapes/placeholders on the slide
 *  @param slide - slide address for the report
 *  @returns a warning for the first covered obstacle, otherwise null
 */
export const lintElementOverlap = (
    name: string,
    frame: Frame,
    obstacles: Obstacle[],
    slide: { id: number | null, index: number, title: string | null }
): LintWarning | null => {
    for (const o of obstacles) {
        const area = intersectionArea(frame, o.frame)
        if (area > 0.05) {
            return {
                code: "W_ELEMENT_OVERLAP",
                slide,
                element: name,
                covers: o.name,
                message: `element '${name}' covers '${o.name}' by ~${area.toFixed(1)} in² -- reposition it`
            }
        }
    }
    return null
}

/**
 *  Flatten rich text content into plain paragraph strings.
 *
 *  @param text - plain string or paragraph list
 *  @returns one plain string per paragraph
 */
export const richTextToPlain = (text: RichText): string[] => {
    if (typeof text === "string")
        return text.split("\n")
    return text.map((p) => p.text ?? (p.runs ?? []).map((r) => r.text).join(""))
}

/**
 *  Collect every explicitly set font size (pt) in rich text -- paragraph-level
 *  `size` and per-run `size`. A plain string sets none (inherits template
 *  sizes), so it yields an empty list.
 *
 *  @param text - plain string or paragraph list
 *  @returns the explicit sizes in pt, in document order (may be empty)
 */
export const richTextSizes = (text: RichText): number[] => {
    if (typeof text === "string")
        return []
    const sizes: number[] = []
    for (const p of text) {
        if (p.size !== undefined)
            sizes.push(p.size)
        for (const r of p.runs ?? [])
            if (r.size !== undefined)
                sizes.push(r.size)
    }
    return sizes
}

/**
 *  Check planned text content against a placeholder's capacity.
 *
 *  @param text - the content about to be written
 *  @param ph - the target placeholder (with capacity model)
 *  @param slide - slide address for the report
 *  @returns a warning when the estimate exceeds the capacity, otherwise null
 */
export const lintPlaceholderText = (
    text: RichText,
    ph: Placeholder,
    slide: { id: number | null, index: number, title: string | null }
): LintWarning | null => {
    if (ph.capacity === null)
        return null
    const used = estimateUsedLines(richTextToPlain(text), ph.capacity)
    if (used <= ph.capacity.lines)
        return null
    return {
        code: "W_TEXT_OVERFLOW",
        slide,
        placeholder: ph.idx,
        estimatedLines: used,
        maxLines: ph.capacity.lines,
        message: `text needs ~${used} lines but placeholder ${ph.idx} fits ~${ph.capacity.lines}`
    }
}

/**
 *  Check explicit font sizes against a readability floor. Enforced for free
 *  `el.add` elements and for placeholder runs that override the template size;
 *  footer/slide-number/date placeholders and the tool's own prompt boxes are
 *  exempt (their footer-scale text is intentional) and never reach this check.
 *
 *  @param target - the element name or placeholder idx the finding refers to
 *  @param sizes - all explicit sizes in pt the element/run carries (may be empty)
 *  @param minPt - the floor in pt; `0` or less disables the check
 *  @param slide - slide address for the report
 *  @returns a warning for the smallest below-floor size, otherwise null
 */
export const lintFontSize = (
    target: { element: string } | { placeholder: number },
    sizes: number[],
    minPt: number,
    slide: SlideAddr
): LintWarning | null => {
    if (minPt <= 0 || sizes.length === 0)
        return null
    const smallest = Math.min(...sizes)
    if (smallest >= minPt)
        return null
    if ("element" in target)
        return {
            code: "W_FONT_TOO_SMALL",
            slide,
            element: target.element,
            fontPt: smallest,
            minPt,
            message: `element '${target.element}' sets ${smallest}pt -- below the ${minPt}pt `
                + "minimum for readable slides; enlarge it or lower --min-font-pt"
        }
    return {
        code: "W_FONT_TOO_SMALL",
        slide,
        placeholder: target.placeholder,
        fontPt: smallest,
        minPt,
        message: `placeholder ${target.placeholder} sets ${smallest}pt -- below the ${minPt}pt `
            + "minimum; enlarge it or lower --min-font-pt"
    }
}
