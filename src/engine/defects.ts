/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/defects: the PowerPoint repair triggers of a single part as one
**  rule set. engine/verify reports them on a written deck; the in-place
**  engine detects them on its input and repairs exactly these
**  (engine/inplace/slides.ts) -- one detector, so detection, repair and
**  verification cannot drift apart.
*/

import type { Document } from "@xmldom/xmldom"
import type { Rel } from "./parts.js"
import { elements, inFallback, relRefs } from "./xml.js"

/**  a part as the rules see it  */
export interface PartView {
    /**  the part name  */
    part: string
    /**  its DOM, null for a non-XML part  */
    doc: Document | null
    /**  its relationships  */
    rels: Rel[]
}

/**
 *  Duplicate `p:cNvPr` shape ids on a slide. The `mc:Fallback` copy of a
 *  shape repeats the ids of its `mc:Choice` by design and is ignored.
 *
 *  @param view - the part
 *  @returns findings, empty when clean
 */
export const shapeIdDefects = (view: PartView): string[] => {
    if (view.doc === null || view.doc.documentElement?.nodeName !== "p:sld")
        return []
    const ids = elements(view.doc, "p:cNvPr").filter((el) => !inFallback(el)).map((el) => el.getAttribute("id") ?? "")
    const dup = [ ...new Set(ids.filter((id, i) => ids.indexOf(id) !== i)) ]
    return dup.length > 0 ? [ `${view.part}: duplicate cNvPr ids ${dup.join(",")}` ] : []
}

/**
 *  Duplicate relationship ids and dead internal targets of a part.
 *
 *  @param view - the part
 *  @param exists - whether a part exists
 *  @returns findings, empty when clean
 */
export const relsDefects = (view: PartView, exists: (part: string) => boolean): string[] => {
    const findings: string[] = []
    const ids = view.rels.map((r) => r.id)
    const dup = [ ...new Set(ids.filter((id, i) => ids.indexOf(id) !== i)) ]
    if (dup.length > 0)
        findings.push(`${view.part}: duplicate relationship ids ${dup.join(",")}`)
    for (const r of view.rels)
        if (r.target !== null && !exists(r.target))
            findings.push(`${view.part}: dead relationship target ${r.el.getAttribute("Target") ?? ""}`)
    return findings
}

/**
 *  References to relationship ids a part does not have (a picture, chart
 *  or link of a slide, layout, chart, notes part, the presentation, ...).
 *  An empty id (`r:id=""` of a jump to the next slide) references nothing.
 *
 *  @param view - the part
 *  @returns findings, empty when clean
 */
export const refDefects = (view: PartView): string[] => {
    const root = view.doc?.documentElement ?? null
    if (root === null)
        return []
    const ids = new Set(view.rels.map((r) => r.id))
    const dangling = [ ...new Set(relRefs(root)
        .map((ref) => ref.attr.value).filter((id) => id !== "" && !ids.has(id))) ]
    return dangling.length > 0 ? [ `${view.part}: references missing relationship ids ${dangling.join(",")}` ] : []
}

/**
 *  Every repair trigger of one part.
 *
 *  @param view - the part
 *  @param exists - whether a part exists
 *  @returns findings, empty when clean
 */
export const partDefects = (view: PartView, exists: (part: string) => boolean): string[] =>
    [ ...relsDefects(view, exists), ...refDefects(view), ...shapeIdDefects(view) ]

/**
 *  The slide a notes part belongs to: the target of its slide relationship
 *  that no slide-jump link inside the notes uses.
 *
 *  @param notes - the notes part
 *  @returns the slide part, null when the notes name none
 */
export const notesOwner = (notes: PartView): string | null => {
    const root = notes.doc?.documentElement ?? null
    const used = new Set(root === null ? [] : relRefs(root).map((ref) => ref.attr.value))
    return notes.rels.find((r) => r.type.endsWith("/slide") && !used.has(r.id))?.target ?? null
}
