/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/defects: the PowerPoint repair triggers of a single part as one
**  rule set. engine/verify reports them on a written deck; the in-place
**  engine detects them on its input and repairs exactly these
**  (engine/inplace/repair.ts) -- one detector, so detection, repair and
**  verification cannot drift apart.
*/

import type { Document } from "@xmldom/xmldom"
import type { Rel } from "./parts.js"
import { elements, inFallback, relRefs } from "./xml.js"

/**  one repair trigger: its rule and part identify it (a finding the input
     deck already carried is recognized by both), the text explains it  */
export interface Finding {
    /**  the rule, e.g. "rels-dead"  */
    rule: string
    /**  the part it concerns  */
    part: string
    /**  the human-readable finding  */
    text: string
}

/**
 *  A finding.
 *
 *  @param rule - the rule
 *  @param part - the part
 *  @param detail - what is wrong, after the part name
 *  @returns the finding, its text `<part>: <detail>`
 */
export const finding = (rule: string, part: string, detail: string): Finding =>
    ({ rule, part, text: `${part}: ${detail}` })

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
export const shapeIdDefects = (view: PartView): Finding[] => {
    if (view.doc === null || view.doc.documentElement?.nodeName !== "p:sld")
        return []
    const ids = elements(view.doc, "p:cNvPr").filter((el) => !inFallback(el)).map((el) => el.getAttribute("id") ?? "")
    const dup = [ ...new Set(ids.filter((id, i) => ids.indexOf(id) !== i)) ]
    return dup.map((id) => finding("shape-ids", view.part, `duplicate cNvPr id ${id}`))
}

/**
 *  Duplicate relationship ids and dead internal targets of a part.
 *
 *  @param view - the part
 *  @param exists - whether a part exists
 *  @returns findings, empty when clean
 */
export const relsDefects = (view: PartView, exists: (part: string) => boolean): Finding[] => {
    const findings: Finding[] = []
    const ids = view.rels.map((r) => r.id)
    const dup = [ ...new Set(ids.filter((id, i) => ids.indexOf(id) !== i)) ]
    for (const id of dup)
        findings.push(finding("rels-dup", view.part, `duplicate relationship id ${id}`))
    for (const r of view.rels)
        if (r.target !== null && !exists(r.target))
            findings.push(finding("rels-dead", view.part, `dead relationship target ${r.target}`))
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
export const refDefects = (view: PartView): Finding[] => {
    const root = view.doc?.documentElement ?? null
    if (root === null)
        return []
    const ids = new Set(view.rels.map((r) => r.id))
    const dangling = [ ...new Set(relRefs(root)
        .map((ref) => ref.attr.value).filter((id) => id !== "" && !ids.has(id))) ]
    return dangling.map((id) => finding("refs", view.part, `references missing relationship id ${id}`))
}

/**
 *  Every repair trigger of one part.
 *
 *  @param view - the part
 *  @param exists - whether a part exists
 *  @returns findings, empty when clean
 */
export const partDefects = (view: PartView, exists: (part: string) => boolean): Finding[] =>
    [ ...relsDefects(view, exists), ...refDefects(view), ...shapeIdDefects(view) ]

/**
 *  The slide a notes part belongs to: the target of its slide relationship
 *  that no slide-jump link inside the notes uses -- or, when a jump link
 *  reuses it, of its first slide relationship.
 *
 *  @param notes - the notes part
 *  @returns the slide part, null when the notes name none
 */
export const notesOwner = (notes: PartView): string | null => {
    const root = notes.doc?.documentElement ?? null
    const used = new Set(root === null ? [] : relRefs(root).map((ref) => ref.attr.value))
    const slides = notes.rels.filter((r) => r.kind === "slide")
    return (slides.find((r) => !used.has(r.id)) ?? slides[0])?.target ?? null
}
