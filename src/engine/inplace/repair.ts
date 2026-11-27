/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/repair: the repairs of the triggers engine/defects
**  detects -- duplicate and dead relationships, references to missing
**  relationships, duplicate shape ids -- on any part (slides, their notes
**  and charts, layouts, masters, the presentation). A dead reference takes
**  down the smallest element that keeps the XML schema-valid, and every
**  repair names what it removed.
*/

import type { Document, Element } from "@xmldom/xmldom"
import { PartAllocator, RELS, relIs } from "../parts.js"
import { elements, firstElement, inFallback, relRefs, serializeXml } from "../xml.js"
import type { Package } from "./package.js"

/**  the elements a shape tree is made of: the smallest unit a dead
     reference to its main content can take down  */
const SHAPES = [ "p:sp", "p:pic", "p:graphicFrame", "p:cxnSp", "p:grpSp", "p:contentPart" ]

/**  elements that only link or tag: they go alone  */
const LINKS = [ "a:hlinkClick", "a:hlinkHover", "a:hlinkMouseOver", "p:custData", "p:tags" ]

/**  elements whose references are all required (a SmartArt needs its four
     parts): one dead reference takes down the whole shape  */
const REQUIRED = [ "dgm:relIds" ]

/**  the nearest ancestor (or self) with one of the given names  */
const closest = (el: Element, names: string[]): Element | null => {
    for (let n: Element | null = el; n !== null && n.nodeType === 1; n = n.parentNode as Element | null)
        if (names.includes(n.nodeName))
            return n
    return null
}

/**  whether an element is still part of its document  */
const attached = (el: Element, root: Element): boolean => {
    for (let n: Element | null = el; n !== null; n = n.parentNode as Element | null)
        if (n === root)
            return true
    return false
}

/**
 *  Remove a shape; inside an `mc:AlternateContent` (OLE objects, SmartArt)
 *  the whole wrapper goes, so no fallback copy stays behind.
 *
 *  @param shape - the shape element
 */
export const removeShape = (shape: Element): void => {
    let doomed = shape
    for (let up = shape.parentNode as Element | null; up !== null && up.nodeName !== "p:spTree" && up.nodeName !== "p:grpSp";
        up = up.parentNode as Element | null)
        if (up.nodeName === "mc:AlternateContent")
            doomed = up
    doomed.parentNode?.removeChild(doomed)
}

/**  what a dead reference takes down: the first rule that applies decides  */
const SCOPES: ((el: Element) => { doomed: Element, what: string } | null)[] = [
    /*  a link or tag goes alone  */
    (el) => LINKS.includes(el.nodeName) ? { doomed: el, what: "a link" } : null,

    /*  animations referencing a missing sound or target go as a whole  */
    (el) => { const t = closest(el, [ "p:timing" ]); return t === null ? null : { doomed: t, what: "the animations" } },

    /*  a dead fallback goes, its choice stays  */
    (el) => { const f = closest(el, [ "mc:Fallback" ]); return f === null ? null : { doomed: f, what: "a fallback" } },

    /*  an extension (e.g. the SVG of a picture) goes, its owner stays  */
    (el) => { const x = closest(el, [ "a:ext", "p:ext" ]); return x === null ? null : { doomed: x, what: "an extension" } },

    /*  a transition sound goes with its sound action  */
    (el) => { const s = closest(el, [ "p:sndAc" ]); return s === null ? null : { doomed: s, what: "a transition sound" } },

    /*  a picture bullet goes, the text keeps its default bullet  */
    (el) => { const b = closest(el, [ "a:buBlip" ]); return b === null ? null : { doomed: b, what: "a picture bullet" } },

    /*  a picture fill of a background, shape, group, cell or text goes;
        the layout's background or the default fill shows again  */
    (el) => {
        const fill = el.nodeName === "a:blip" && (el.parentNode as Element | null)?.nodeName === "a:blipFill" ? el.parentNode as Element : null
        if (fill === null)
            return null
        const holder = fill.parentNode as Element | null
        return holder?.nodeName === "p:bgPr" ? { doomed: holder.parentNode as Element, what: "a background picture" }
            : { doomed: fill, what: "a picture fill" }
    }
]

/**  drop what references a relationship the part does not have (rules
     above; a dead attribute next to a valid reference goes alone; anything
     else takes down its smallest shape, or outside shapes itself)
     @returns what was dropped  */
const dropDangling = (doc: Document, rels: Document): string[] => {
    const root = doc.documentElement
    if (root === null)
        return []
    const ids = new Set(elements(rels, "Relationship").map((r) => r.getAttribute("Id") ?? ""))
    const dropped: string[] = []
    for (const { el, attr } of relRefs(root).filter((ref) => ref.attr.value !== "" && !ids.has(ref.attr.value))) {
        if (!attached(el, root))
            continue
        const kept = relRefs(el).some((ref) => ref.el === el && ref.attr !== attr && ids.has(ref.attr.value))
        if (el === root || (kept && !REQUIRED.includes(el.nodeName))) {
            el.removeAttributeNode(attr)
            dropped.push(`the reference ${attr.name}="${attr.value}"`)
            continue
        }
        const scope = REQUIRED.includes(el.nodeName) ? null : SCOPES.map((rule) => rule(el)).find((s) => s !== null) ?? null
        if (scope !== null) {
            scope.doomed.parentNode?.removeChild(scope.doomed)
            dropped.push(scope.what)
            continue
        }
        const shape = closest(el, SHAPES)
        if (shape !== null) {
            dropped.push(`the shape '${firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? ""}'`)
            removeShape(shape)
        }
        else {
            dropped.push(`<${el.nodeName}>`)
            el.parentNode?.removeChild(el)
        }
    }
    return dropped
}

/**
 *  Make cNvPr shape ids unique within a slide (duplicates trigger repair);
 *  an `mc:Fallback` repeats the ids of its choice by design.
 *
 *  @param slide - the slide document
 *  @returns whether an id changed
 */
export const uniquifyShapeIds = (slide: Document): boolean => {
    const nextId = PartAllocator.shapeIds(slide)
    const seen = new Set<string>()
    let changed = false
    for (const el of elements(slide, "p:cNvPr").filter((e) => !inFallback(e))) {
        const id = el.getAttribute("id") ?? "0"
        if (seen.has(id)) {
            el.setAttribute("id", String(nextId()))
            changed = true
        }
        else
            seen.add(id)
    }
    return changed
}

/**
 *  Repair the relationships of a part: duplicate ids keep their first
 *  occurrence, dead internal targets go; rels only removed shapes used
 *  (`unused`, unless the XML still references them) and the drawings of
 *  removed SmartArt (`forced`) go as well.
 *
 *  @param pkg - the deck package
 *  @param part - the owning part
 *  @param doc - its DOM, null for a non-XML part
 *  @param rels - its relationships (edited in place)
 *  @param unused - ids only removed shapes used
 *  @param forced - ids to drop unless still referenced
 *  @returns the ids of the dropped relationships
 */
export const repairRels = (pkg: Package, part: string, doc: Document | null, rels: Document,
    unused: Set<string> = new Set(), forced: Set<string> = new Set()): string[] => {
    const xml = doc === null ? "" : serializeXml(doc)
    const seen = new Set<string>()
    const dropped: string[] = []
    for (const rel of pkg.targets(part, rels)) {
        const { el, id, target } = rel
        const dead = target !== null && !pkg.has(target)
        const stale = (forced.has(id) || (unused.has(id) && !relIs(rel, RELS.implicit))) && !xml.includes(`"${id}"`)
        if (seen.has(id) || dead || stale) {
            el.parentNode?.removeChild(el)
            dropped.push(id)
        }
        seen.add(id)
    }
    return dropped
}

/**
 *  Drop the references to missing relationships of a part and renumber
 *  duplicate shape ids of a slide.
 *
 *  @param doc - the part's DOM (edited in place)
 *  @param rels - its relationships
 *  @returns what was repaired, empty when nothing
 */
export const repairContent = (doc: Document, rels: Document): string[] => {
    const repaired = dropDangling(doc, rels).map((what) => `removed ${what}`)
    if (doc.documentElement?.nodeName === "p:sld" && uniquifyShapeIds(doc))
        repaired.push("renumbered duplicate shape ids")
    return repaired
}

/**
 *  Repair a part carrying repair triggers (engine/defects): duplicate and
 *  dead relationships go, what references a missing relationship goes,
 *  duplicate shape ids of a slide are renumbered.
 *
 *  @param pkg - the deck package
 *  @param part - the part to repair (a slide's notes or chart, a layout, a
 *                master, the presentation, "" for the package rels, ...)
 *  @returns what was repaired, empty when nothing
 */
export const repairPart = async (pkg: Package, part: string): Promise<string[]> => {
    const rels = await pkg.rels(part)
    const doc = /\.xml$/.test(part) ? await pkg.xml(part) : null
    const dropped = repairRels(pkg, part, doc, rels)
    const repaired = doc === null ? [] : repairContent(doc, rels)
    if (dropped.length > 0)
        pkg.writeRels(part, rels)
    if (doc !== null && repaired.length > 0)
        pkg.write(part, doc)
    return [ ...dropped.length > 0 ? [ `dropped relationships ${dropped.join(",")}` ] : [], ...repaired ]
}
