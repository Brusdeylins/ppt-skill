/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/slides: everything the in-place engine does to ONE slide.
**  A touched slide is parsed once, receives all of its planned edits
**  (removals, text, pictures, background, footer, notes, generated
**  elements, hyperlinks, visibility, repairs) on its DOM and is serialized
**  once. New slides are created from the deck's own layout part.
*/

import { readFileSync } from "node:fs"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../../infra/errors.js"
import { inchToEmu, type Frame } from "../../core/model.js"
import type { SlidePlanEntry } from "../../core/ops/registry.js"
import { IMPLICIT_RELS } from "../guard.js"
import { PartAllocator, relsPartOf } from "../parts.js"
import { slideFromLayout } from "../seed.js"
import { HLINK_ATTR, setShapeText } from "../text.js"
import { NS_A, NS_P, NS_R, elements, firstElement, parseXml, serializeXml } from "../xml.js"
import { Package } from "./package.js"
import type { Generated } from "./transplant.js"

/**  OfficeDocument relationship type base URI  */
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/**  relationship types a slide carries without referencing them by id in
     its XML (layout, notes, comments, theme override, ...; the SmartArt
     drawing is referenced from the SmartArt data part)  */
const UNREFERENCED_RELS = [ "/slideLayout", "/notesSlide", "/diagramDrawing", ...IMPLICIT_RELS ]

/**  is the shape the placeholder with the given idx  */
const isPlaceholder = (idx: number) => (shape: Element): boolean => {
    const ph = firstElement(shape, "p:ph")
    return ph !== null && Number(ph.getAttribute("idx") ?? "0") === idx
}

/**  is the shape a top-level shape of the slide (what el.* addresses)  */
const isTopLevel = (shape: Element): boolean =>
    (shape.parentNode as Element | null)?.nodeName === "p:spTree"

/**  the shape owning a `p:cNvPr` (its grandparent)  */
const shapeOf = (cNvPr: Element): Element => cNvPr.parentNode?.parentNode as Element

/**  the first shape of a name that `pick` accepts -- names may repeat, so a
     textbox named like a placeholder must never be edited instead of it  */
const findShape = (slide: Document, name: string, pick: (shape: Element) => boolean): Element | null =>
    elements(slide, "p:cNvPr").map(shapeOf).find((shape) =>
        firstElement(shape, "p:cNvPr")?.getAttribute("name") === name && pick(shape)) ?? null

/**  every relationship id an element subtree references  */
const relIdsOf = (scope: Element): string[] => {
    const ids: string[] = []
    const walk = (el: Element): void => {
        for (let i = 0; i < el.attributes.length; i++) {
            const attr = el.attributes.item(i)
            if (attr !== null && attr.namespaceURI === NS_R)
                ids.push(attr.value)
        }
        for (let n = el.firstChild; n !== null; n = n.nextSibling)
            if (n.nodeType === 1)
                walk(n as Element)
    }
    walk(scope)
    return ids
}

/**  remove a shape; inside an `mc:AlternateContent` (OLE objects,
     SmartArt) the whole wrapper goes, so no fallback copy stays behind  */
const removeShape = (shape: Element): void => {
    let doomed = shape
    for (let up = shape.parentNode as Element | null; up !== null && up.nodeName !== "p:spTree"; up = up.parentNode as Element | null)
        if (up.nodeName === "mc:AlternateContent")
            doomed = up
    doomed.parentNode?.removeChild(doomed)
}

/**  set `show="0"` on the slide root for a hidden slide, clear it otherwise  */
const setSlideVisibility = (slide: Document, hidden: boolean): void => {
    const root = slide.documentElement
    if (root === null)
        return
    if (hidden)
        root.setAttribute("show", "0")
    else if (root.hasAttribute("show"))
        root.removeAttribute("show")
}

/**  set a solid background color on a slide  */
const setBackground = (slide: Document, color: string): void => {
    const cSld = firstElement(slide, "p:cSld")
    if (cSld === null)
        return
    const old = firstElement(slide, "p:bg")
    if (old !== null)
        cSld.removeChild(old)
    const bg = slide.createElementNS(NS_P, "p:bg")
    const bgPr = slide.createElementNS(NS_P, "p:bgPr")
    const fill = slide.createElementNS(NS_A, "a:solidFill")
    const clr = slide.createElementNS(NS_A, "a:srgbClr")
    clr.setAttribute("val", color)
    fill.appendChild(clr)
    bgPr.appendChild(fill)
    bgPr.appendChild(slide.createElementNS(NS_A, "a:effectLst"))
    bg.appendChild(bgPr)
    cSld.insertBefore(bg, cSld.firstChild)
}

/**  clone the footer placeholder from the slide's layout and set its text  */
const setFooter = (slide: Document, layout: Document | null, text: string, nextId: () => number): void => {
    const spTree = firstElement(slide, "p:spTree")
    if (spTree === null)
        return
    for (const sp of elements(slide, "p:sp"))
        if (firstElement(sp, "p:ph")?.getAttribute("type") === "ftr")
            sp.parentNode?.removeChild(sp)
    const ftr = layout === null ? undefined
        : elements(layout, "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === "ftr")
    if (ftr === undefined)
        return
    const clone = slide.importNode(ftr, true) as Element
    firstElement(clone, "p:cNvPr")?.setAttribute("id", String(nextId()))
    const txBody = firstElement(clone, "p:txBody")
    if (txBody !== null) {
        for (const p of elements(txBody, "a:p"))
            txBody.removeChild(p)
        const p = slide.createElementNS(NS_A, "a:p")
        const r = slide.createElementNS(NS_A, "a:r")
        const t = slide.createElementNS(NS_A, "a:t")
        t.appendChild(slide.createTextNode(text))
        r.appendChild(t)
        p.appendChild(r)
        txBody.appendChild(p)
    }
    spTree.appendChild(clone)
}

/**  replace an empty picture placeholder with a picture shape  */
const fillPicturePlaceholder = (slide: Document, rid: string, phIdx: number, frame: Frame | undefined): boolean => {
    const sp = elements(slide, "p:sp").find((s) => {
        const ph = firstElement(s, "p:ph")
        return ph !== null && ph.getAttribute("type") === "pic" && Number(ph.getAttribute("idx") ?? "0") === phIdx
    })
    if (sp === undefined)
        return false
    const cNvPr = firstElement(sp, "p:cNvPr")
    const ph = firstElement(sp, "p:ph") as Element
    const xfrm = frame === undefined ? firstElement(sp, "a:xfrm") : null

    const pic = slide.createElementNS(NS_P, "p:pic")
    const nvPicPr = slide.createElementNS(NS_P, "p:nvPicPr")
    const newCNvPr = slide.createElementNS(NS_P, "p:cNvPr")
    newCNvPr.setAttribute("id", cNvPr?.getAttribute("id") ?? "99")
    newCNvPr.setAttribute("name", cNvPr?.getAttribute("name") ?? `Picture ${phIdx}`)
    const cNvPicPr = slide.createElementNS(NS_P, "p:cNvPicPr")
    const picLocks = slide.createElementNS(NS_A, "a:picLocks")
    picLocks.setAttribute("noChangeAspect", "1")
    cNvPicPr.appendChild(picLocks)
    const nvPr = slide.createElementNS(NS_P, "p:nvPr")
    nvPr.appendChild(slide.importNode(ph, true))
    nvPicPr.appendChild(newCNvPr)
    nvPicPr.appendChild(cNvPicPr)
    nvPicPr.appendChild(nvPr)
    pic.appendChild(nvPicPr)

    const blipFill = slide.createElementNS(NS_P, "p:blipFill")
    const blip = slide.createElementNS(NS_A, "a:blip")
    blip.setAttributeNS(NS_R, "r:embed", rid)
    const stretch = slide.createElementNS(NS_A, "a:stretch")
    stretch.appendChild(slide.createElementNS(NS_A, "a:fillRect"))
    blipFill.appendChild(blip)
    blipFill.appendChild(stretch)
    pic.appendChild(blipFill)

    const spPr = slide.createElementNS(NS_P, "p:spPr")
    if (frame !== undefined) {
        const xf = slide.createElementNS(NS_A, "a:xfrm")
        const off = slide.createElementNS(NS_A, "a:off")
        off.setAttribute("x", String(inchToEmu(frame.x)))
        off.setAttribute("y", String(inchToEmu(frame.y)))
        const ext = slide.createElementNS(NS_A, "a:ext")
        ext.setAttribute("cx", String(inchToEmu(frame.w)))
        ext.setAttribute("cy", String(inchToEmu(frame.h)))
        xf.appendChild(off)
        xf.appendChild(ext)
        spPr.appendChild(xf)
    }
    else if (xfrm !== null)
        spPr.appendChild(slide.importNode(xfrm, true))
    pic.appendChild(spPr)

    sp.parentNode?.replaceChild(pic, sp)
    return true
}

/**  escape text for an XML text node  */
const xmlEscape = (s: string): string =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")

/**  write speaker notes into a fresh notesSlide part and point the slide at
     it (the previous notes part is collected by the package sweep)  */
const setNotes = async (pkg: Package, slidePart: string, slideRels: Document, text: string, notesMaster: string): Promise<void> => {
    const notesPart = pkg.alloc.part("ppt/notesSlides", "notesSlide", ".xml")
    const paras = text.split("\n").map((line) => `<a:p><a:r><a:t>${xmlEscape(line)}</a:t></a:r></a:p>`).join("")
    pkg.write(notesPart,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<p:notes xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
        + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        + " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
        + "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        + "<p:grpSpPr/>"
        + "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Notes Placeholder\"/>"
        + "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>"
        + "<p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr></p:nvSpPr>"
        + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>`
        + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>")
    pkg.write(relsPartOf(notesPart),
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + `<Relationship Id="rId1" Type="${REL}/notesMaster" Target="${path.posix.relative("ppt/notesSlides", notesMaster)}"/>`
        + `<Relationship Id="rId2" Type="${REL}/slide" Target="../slides/${path.posix.basename(slidePart)}"/>`
        + "</Relationships>")
    await pkg.typeAs(notesPart, "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
    for (const rel of elements(slideRels, "Relationship"))
        if ((rel.getAttribute("Type") ?? "").endsWith("/notesSlide"))
            rel.parentNode?.removeChild(rel)
    pkg.addRel(slideRels, `${REL}/notesSlide`, `../notesSlides/${path.posix.basename(notesPart)}`)
}

/**  give a slide its own copy of a notes part another slide uses, with the
     back-reference pointing at this slide  */
const unshareNotes = async (pkg: Package, slidePart: string, slideRels: Document): Promise<void> => {
    const rel = elements(slideRels, "Relationship").find((r) => (r.getAttribute("Type") ?? "").endsWith("/notesSlide"))
    const notesPart = rel === undefined ? null : Package.targets(slidePart, slideRels).find((t) => t.rel === rel)?.target ?? null
    if (rel === undefined || notesPart === null || !pkg.has(notesPart))
        return
    const notesXml = await pkg.text(notesPart)
    const backref = Package.targets(notesPart, await pkg.rels(notesPart)).find((t) =>
        t.type.endsWith("/slide") && !notesXml.includes(`"${t.rel.getAttribute("Id") ?? ""}"`))?.target ?? slidePart
    const copy = await pkg.clonePart(pkg, notesPart, new Map([ [ backref, slidePart ] ]))
    rel.setAttribute("Target", `../notesSlides/${path.posix.basename(copy)}`)
}

/**  replace the `pptc-hlink` markers of the text writer by external
     hyperlink relationships  */
const wireHyperlinks = (pkg: Package, slide: Document, slideRels: Document): void => {
    for (const hlink of elements(slide, "a:hlinkClick")) {
        const url = hlink.getAttribute(HLINK_ATTR)
        if (url === null)
            continue
        const rid = pkg.addRel(slideRels, `${REL}/hyperlink`, url, true)
        hlink.removeAttribute(HLINK_ATTR)
        hlink.setAttributeNS(NS_R, "r:id", rid)
    }
}

/**  drop links to removed slides (hyperlink jumps, zoom) with their
     relationship; returns whether anything was dropped  */
export const dropLinksToRemoved = (owner: string, doc: Document | null, rels: Document, removed: Set<string>): boolean => {
    let changed = false
    for (const { rel, type, target } of Package.targets(owner, rels)) {
        if (!/\/relationships\/slide$/.test(type) || target === null || !removed.has(target))
            continue
        const id = rel.getAttribute("Id") ?? ""
        rel.parentNode?.removeChild(rel)
        changed = true
        if (doc !== null)
            for (const tag of [ "a:hlinkClick", "a:hlinkHover", "p:sld" ])
                for (const el of elements(doc, tag))
                    if (el.getAttribute("r:id") === id)
                        el.parentNode?.removeChild(el)
    }
    return changed
}

/**  make cNvPr shape ids unique within a slide (duplicates trigger repair)  */
const uniquifyShapeIds = (slide: Document): void => {
    const nextId = PartAllocator.shapeIds(slide)
    const seen = new Set<string>()
    for (const el of elements(slide, "p:cNvPr")) {
        const id = el.getAttribute("id") ?? "0"
        if (seen.has(id))
            el.setAttribute("id", String(nextId()))
        else
            seen.add(id)
    }
}

/**  repair the relationships of a slide: duplicate ids keep their first
     occurrence, dead internal targets go; rels only removed shapes used are
     dropped as well  */
const repairRels = (pkg: Package, part: string, slide: Document, rels: Document, unused: Set<string>): void => {
    const xml = serializeXml(slide)
    const seen = new Set<string>()
    for (const { rel, type, target } of Package.targets(part, rels)) {
        const id = rel.getAttribute("Id") ?? ""
        const dead = target !== null && !pkg.has(target)
        const stale = unused.has(id) && !UNREFERENCED_RELS.some((k) => type.endsWith(k)) && !xml.includes(`"${id}"`)
        if (seen.has(id) || dead || stale)
            rel.parentNode?.removeChild(rel)
        seen.add(id)
    }
}

/**
 *  Create a new, empty slide from a layout of the deck.
 *
 *  @param pkg - the deck package
 *  @param layoutPart - the layout part to base the slide on
 *  @returns the new slide part
 */
export const createSlide = async (pkg: Package, layoutPart: string): Promise<string> => {
    const part = pkg.alloc.part("ppt/slides", "slide", ".xml")
    pkg.write(part, slideFromLayout(await pkg.text(layoutPart)))
    const rels = parseXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")
    pkg.addRel(rels, `${REL}/slideLayout`, path.posix.relative("ppt/slides", layoutPart))
    pkg.writeRels(part, rels)
    await pkg.typeAs(part, "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
    return part
}

/**  what one touched output slide needs beyond its plan entry  */
export interface SlideWork {
    /**  the output slide part  */
    part: string
    /**  the plan entry  */
    entry: SlidePlanEntry
    /**  input slide parts removed by this apply  */
    removed: Set<string>
    /**  whether the slide's notes part is shared with an earlier slide  */
    sharedNotes: boolean
    /**  the deck's notes master part  */
    notesMaster: string
    /**  generated elements of this apply, null when there are none  */
    generated: Generated | null
}

/**
 *  Apply every planned edit of one slide in a single parse/serialize pass.
 *
 *  @param pkg - the deck package
 *  @param work - the slide and its plan entry
 *  @throws PptcError E_ADDR_NOTFOUND when an addressed shape is missing
 */
export const editSlide = async (pkg: Package, work: SlideWork): Promise<void> => {
    const { part, entry } = work
    const slide = await pkg.xml(part)
    const rels = await pkg.rels(part)
    const missing = (what: string): PptcError =>
        new PptcError("E_ADDR_NOTFOUND", `slide ${part}: no ${what}`)
    const need = (shape: Element | null, what: string): Element => {
        if (shape === null)
            throw missing(what)
        return shape
    }

    /*  resolve every addressed shape on the unchanged slide first  */
    const fills = entry.fills.filter((f) => f.text !== undefined).map((f) => ({
        fill: f, shape: need(findShape(slide, f.phName, isPlaceholder(f.phIdx)), `placeholder '${f.phName}'`)
    }))
    const texts = entry.setTexts.map((t) => ({
        text: t.text, shape: need(findShape(slide, t.name, isTopLevel), `element '${t.name}'`)
    }))
    const doomed = new Set<Element>()
    for (const name of entry.removeNames) {
        const shape = elements(slide, "p:cNvPr").map(shapeOf).find((s) =>
            !doomed.has(s) && isTopLevel(s) && firstElement(s, "p:cNvPr")?.getAttribute("name") === name)
        if (shape !== undefined)
            doomed.add(shape)
    }

    /*  removals, then text; relationships only removed shapes or replaced
        text used (hyperlinks) are dropped below  */
    const unused = new Set<string>()
    for (const shape of [ ...doomed, ...fills.filter((f) => !f.fill.append).map((f) => f.shape), ...texts.map((t) => t.shape) ])
        for (const id of relIdsOf(shape))
            unused.add(id)
    for (const shape of doomed)
        removeShape(shape)
    for (const { fill, shape } of fills)
        setShapeText(shape, fill.text as NonNullable<typeof fill.text>, fill.append)
    for (const { text, shape } of texts)
        setShapeText(shape, text)

    /*  pictures into picture placeholders  */
    for (const fill of entry.fills.filter((f) => f.image !== undefined)) {
        const file = fill.image as string
        const ext = path.extname(file).slice(1).toLowerCase() || "png"
        const media = pkg.alloc.part("ppt/media", "image", `.${ext}`)
        pkg.write(media, readFileSync(file))
        await pkg.typePart(media, "", media)
        const rid = pkg.addRel(rels, `${REL}/image`, `../media/${path.posix.basename(media)}`)
        if (!fillPicturePlaceholder(slide, rid, fill.phIdx, fill.frame))
            throw missing(`empty picture placeholder idx ${fill.phIdx}`)
    }

    /*  background, footer, notes  */
    if (entry.background !== null)
        setBackground(slide, entry.background)
    if (entry.footer !== null) {
        const layout = Package.targets(part, rels).find((t) => t.type.endsWith("/slideLayout"))?.target ?? null
        setFooter(slide, layout !== null && pkg.has(layout) ? await pkg.xml(layout) : null, entry.footer, PartAllocator.shapeIds(slide))
    }
    if (entry.notes !== null)
        await setNotes(pkg, part, rels, entry.notes, work.notesMaster)
    else if (work.sharedNotes)
        await unshareNotes(pkg, part, rels)

    /*  generated elements, hyperlinks, visibility, links, repairs  */
    if (work.generated !== null)
        await work.generated.transplant(pkg, entry, slide, rels, part)
    wireHyperlinks(pkg, slide, rels)
    setSlideVisibility(slide, entry.hidden)
    dropLinksToRemoved(part, slide, rels, work.removed)
    uniquifyShapeIds(slide)
    repairRels(pkg, part, slide, rels, unused)
    pkg.write(part, slide)
    pkg.writeRels(part, rels)
}
