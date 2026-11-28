/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/slides: everything the in-place engine does to ONE slide.
**  A touched slide is parsed once, receives all of its planned edits
**  (removals, text, pictures, background, footer, notes, generated
**  elements, the stacking order, hyperlinks, visibility, repairs) on its
**  DOM and is serialized
**  once. New slides are created from the deck's own layout part.
*/

import { readFileSync } from "node:fs"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../../infra/errors.js"
import { inchToEmu, type Frame } from "../../core/model.js"
import type { PlannedEdit, PlannedElement, SlidePlanEntry } from "../../core/ops/registry.js"
import type { ZOrder } from "../../schema/payloads.js"
import { notesOwner } from "../defects.js"
import { EMPTY_RELS, PartAllocator, REL, relTarget, relsPartOf } from "../parts.js"
import { fitImage, imageSize, type Box, type ImageFit, type ImageSize } from "../image-size.js"
import { slideFromLayout } from "../seed.js"
import { digestBytes } from "../guard.js"
import { HLINK_ATTR, placeholderListStyle, setBodyText, setShapeText, type ListStyle } from "../text.js"
import { NS_A, NS_P, NS_R, colorElement, elements, firstElement, parseXml, relRefs, xmlEscape } from "../xml.js"
import type { Package } from "./package.js"
import { removeShape, repairContent, repairRels, type Repair } from "./repair.js"
import type { Generated } from "./transplant.js"
import { applyEdit, replaceExisting, replaceNotes, type ImageSwap } from "./edits.js"

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

/**  the shapes of a shape tree, in stacking order (back to front): every
     child but the group's own properties and extension list  */
const stacked = (spTree: Element): Element[] => {
    const out: Element[] = []
    for (let n = spTree.firstChild; n !== null; n = n.nextSibling)
        if (n.nodeType === 1 && ![ "p:nvGrpSpPr", "p:grpSpPr", "p:extLst" ].includes(n.nodeName))
            out.push(n as Element)
    return out
}

/**
 *  Move a top-level shape in the stacking order of its slide: to the
 *  front (before the tree's extension list), to the back (right after the
 *  tree's group properties), or one shape forward or backward.
 *
 *  @param shape - the shape (a child of `p:spTree`)
 *  @param zOrder - where it goes
 */
export const moveShape = (shape: Element, zOrder: ZOrder): void => {
    const spTree = shape.parentNode as Element | null
    if (spTree === null)
        return
    const order = stacked(spTree)
    const at = order.indexOf(shape)
    let before: Element | null
    switch (zOrder) {
        case "front":
            before = elements(spTree, "p:extLst").find((el) => el.parentNode === spTree) ?? null
            break
        case "back":
            before = order[0] ?? null
            break
        case "forward":
            if (at < 0 || at === order.length - 1)
                return
            before = (order[at + 1] as Element).nextSibling as Element | null
            break
        case "backward":
            if (at <= 0)
                return
            before = order[at - 1] as Element
            break
    }
    if (before !== shape)
        spTree.insertBefore(shape, before)
}

/**  set `show="0"` on the slide root for a hidden slide, clear it otherwise  */
const setSlideVisibility = (slide: Document, hidden: boolean): void => {
    const root = slide.documentElement
    if (root === null)
        return
    const current = root.getAttribute("show")
    if (hidden && current !== "0" && current !== "false")
        root.setAttribute("show", "0")
    else if (!hidden && root.hasAttribute("show"))
        root.removeAttribute("show")
}

/**  set a solid background color (hex or theme color) on a slide  */
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
    fill.appendChild(colorElement(slide, color))
    bgPr.appendChild(fill)
    bgPr.appendChild(slide.createElementNS(NS_A, "a:effectLst"))
    bg.appendChild(bgPr)
    cSld.insertBefore(bg, cSld.firstChild)
}

/**
 *  Give a slide relationships of its own for the references a content
 *  cloned from its layout carries (a picture fill, a hyperlink): the
 *  layout's ids mean nothing on the slide.
 *
 *  @param pkg - the deck package
 *  @param part - the slide part
 *  @param rels - the slide's relationships (edited in place)
 *  @param layoutPart - the layout part the content comes from
 *  @param scope - the cloned content (changed in place)
 *  @returns the image parts the slide now relates to
 */
const adoptLayoutRels = async (pkg: Package, part: string, rels: Document, layoutPart: string, scope: Element): Promise<string[]> => {
    const layoutRels = new Map(pkg.targets(layoutPart, await pkg.rels(layoutPart)).map((t) => [ t.id, t ]))
    const ids = new Map<string, string>()
    const images: string[] = []
    for (const { el, attr } of relRefs(scope)) {
        const rel = layoutRels.get(attr.value)
        if (rel !== undefined && !ids.has(attr.value)) {
            ids.set(attr.value, pkg.addRel(rels, rel.type,
                rel.target === null ? rel.el.getAttribute("Target") ?? "" : relTarget(part, rel.target),
                rel.target === null))
            if (rel.kind === "image" && rel.target !== null)
                images.push(rel.target)
        }
        const id = ids.get(attr.value)
        if (id !== undefined)
            el.setAttributeNS(NS_R, attr.name, id)
    }
    return images
}

/**  the footer-band placeholders: footer, slide number and date  */
type BandType = "ftr" | "sldNum" | "dt"

/**  the slide's own placeholders of a footer-band type  */
const bandShapes = (slide: Document, type: BandType): Element[] =>
    elements(slide, "p:sp").filter((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === type)

/**
 *  Set a footer-band placeholder of a slide: remove the slide's own, then
 *  (unless `content` is false) clone the layout's with its relationships --
 *  the footer with the given text, the slide number and date with their
 *  field -- in front of the shape tree's extension list. Without such a
 *  placeholder on the layout, the slide keeps its own (shown means
 *  present, never removed), a footer text goes into it.
 *
 *  @param pkg - the deck package
 *  @param part - the slide part
 *  @param slide - the slide DOM
 *  @param rels - the slide's relationships
 *  @param layoutPart - the slide's layout part, null when unknown
 *  @param type - the placeholder type
 *  @param content - the footer text, true to show the layout's content,
 *                   false to remove the placeholder
 *  @param lang - the language of the footer text
 *  @returns the image parts the slide now relates to (layout content), the
 *           image parts the removed placeholder showed, the relationship
 *           ids it referenced and the layout's copy (null when none came)
 */
const setBandPlaceholder = async (pkg: Package, part: string, slide: Document, rels: Document, layoutPart: string | null,
    type: BandType, content: string | boolean, lang: string): Promise<{ images: string[], removed: string[], refs: string[],
    clone: Element | null }> => {
    const refs: string[] = []
    const spTree = firstElement(slide, "p:spTree")
    if (spTree === null)
        return { images: [], removed: [], refs, clone: null }
    const source = content === false || layoutPart === null || !pkg.has(layoutPart) ? undefined
        : elements(await pkg.xml(layoutPart), "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === type)

    /*  none on the layout: the slide's own stays, with the footer text  */
    if (content !== false && source === undefined) {
        if (typeof content === "string")
            for (const txBody of bandShapes(slide, type).map((sp) => firstElement(sp, "p:txBody")))
                if (txBody !== null)
                    setBodyText(txBody, content, false, lang, false, true)
        return { images: [], removed: [], refs, clone: null }
    }
    for (const sp of bandShapes(slide, type)) {
        refs.push(...relRefs(sp).map((ref) => ref.attr.value))
        sp.parentNode?.removeChild(sp)
    }
    const removed = pkg.targets(part, rels).filter((t) => t.kind === "image" && t.target !== null && refs.includes(t.id))
        .map((t) => t.target as string)
    if (source === undefined || layoutPart === null)
        return { images: [], removed, refs, clone: null }
    const clone = slide.importNode(source, true) as Element
    firstElement(clone, "p:cNvPr")?.setAttribute("id", String(PartAllocator.shapeIds(slide)()))
    const txBody = firstElement(clone, "p:txBody")

    /*  the footer text as the deck's text; the layout's own content (the
        slide-number or date field, a footer text) takes the deck language  */
    if (typeof content === "string" && txBody !== null)
        setBodyText(txBody, content, false, lang, false, true)
    else
        for (const props of [ ...elements(clone, "a:rPr"), ...elements(clone, "a:endParaRPr") ])
            props.setAttribute("lang", lang)
    const images = await adoptLayoutRels(pkg, part, rels, layoutPart, clone)
    spTree.insertBefore(clone, elements(spTree, "p:extLst").find((el) => el.parentNode === spTree) ?? null)
    return { images, removed, refs, clone }
}

/**  the frame of a shape's own `a:xfrm` in EMU, null when it has none  */
const xfrmBox = (xfrm: Element | null): Box | null => {
    const off = xfrm === null ? null : firstElement(xfrm, "a:off")
    const ext = xfrm === null ? null : firstElement(xfrm, "a:ext")
    if (off === null || ext === null)
        return null
    const box = { x: Number(off.getAttribute("x")), y: Number(off.getAttribute("y")),
        w: Number(ext.getAttribute("cx")), h: Number(ext.getAttribute("cy")) }
    return Object.values(box).every(Number.isFinite) ? box : null
}

/**
 *  Replace an empty picture placeholder with a picture shape. The image
 *  takes the placeholder frame by its fit: cropped centered (`a:srcRect`,
 *  as PowerPoint does), contained in a centered, shrunk frame, or
 *  stretched; an unknown image size or frame stretches.
 *
 *  @param slide - the slide DOM
 *  @param rid - the image relationship id
 *  @param phIdx - the picture placeholder's idx
 *  @param frame - the layout's placeholder frame in inches, if known
 *  @param size - the image's pixel size, null when unknown
 *  @param fit - how the image takes the frame
 *  @returns false when the slide has no such empty placeholder
 */
const fillPicturePlaceholder = (slide: Document, rid: string, phIdx: number, frame: Frame | undefined,
    size: ImageSize | null, fit: ImageFit): boolean => {
    const sp = elements(slide, "p:sp").find((s) => {
        const ph = firstElement(s, "p:ph")
        return ph !== null && ph.getAttribute("type") === "pic" && Number(ph.getAttribute("idx") ?? "0") === phIdx
    })
    if (sp === undefined)
        return false
    const cNvPr = firstElement(sp, "p:cNvPr")
    const ph = firstElement(sp, "p:ph") as Element
    const xfrm = frame === undefined ? firstElement(sp, "a:xfrm") : null
    const box = frame !== undefined
        ? { x: inchToEmu(frame.x), y: inchToEmu(frame.y), w: inchToEmu(frame.w), h: inchToEmu(frame.h) }
        : xfrmBox(xfrm)
    const fitted = box === null ? null : fitImage(box, size, fit)

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
    blipFill.appendChild(blip)
    if (fitted !== null && fitted.crop !== null) {
        const srcRect = slide.createElementNS(NS_A, "a:srcRect")
        for (const [ side, value ] of Object.entries(fitted.crop))
            if (value !== 0)
                srcRect.setAttribute(side, String(value))
        blipFill.appendChild(srcRect)
    }
    const stretch = slide.createElementNS(NS_A, "a:stretch")
    stretch.appendChild(slide.createElementNS(NS_A, "a:fillRect"))
    blipFill.appendChild(stretch)
    pic.appendChild(blipFill)

    /*  the slide's own frame is kept as it is unless the fit shrinks it  */
    const spPr = slide.createElementNS(NS_P, "p:spPr")
    if (fitted !== null && (frame !== undefined || fitted.box !== box)) {
        const xf = slide.createElementNS(NS_A, "a:xfrm")
        const off = slide.createElementNS(NS_A, "a:off")
        off.setAttribute("x", String(fitted.box.x))
        off.setAttribute("y", String(fitted.box.y))
        const ext = slide.createElementNS(NS_A, "a:ext")
        ext.setAttribute("cx", String(fitted.box.w))
        ext.setAttribute("cy", String(fitted.box.h))
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

/**  the placeholders of a notes slide in PowerPoint's order: the slide
     image, the notes body and the slide number  */
const NOTES_PLACEHOLDERS = [
    { type: "sldImg", name: "Slide Image Placeholder", locks: "<a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/>",
        fallback: "<p:ph type=\"sldImg\"/>" },
    { type: "body", name: "Notes Placeholder", locks: "<a:spLocks noGrp=\"1\"/>", fallback: "<p:ph type=\"body\" idx=\"1\"/>" },
    { type: "sldNum", name: "Slide Number Placeholder", locks: "<a:spLocks noGrp=\"1\"/>", fallback: null }
] as const

/**  the shape of a placeholder type in a notes part or the notes master  */
const notesPlaceholder = (doc: Document, type: string): Element | null =>
    elements(doc, "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === type) ?? null

/**  a `p:ph` element as XML text, with the attributes of the given one  */
const phXml = (ph: Element): string => {
    let attrs = ""
    for (let i = 0; i < ph.attributes.length; i++) {
        const attr = ph.attributes.item(i)
        if (attr !== null && !attr.name.startsWith("xmlns"))
            attrs += ` ${attr.name}="${xmlEscape(attr.value, true)}"`
    }
    return `<p:ph${attrs}/>`
}

/**  the placeholder shapes of a new notes slide, as PowerPoint creates
     them: each placeholder the notes master has (by type, with its idx --
     an idx of another master placeholder would inherit that one's frame),
     with an empty notes body and the master's slide-number field  */
const notesShapes = (master: Document | null, lang: string): string => {
    let id = 2
    let xml = ""
    for (const spec of NOTES_PLACEHOLDERS) {
        const sp = master === null ? null : notesPlaceholder(master, spec.type)
        const ph = sp === null ? null : firstElement(sp, "p:ph")
        const field = spec.type === "sldNum" && sp !== null ? firstElement(sp, "a:fld") : null
        if ((ph === null && spec.fallback === null) || (spec.type === "sldNum" && field === null))
            continue
        let body = ""
        if (spec.type === "body")
            body = "<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody>"
        else if (field !== null)
            body = "<p:txBody><a:bodyPr/><a:lstStyle/><a:p>"
                + `<a:fld id="${xmlEscape(field.getAttribute("id") ?? "", true)}" type="slidenum">`
                + `<a:rPr lang="${xmlEscape(lang, true)}"/><a:t>\u2039#\u203A</a:t></a:fld></a:p></p:txBody>`
        xml += `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${spec.name} ${id - 1}"/><p:cNvSpPr>${spec.locks}</p:cNvSpPr>`
            + `<p:nvPr>${ph === null ? spec.fallback ?? "" : phXml(ph)}</p:nvPr></p:nvSpPr><p:spPr/>${body}</p:sp>`
        id++
    }
    return xml
}

/**  the notes body of a notes part, created from the notes master when
     the part has none  */
const notesBodyText = (notes: Document, master: Document | null, lang: string): Element => {
    let sp = notesPlaceholder(notes, "body")
    if (sp === null) {
        const spec = parseXml(`<p:spTree xmlns:a="${NS_A}" xmlns:p="${NS_P}">${notesShapes(master, lang)}</p:spTree>`)
        const spTree = firstElement(notes, "p:spTree")
        const body = notesPlaceholder(spec, "body") as Element
        const ids = PartAllocator.shapeIds(notes)
        firstElement(body, "p:cNvPr")?.setAttribute("id", String(ids()))
        sp = notes.importNode(body, true) as Element
        spTree?.insertBefore(sp, elements(spTree, "p:extLst").find((el) => el.parentNode === spTree) ?? null)
    }
    let txBody = firstElement(sp, "p:txBody")
    if (txBody === null) {
        txBody = notes.createElementNS(NS_P, "p:txBody")
        txBody.appendChild(notes.createElementNS(NS_A, "a:bodyPr"))
        txBody.appendChild(notes.createElementNS(NS_A, "a:lstStyle"))
        sp.appendChild(txBody)
    }
    return txBody
}

/**
 *  Write the speaker notes of a slide. An existing notes part keeps
 *  everything but the text of its notes body (the slide's own copy when
 *  it shares one); a new one is built from the notes master's
 *  placeholders, as PowerPoint creates it, and related to the slide and
 *  the notes master.
 *
 *  @param pkg - the deck package
 *  @param slidePart - the slide part
 *  @param slideRels - the slide's relationships (edited in place)
 *  @param text - the notes text
 *  @param notesMaster - the deck's notes master part
 *  @param lang - the language of the notes text
 *  @param shared - whether the slide's notes part belongs to another
 *                  slide or none
 */
const setNotes = async (pkg: Package, slidePart: string, slideRels: Document, text: string, notesMaster: string,
    lang: string, shared: boolean): Promise<void> => {
    const master = notesMaster !== "" && pkg.has(notesMaster) ? await pkg.xml(notesMaster) : null
    const current = pkg.targets(slidePart, slideRels).find((t) => t.kind === "notesSlide")?.target ?? null
    if (current !== null && pkg.has(current)) {
        const own = shared ? await unshareNotes(pkg, slidePart, slideRels) ?? current : current
        const notes = await pkg.xml(own)
        setBodyText(notesBodyText(notes, master, lang), text, false, lang, false, true)
        pkg.write(own, notes)
        return
    }
    const notesPart = pkg.alloc.part("ppt/notesSlides", "notesSlide", ".xml")
    const notes = parseXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + `<p:notes xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">`
        + "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        + `<p:grpSpPr/>${notesShapes(master, lang)}</p:spTree></p:cSld>`
        + "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>")
    setBodyText(notesBodyText(notes, master, lang), text, false, lang, false, true)
    pkg.write(notesPart, notes)
    pkg.write(relsPartOf(notesPart),
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + `<Relationship Id="rId1" Type="${REL}/notesMaster" Target="${xmlEscape(relTarget(notesPart, notesMaster), true)}"/>`
        + `<Relationship Id="rId2" Type="${REL}/slide" Target="../slides/${path.posix.basename(slidePart)}"/>`
        + "</Relationships>")
    await pkg.typeAs(notesPart, "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
    for (const rel of pkg.targets(slidePart, slideRels))
        if (rel.kind === "notesSlide")
            rel.el.parentNode?.removeChild(rel.el)
    pkg.addRel(slideRels, `${REL}/notesSlide`, `../notesSlides/${path.posix.basename(notesPart)}`)
}

/**  give a slide its own copy of its notes part (one that may belong to
     another slide, or to none), with the back-reference pointing at this
     slide; returns the copy, null when the slide has no notes  */
const unshareNotes = async (pkg: Package, slidePart: string, slideRels: Document): Promise<string | null> => {
    const rel = pkg.targets(slidePart, slideRels).find((t) => t.kind === "notesSlide")
    const notesPart = rel?.target ?? null
    if (rel === undefined || notesPart === null || !pkg.has(notesPart))
        return null
    const owner = notesOwner({ part: notesPart, doc: await pkg.xml(notesPart), rels: pkg.targets(notesPart, await pkg.rels(notesPart)) })
    const copy = await pkg.clonePart(pkg, notesPart, new Map([ [ owner ?? slidePart, slidePart ] ]))
    if (owner === null) {
        const copyRels = await pkg.rels(copy)
        pkg.addRel(copyRels, `${REL}/slide`, relTarget(copy, slidePart))
        pkg.writeRels(copy, copyRels)
    }
    rel.el.setAttribute("Target", relTarget(slidePart, copy))
    return copy
}

/**  replace the `pptc-hlink` markers of the text writer, the el.add
     finishing and el.set by relationships: an external hyperlink for a
     URL, a slide relationship and a slide jump for a slide link  */
const wireHyperlinks = (pkg: Package, part: string, slide: Document, slideRels: Document, links: Map<string, string>): void => {
    for (const hlink of elements(slide, "a:hlinkClick")) {
        const target = hlink.getAttribute(HLINK_ATTR)
        if (target === null)
            continue
        hlink.removeAttribute(HLINK_ATTR)
        let rid: string
        if (target.startsWith("slide:")) {
            const to = links.get(target)
            if (to === undefined)
                throw new PptcError("E_ENGINE", `slide link '${target}' has no resolved target`)
            rid = pkg.addRel(slideRels, `${REL}/slide`, relTarget(part, to))
            hlink.setAttribute("action", "ppaction://hlinksldjump")
        }
        else
            rid = pkg.addRel(slideRels, `${REL}/hyperlink`, target, true)
        hlink.setAttributeNS(NS_R, "r:id", rid)
    }
}

/**  drop links to removed slides (hyperlink jumps, zoom) with their
     relationship; returns whether anything was dropped  */
export const dropLinksToRemoved = (pkg: Package, owner: string, doc: Document | null, rels: Document, removed: Set<string>): boolean => {
    let changed = false
    for (const { el: rel, id, kind, target } of pkg.targets(owner, rels)) {
        if (kind !== "slide" || target === null || !removed.has(target))
            continue
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

/**
 *  Create a new, empty slide from a layout of the deck.
 *
 *  @param pkg - the deck package
 *  @param layoutPart - the layout part to base the slide on
 *  @returns the new slide part
 */
export const createSlide = async (pkg: Package, layoutPart: string): Promise<string> => {
    const part = pkg.alloc.part("ppt/slides", "slide", ".xml")
    const slide = parseXml(slideFromLayout(await pkg.text(layoutPart)))
    const rels = parseXml(EMPTY_RELS)
    pkg.addRel(rels, `${REL}/slideLayout`, relTarget(part, layoutPart))

    /*  cloned placeholders may reference layout relationships (a picture
        fill, a hyperlink): the slide gets its own relationship to the same
        target, as the layout's ids mean nothing on the slide  */
    if (slide.documentElement !== null)
        await adoptLayoutRels(pkg, part, rels, layoutPart, slide.documentElement)
    pkg.write(part, slide)
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
    /**  whether the slide's notes part belongs to another slide or none  */
    sharedNotes: boolean
    /**  the deck's notes master part  */
    notesMaster: string
    /**  generated elements of this apply, null when there are none  */
    generated: Generated | null
    /**  the language of text that states none (the deck language)  */
    lang: string
    /**  output slide part of every slide link of the plan  */
    links: Map<string, string>
}

/**  what one slide pass did beyond its repairs: the images el.set
     replaced, the images layout content brought in, the replacements of
     each text.replace in the slide's existing content, whether a
     text.replace removed a run that showed an image (a picture-filled run),
     and how many replacements it made in the slide's existing notes  */
export interface SlideEdited extends Repair {
    swaps: ImageSwap[]
    adopted: ImageSwap[]
    replaced: { op: number, count: number }[]
    runImages: boolean
    notesReplaced: number
}

/**
 *  Apply every planned edit of one slide in a single parse/serialize pass.
 *
 *  @param pkg - the deck package
 *  @param work - the slide and its plan entry
 *  @returns what was repaired on the slide (input defects), the images
 *           replaced or adopted and the text replacements
 *  @throws PptcError E_ADDR_NOTFOUND when an addressed shape is missing
 */
export const editSlide = async (pkg: Package, work: SlideWork): Promise<SlideEdited> => {
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
    const doomed = new Set<Element>()
    for (const name of entry.removeNames) {
        const shape = elements(slide, "p:cNvPr").map(shapeOf).find((s) =>
            !doomed.has(s) && isTopLevel(s) && firstElement(s, "p:cNvPr")?.getAttribute("name") === name)
        if (shape !== undefined)
            doomed.add(shape)
    }

    /*  text.replace on the content the slide carried before the apply,
        before any other op changes it (earlier ops' changes are left out)  */
    const unused = new Set<string>()
    const replaced: { op: number, count: number }[] = []
    const images = new Set(pkg.targets(part, rels).filter((t) => t.kind === "image").map((t) => t.id))
    let runImages = false
    for (const replace of entry.replaces) {
        const done = replaceExisting(slide, replace)
        replaced.push({ op: replace.op, count: done.count })
        for (const id of done.refs) {
            unused.add(id)
            runImages ||= images.has(id)
        }
    }

    /*  removals, then text; relationships only removed shapes or replaced
        text used (hyperlinks) are dropped below  */
    for (const shape of [ ...doomed, ...fills.filter((f) => !f.fill.append).map((f) => f.shape) ])
        for (const ref of relRefs(shape))
            unused.add(ref.attr.value)

    /*  a removed SmartArt's drawing is referenced from its data part  */
    const forced = new Set<string>()
    const targetOf = new Map(pkg.targets(part, rels).map((t) => [ t.id, t.target ]))
    for (const shape of doomed)
        for (const relIds of elements(shape, "dgm:relIds")) {
            const data = targetOf.get(relIds.getAttribute("r:dm") ?? "") ?? null
            const relId = data !== null && pkg.has(data)
                ? firstElement(await pkg.xml(data), "dsp:dataModelExt")?.getAttribute("relId") ?? undefined : undefined
            if (relId !== undefined)
                forced.add(relId)
        }
    for (const shape of doomed)
        removeShape(shape)

    /*  the list style a placeholder inherits (bullets by level)  */
    const layoutPart = pkg.targets(part, rels).find((t) => t.kind === "slideLayout")?.target ?? null
    const layoutDoc = layoutPart !== null && pkg.has(layoutPart) ? await pkg.xml(layoutPart) : null
    const masterPart = layoutPart === null || layoutDoc === null ? null
        : pkg.targets(layoutPart, await pkg.rels(layoutPart)).find((t) => t.kind === "slideMaster")?.target ?? null
    const masterDoc = masterPart !== null && pkg.has(masterPart) ? await pkg.xml(masterPart) : null
    const listStyle = (shape: Element): ListStyle => placeholderListStyle(shape, layoutDoc, masterDoc)
    for (const { fill, shape } of fills)
        setShapeText(shape, fill.text as NonNullable<typeof fill.text>, fill.append, work.lang, false, listStyle(shape))

    /*  pictures into picture placeholders  */
    for (const fill of entry.fills.filter((f) => f.image !== undefined)) {
        const file = fill.image as string
        const ext = path.extname(file).slice(1).toLowerCase() || "png"
        const media = pkg.alloc.part("ppt/media", "image", `.${ext}`)
        const bytes = readFileSync(file)
        pkg.write(media, bytes)
        await pkg.typePart(media, "", media)
        const rid = pkg.addRel(rels, `${REL}/image`, `../media/${path.posix.basename(media)}`)
        if (!fillPicturePlaceholder(slide, rid, fill.phIdx, fill.frame, imageSize(bytes), fill.fit ?? "crop"))
            throw missing(`empty picture placeholder idx ${fill.phIdx}`)
    }

    /*  the footer-band placeholders set anew: their edits follow the
        layout's copy that replaces them (below)  */
    const bandSet = [ [ "ftr", entry.footer ], [ "sldNum", entry.slideNumber ], [ "dt", entry.date ] ] as const
    const banded = new Map<string, BandType>()
    for (const [ type, content ] of bandSet)
        for (const sp of content === null ? [] : bandShapes(slide, type))
            banded.set(firstElement(sp, "p:cNvPr")?.getAttribute("name") ?? "", type)
    const clones = new Map<BandType, Element>()
    /*  a shape by name; a footer-band placeholder the layout's copy
        replaced is that copy (one the slide kept is found by its name)  */
    const bandShape = (name: string): Element | null => {
        const band = banded.get(name)
        const clone = band === undefined ? undefined : clones.get(band)
        return clone ?? findShape(slide, name, isTopLevel)
    }

    /*  el.set changes of existing elements, in op order; what the changed
        elements referenced before may become unused (links, images)  */
    const swaps: ImageSwap[] = []
    const applyEdits = async (edits: PlannedEdit[]): Promise<void> => {
        for (const edit of edits) {
            const shape = need(bandShape(edit.name), `element '${edit.name}'`)
            for (const ref of relRefs(shape))
                unused.add(ref.attr.value)
            /*  every swap counts (the content guard checks them as a
                multiset, an image swapped again included)  */
            const swap = await applyEdit(pkg, part, rels, shape, edit, work.lang, listStyle(shape))
            if (swap !== null)
                swaps.push(swap)
        }
    }
    await applyEdits(entry.edits.filter((e) => !banded.has(e.name)))

    /*  background, footer, slide number, date, notes  */
    if (entry.background !== null)
        setBackground(slide, entry.background)
    const bands: { images: string[], removed: string[] }[] = []
    for (const [ type, content ] of bandSet)
        if (content !== null) {
            const band = await setBandPlaceholder(pkg, part, slide, rels, layoutPart, type, content, work.lang)
            bands.push(band)
            if (band.clone !== null)
                clones.set(type, band.clone)
            for (const id of band.refs)
                unused.add(id)
        }
    await applyEdits(entry.edits.filter((e) => banded.has(e.name)))
    const notesReplaces = entry.notes === null ? entry.replaces.filter((r) => r.notes) : []
    let notesReplaced = 0
    if (entry.notes !== null)
        await setNotes(pkg, part, rels, entry.notes, work.notesMaster, work.lang, work.sharedNotes)
    else if (work.sharedNotes || notesReplaces.length > 0) {
        const copy = await unshareNotes(pkg, part, rels)
        if (copy !== null && notesReplaces.length > 0) {
            const notes = await pkg.xml(copy)
            for (const replace of notesReplaces) {
                const count = replaceNotes(notes, replace)
                replaced.push({ op: replace.op, count })
                notesReplaced += count
            }
            pkg.write(copy, notes)
        }
    }

    /*  generated elements, the stacking order, hyperlinks, visibility,
        links, repairs  */
    const placed = work.generated === null ? new Map<PlannedElement, Element>()
        : await work.generated.transplant(pkg, entry, slide, rels, part)

    /*  a generated element enters the shape tree in front when its el.add
        comes, so a move counts only the shapes there at that point  */
    const spTree = firstElement(slide, "p:spTree")
    const enter = (shape: Element): void => {
        if (shape.parentNode === null && spTree !== null)
            spTree.insertBefore(shape, elements(spTree, "p:extLst").find((el) => el.parentNode === spTree) ?? null)
    }
    for (const z of entry.zOrder) {
        const shape = typeof z.target === "string" ? bandShape(z.target) : placed.get(z.target) ?? null
        if (shape === null)
            continue
        if (z.add) {
            enter(shape)
            moveShape(shape, "front")
        }
        moveShape(shape, z.zOrder)
    }
    for (const shape of placed.values())
        enter(shape)
    wireHyperlinks(pkg, part, slide, rels, work.links)
    setSlideVisibility(slide, entry.hidden)
    dropLinksToRemoved(pkg, part, slide, rels, work.removed)
    const dropped = repairRels(pkg, part, slide, rels, unused, forced).filter((id) => !unused.has(id) && !forced.has(id))
    const repaired = repairContent(pkg, part, slide, rels)
    if (dropped.length > 0)
        repaired.actions.unshift(`dropped relationships ${dropped.join(",")}`)
    pkg.write(part, slide)
    pkg.writeRels(part, rels)
    /*  a band placeholder's images: the layout's come, the removed
        placeholder's may go  */
    const digestOf = async (image: string): Promise<string[]> => {
        const file = pkg.zip.file(image)
        return file === null ? [] : [ digestBytes(await file.async("nodebuffer")) ]
    }
    const adopted: ImageSwap[] = []
    for (const band of bands) {
        const from: string[] = []
        for (const image of band.removed)
            from.push(...await digestOf(image))
        const to: string[] = []
        for (const image of band.images)
            to.push(...await digestOf(image))
        if (to.length === 0 && from.length > 0)
            adopted.push({ from, to: null })
        to.forEach((hash, i) => adopted.push({ from: i === 0 ? from : [], to: hash }))
    }
    return { ...repaired, swaps, adopted, replaced, runImages, notesReplaced }
}
