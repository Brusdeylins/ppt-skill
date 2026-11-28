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
import { notesOwner } from "../defects.js"
import { EMPTY_RELS, PartAllocator, REL, relTarget, relsPartOf } from "../parts.js"
import { fitImage, imageSize, type Box, type ImageFit, type ImageSize } from "../image-size.js"
import { slideFromLayout } from "../seed.js"
import { digestBytes } from "../guard.js"
import { HLINK_ATTR, setShapeText } from "../text.js"
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

/**
 *  Set a footer-band placeholder of a slide: remove the slide's own, then
 *  (unless `content` is false) clone the layout's with its relationships --
 *  the footer with the given text, the slide number and date with their
 *  field -- in front of the shape tree's extension list.
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
 *  @returns the image parts the slide now relates to (layout content) and
 *           the relationship ids the removed placeholder referenced
 */
const setBandPlaceholder = async (pkg: Package, part: string, slide: Document, rels: Document, layoutPart: string | null,
    type: BandType, content: string | boolean, lang: string): Promise<{ images: string[], refs: string[] }> => {
    const refs: string[] = []
    const spTree = firstElement(slide, "p:spTree")
    if (spTree === null)
        return { images: [], refs }
    for (const sp of elements(slide, "p:sp"))
        if (firstElement(sp, "p:ph")?.getAttribute("type") === type) {
            refs.push(...relRefs(sp).map((ref) => ref.attr.value))
            sp.parentNode?.removeChild(sp)
        }
    if (content === false || layoutPart === null || !pkg.has(layoutPart))
        return { images: [], refs }
    const source = elements(await pkg.xml(layoutPart), "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === type)
    if (source === undefined)
        return { images: [], refs }
    const clone = slide.importNode(source, true) as Element
    firstElement(clone, "p:cNvPr")?.setAttribute("id", String(PartAllocator.shapeIds(slide)()))
    const txBody = firstElement(clone, "p:txBody")
    if (typeof content === "string" && txBody !== null) {
        for (const p of elements(txBody, "a:p"))
            txBody.removeChild(p)
        const p = slide.createElementNS(NS_A, "a:p")
        const r = slide.createElementNS(NS_A, "a:r")
        const rPr = slide.createElementNS(NS_A, "a:rPr")
        rPr.setAttribute("lang", lang)
        r.appendChild(rPr)
        const t = slide.createElementNS(NS_A, "a:t")
        t.appendChild(slide.createTextNode(content))
        r.appendChild(t)
        p.appendChild(r)
        txBody.appendChild(p)
    }
    const images = await adoptLayoutRels(pkg, part, rels, layoutPart, clone)
    spTree.insertBefore(clone, elements(spTree, "p:extLst").find((el) => el.parentNode === spTree) ?? null)
    return { images, refs }
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

/**  write speaker notes into a fresh notesSlide part and point the slide at
     it (the previous notes part is collected by the package sweep)  */
const setNotes = async (pkg: Package, slidePart: string, slideRels: Document, text: string, notesMaster: string,
    lang: string): Promise<void> => {
    const notesPart = pkg.alloc.part("ppt/notesSlides", "notesSlide", ".xml")
    const paras = text.split("\n").map((line) =>
        `<a:p><a:r><a:rPr lang="${xmlEscape(lang, true)}"/><a:t>${xmlEscape(line)}</a:t></a:r></a:p>`).join("")
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
     each text.replace in the slide's existing content, and whether a
     text.replace removed a run that showed an image (a picture-filled run)  */
export interface SlideEdited extends Repair {
    swaps: ImageSwap[]
    adopted: string[]
    replaced: { op: number, count: number }[]
    runImages: boolean
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
    for (const { fill, shape } of fills)
        setShapeText(shape, fill.text as NonNullable<typeof fill.text>, fill.append, work.lang)

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

    /*  el.set changes of existing elements, in op order; what the changed
        elements referenced before may become unused (links, images)  */
    const swaps: ImageSwap[] = []
    for (const edit of entry.edits) {
        const shape = need(findShape(slide, edit.name, isTopLevel), `element '${edit.name}'`)
        for (const ref of relRefs(shape))
            unused.add(ref.attr.value)
        const swap = await applyEdit(pkg, part, rels, shape, edit, work.lang)

        /*  an image swapped again in the same apply: only the last stays  */
        const again = swap === null ? undefined : swaps.find((s) => swap.from.includes(s.to))
        if (swap !== null && again !== undefined) {
            again.from.push(...swap.from.filter((f) => f !== again.to))
            again.to = swap.to
        }
        else if (swap !== null)
            swaps.push(swap)
    }

    /*  background, footer, slide number, date, notes  */
    if (entry.background !== null)
        setBackground(slide, entry.background)
    const layout = pkg.targets(part, rels).find((t) => t.kind === "slideLayout")?.target ?? null
    const adoptedParts: string[] = []
    for (const [ type, content ] of [ [ "ftr", entry.footer ], [ "sldNum", entry.slideNumber ], [ "dt", entry.date ] ] as const)
        if (content !== null) {
            const band = await setBandPlaceholder(pkg, part, slide, rels, layout, type, content, work.lang)
            adoptedParts.push(...band.images)
            for (const id of band.refs)
                unused.add(id)
        }
    const notesReplaces = entry.notes === null ? entry.replaces.filter((r) => r.notes) : []
    if (entry.notes !== null)
        await setNotes(pkg, part, rels, entry.notes, work.notesMaster, work.lang)
    else if (work.sharedNotes || notesReplaces.length > 0) {
        const copy = await unshareNotes(pkg, part, rels)
        if (copy !== null && notesReplaces.length > 0) {
            const notes = await pkg.xml(copy)
            for (const replace of notesReplaces)
                replaced.push({ op: replace.op, count: replaceNotes(notes, replace) })
            pkg.write(copy, notes)
        }
    }

    /*  generated elements, hyperlinks, visibility, links, repairs  */
    if (work.generated !== null)
        await work.generated.transplant(pkg, entry, slide, rels, part)
    wireHyperlinks(pkg, part, slide, rels, work.links)
    setSlideVisibility(slide, entry.hidden)
    dropLinksToRemoved(pkg, part, slide, rels, work.removed)
    const dropped = repairRels(pkg, part, slide, rels, unused, forced).filter((id) => !unused.has(id) && !forced.has(id))
    const repaired = repairContent(pkg, part, slide, rels)
    if (dropped.length > 0)
        repaired.actions.unshift(`dropped relationships ${dropped.join(",")}`)
    pkg.write(part, slide)
    pkg.writeRels(part, rels)
    const adopted: string[] = []
    for (const image of adoptedParts) {
        const file = pkg.zip.file(image)
        if (file !== null)
            adopted.push(digestBytes(await file.async("nodebuffer")))
    }
    return { ...repaired, swaps, adopted, replaced, runImages }
}
