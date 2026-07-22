/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/post: the zip-level post-pass that runs after pptx-automizer has
**  written the slide structure. It garbage-collects unreferenced slide parts
**  and applies everything automizer cannot express: speaker notes, footer
**  cloning, backgrounds, slide visibility, images inside picture
**  placeholders, hyperlink relationships and document properties.
**
**  LOCKSTEP: every repair-trigger this pass fixes is asserted read-only by
**  engine/verify.ts -- when a trigger is added or changed here, teach the
**  verifier the same rule (and vice versa), or self-verify turns green
**  fixes into false E_INTEGRITY failures.
*/

import JSZip from "jszip"
import { readFileSync } from "node:fs"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../infra/errors.js"
import { inchToEmu, type Frame } from "../core/model.js"
import { HLINK_ATTR } from "./text.js"
import { partText, cleanContentTypes } from "./parts.js"
import { NS_A, NS_P, NS_R, NS_REL, elements, firstElement, parseXml, serializeXml } from "./xml.js"

/**  one captured p:custDataLst child: the element kind plus the slide
     relationship it referenced in the source deck  */
export interface SlideTagRef {
    /**  element kind: "p:custData" or "p:tags"  */
    kind: string
    /**  relationship type of the source rel  */
    relType: string
    /**  relationship target, relative to ppt/slides/  */
    target: string
}

/**  per-output-slide work items, in final slide order  */
export interface PostSlideWork {
    /**  speaker notes text to set, null = leave untouched  */
    notes: string | null
    /**  footer text to clone in from the layout, null = leave untouched  */
    footer: string | null
    /**  solid background color (RRGGBB), null = leave untouched  */
    background: string | null
    /**  whether the output slide must be hidden ("Hide Slide", `show="0"` on
         the `p:sld` root). automizer drops this attribute on every rebuild, so
         it is re-applied here from the plan to keep hidden slides hidden  */
    hidden: boolean
    /**  images to insert into picture placeholders  */
    images: { phIdx: number, path: string, frame?: Frame }[]
    /**  slide-tags references (p:custDataLst) captured from the source
         slide, to restore after the automizer rebuild; empty = none  */
    slideTags: SlideTagRef[]
}

/**  mutable archive context shared by the post steps  */
interface Post {
    zip: JSZip
    rid: number
    /**  tags parts already wired to an output slide in this pass  */
    tagsSeen: Set<string>
}

/**  slide part paths referenced from the presentation, in sldIdLst order  */
const referencedSlides = async (zip: JSZip): Promise<string[]> => {
    const pres = parseXml(await partText(zip, "ppt/presentation.xml"))
    const rels = parseXml(await partText(zip, "ppt/_rels/presentation.xml.rels"))
    const relMap = new Map(elements(rels, "Relationship")
        .map((r) => [r.getAttribute("Id") ?? "", r.getAttribute("Target") ?? ""]))
    return elements(pres, "p:sldId").map((s) =>
        path.posix.normalize(path.posix.join("ppt", relMap.get(s.getAttribute("r:id") ?? "") ?? "")))
}

/**  remove a list of parts plus their rels and content-type overrides  */
const removeParts = async (zip: JSZip, parts: string[]): Promise<void> => {
    if (parts.length === 0)
        return
    let ct = await partText(zip, "[Content_Types].xml")
    for (const part of parts) {
        zip.remove(part)
        zip.remove(`${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`)
        ct = ct.replace(new RegExp(`<Override PartName="/${part.replace(/[.\\/]/g, "\\$&")}"[^>]*/>`, "g"), "")
    }
    zip.file("[Content_Types].xml", ct)
}

/**  garbage-collect slide and notes parts no longer referenced  */
const gcParts = async (zip: JSZip, kept: string[]): Promise<void> => {
    const keptSet = new Set(kept)
    const orphanSlides = Object.keys(zip.files)
        .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f) && !keptSet.has(f))
    await removeParts(zip, orphanSlides)
    /*  notes slides referenced from kept slides survive; their slide
        back-reference must follow the (possibly renamed) parent part  */
    const referencedNotes = new Set<string>()
    for (const slide of kept) {
        const relsPart = `ppt/slides/_rels/${path.posix.basename(slide)}.rels`
        const relsText = await zip.file(relsPart)?.async("string")
        if (relsText === undefined)
            continue
        for (const rel of elements(parseXml(relsText), "Relationship"))
            if ((rel.getAttribute("Type") ?? "").endsWith("/notesSlide")) {
                const notesPart = path.posix.normalize(
                    path.posix.join("ppt/slides", rel.getAttribute("Target") ?? ""))
                referencedNotes.add(notesPart)
                const noteRelsPart = `ppt/notesSlides/_rels/${path.posix.basename(notesPart)}.rels`
                const noteRelsText = await zip.file(noteRelsPart)?.async("string")
                if (noteRelsText === undefined)
                    continue
                const noteRels = parseXml(noteRelsText)
                for (const back of elements(noteRels, "Relationship"))
                    if ((back.getAttribute("Type") ?? "").endsWith("/slide"))
                        back.setAttribute("Target", `../slides/${path.posix.basename(slide)}`)
                zip.file(noteRelsPart, serializeXml(noteRels))
            }
    }
    const orphanNotes = Object.keys(zip.files)
        .filter((f) => /^ppt\/notesSlides\/notesSlide[^/]*\.xml$/.test(f) && !referencedNotes.has(f))
    await removeParts(zip, orphanNotes)
    /*  drop presentation relationships that are stale: either the target
        part vanished, or it is a /slide rel that no <p:sldId> points at.
        automizer mints a fresh "-created" slide rel on every re-import but
        only wires ONE into the sldIdLst; the rest pile up pointing at parts
        that still exist (so the vanished-target check misses them) until the
        thousand-strong rels file makes PowerPoint demand a repair.  */
    const presRelsPart = "ppt/_rels/presentation.xml.rels"
    const presRels = parseXml(await partText(zip, presRelsPart))
    const pres = parseXml(await partText(zip, "ppt/presentation.xml"))
    const liveSlideRids = new Set(elements(pres, "p:sldId")
        .map((s) => s.getAttribute("r:id") ?? ""))
    let pruned = false
    for (const rel of elements(presRels, "Relationship")) {
        if (rel.getAttribute("TargetMode") === "External")
            continue
        const target = path.posix.normalize(path.posix.join("ppt", rel.getAttribute("Target") ?? ""))
        const isSlide = (rel.getAttribute("Type") ?? "").endsWith("/slide")
        if (zip.file(target) === null
            || (isSlide && !liveSlideRids.has(rel.getAttribute("Id") ?? ""))) {
            rel.parentNode?.removeChild(rel)
            pruned = true
        }
    }
    if (pruned)
        zip.file(presRelsPart, serializeXml(presRels))
}

/**  drop section slide references whose sldId vanished (user-created
     sections survive edits; stale refs would point at removed slides)  */
const pruneSectionRefs = async (zip: JSZip): Promise<void> => {
    const part = "ppt/presentation.xml"
    const xml = await partText(zip, part)
    if (!xml.includes("<p14:sectionLst"))
        return
    const ids = new Set([...xml.matchAll(/<p:sldId [^>]*\bid="(\d+)"/g)].map((m) => m[1] as string))
    const out = xml.replace(/<p14:sldId id="(\d+)"[^>]*\/>/g,
        (m, id: string) => ids.has(id) ? m : "")
    if (out !== xml)
        zip.file(part, out)
}

/**  rel types that are never referenced from the slide XML itself  */
const STRUCTURAL_RELS = ["/slideLayout", "/notesSlide"]

/**  drop slide relationships whose id no longer occurs in the slide XML
     (automizer accumulates one stale chart/image rel per re-import)  */
const pruneUnusedSlideRels = (slide: Document, slideRels: Document): boolean => {
    const xml = serializeXml(slide)
    let changed = false
    const seen = new Set<string>()
    for (const rel of elements(slideRels, "Relationship")) {
        const id = rel.getAttribute("Id") ?? ""
        /*  duplicate ids (left by the pre-0.2.10 rid counter) heal here  */
        if (seen.has(id)) {
            rel.parentNode?.removeChild(rel)
            changed = true
            continue
        }
        seen.add(id)
        const type = rel.getAttribute("Type") ?? ""
        if (STRUCTURAL_RELS.some((k) => type.endsWith(k)))
            continue
        if (!xml.includes(`"${id}"`)) {
            rel.parentNode?.removeChild(rel)
            changed = true
        }
    }
    return changed
}

/**  GC charts, embeddings and media unreachable from any live part --
     automizer re-imports duplicate chart parts on every apply, so without
     this the archive doubles its charts per round  */
const gcAssets = async (zip: JSZip): Promise<void> => {
    const isAsset = (f: string): boolean =>
        /^ppt\/(charts|embeddings|media)\//.test(f) && !f.endsWith("/") && !f.includes("/_rels/")
    const assets = Object.keys(zip.files).filter(isAsset)
    if (assets.length === 0)
        return
    const live = new Set<string>()
    const queue = Object.keys(zip.files).filter((f) =>
        f.endsWith(".rels") && !/^ppt\/(charts|embeddings|media)\//.test(f))
    while (queue.length > 0) {
        const relsPart = queue.pop() as string
        const relsText = await zip.file(relsPart)?.async("string")
        if (relsText === undefined)
            continue
        const base = path.posix.dirname(path.posix.dirname(relsPart))
        for (const rel of elements(parseXml(relsText), "Relationship")) {
            if (rel.getAttribute("TargetMode") === "External")
                continue
            const target = path.posix.normalize(path.posix.join(base, rel.getAttribute("Target") ?? ""))
            if (isAsset(target) && !live.has(target)) {
                live.add(target)
                const tRels = `${path.posix.dirname(target)}/_rels/${path.posix.basename(target)}.rels`
                if (zip.file(tRels) !== null)
                    queue.push(tRels)
            }
        }
    }
    await removeParts(zip, assets.filter((a) => !live.has(a)))
}

/**  make cNvPr shape ids unique within a slide (duplicates trigger repair)  */
const uniquifyShapeIds = (slide: Document): boolean => {
    const all = elements(slide, "p:cNvPr")
    let max = 0
    for (const el of all)
        max = Math.max(max, Number(el.getAttribute("id") ?? "0"))
    const seen = new Set<number>()
    let changed = false
    for (const el of all) {
        const id = Number(el.getAttribute("id") ?? "0")
        if (seen.has(id)) {
            el.setAttribute("id", String(++max))
            changed = true
        }
        else
            seen.add(id)
    }
    return changed
}

/**  next unique pptc relationship id within a rels document  */
const nextRid = (post: Post): string => `rIdPptc${post.rid++}`

/**  append a relationship to a rels document  */
const addRel = (rels: Document, id: string, type: string, target: string, external = false): void => {
    const rel = rels.createElementNS(NS_REL, "Relationship")
    rel.setAttribute("Id", id)
    rel.setAttribute("Type", type)
    rel.setAttribute("Target", target)
    if (external)
        rel.setAttribute("TargetMode", "External")
    rels.documentElement?.appendChild(rel)
}

/**  apply the slide-hidden flag: set `show="0"` on the `p:sld` root when the
     slide must be hidden, clear it otherwise. automizer rebuilds every slide
     and never copies this attribute, so a kept hidden slide would re-appear
     without this. Returns whether the slide XML changed.  */
const setSlideVisibility = (slide: Document, hidden: boolean): boolean => {
    const root = slide.documentElement
    if (root === null)
        return false
    const current = root.getAttribute("show")
    if (hidden) {
        if (current === "0")
            return false
        root.setAttribute("show", "0")
        return true
    }
    if (current === null || current === "")
        return false
    root.removeAttribute("show")
    return true
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
const setFooter = async (zip: JSZip, slidePart: string, slide: Document, text: string): Promise<void> => {
    /*  drop an existing footer shape first  */
    const spTree = firstElement(slide, "p:spTree")
    if (spTree === null)
        return
    for (const sp of elements(slide, "p:sp")) {
        const ph = firstElement(sp, "p:ph")
        if (ph !== null && ph.getAttribute("type") === "ftr")
            spTree.removeChild(sp)
    }
    /*  locate the layout and its footer placeholder  */
    const relsText = await partText(zip, `ppt/slides/_rels/${path.posix.basename(slidePart)}.rels`)
    const layoutTarget = elements(parseXml(relsText), "Relationship")
        .find((r) => (r.getAttribute("Type") ?? "").endsWith("/slideLayout"))?.getAttribute("Target")
    if (layoutTarget === undefined || layoutTarget === null)
        return
    const layoutPart = path.posix.normalize(path.posix.join("ppt/slides", layoutTarget))
    const layout = parseXml(await partText(zip, layoutPart))
    const ftr = elements(layout, "p:sp").find((sp) =>
        firstElement(sp, "p:ph")?.getAttribute("type") === "ftr")
    if (ftr === undefined)
        return
    const clone = slide.importNode(ftr, true) as Element
    /*  unique shape id and the footer text  */
    const ids = elements(slide, "p:cNvPr").map((c) => Number(c.getAttribute("id") ?? "0"))
    firstElement(clone, "p:cNvPr")?.setAttribute("id", String(Math.max(0, ...ids) + 1))
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

/**  content-type defaults for image media  */
const IMAGE_TYPES: Record<string, string> = {
    png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
    gif: "image/gif", svg: "image/svg+xml", webp: "image/webp"
}

/**  ensure a `<Default Extension=...>` entry exists for an image extension  */
const ensureImageDefault = async (zip: JSZip, ext: string): Promise<void> => {
    const ct = await partText(zip, "[Content_Types].xml")
    if (ct.includes(`Extension="${ext}"`))
        return
    zip.file("[Content_Types].xml",
        ct.replace("<Default", `<Default Extension="${ext}" ContentType="${IMAGE_TYPES[ext] ?? "image/png"}"/><Default`))
}

/**  replace an empty picture placeholder with an actual picture shape  */
const fillPicturePlaceholder = (
    slide: Document, rid: string, phIdx: number, frame: Frame | undefined
): boolean => {
    const sp = elements(slide, "p:sp").find((s) => {
        const ph = firstElement(s, "p:ph")
        return ph !== null && ph.getAttribute("type") === "pic"
            && Number(ph.getAttribute("idx") ?? "0") === phIdx
    })
    if (sp === undefined)
        return false
    const cNvPr = firstElement(sp, "p:cNvPr")
    const name = cNvPr?.getAttribute("name") ?? `Picture ${phIdx}`
    const id = cNvPr?.getAttribute("id") ?? "99"
    const ph = firstElement(sp, "p:ph") as Element
    const xfrm = frame === undefined ? firstElement(sp, "a:xfrm") : null

    const pic = slide.createElementNS(NS_P, "p:pic")
    const nvPicPr = slide.createElementNS(NS_P, "p:nvPicPr")
    const newCNvPr = slide.createElementNS(NS_P, "p:cNvPr")
    newCNvPr.setAttribute("id", id)
    newCNvPr.setAttribute("name", name)
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

/**  ensure a notesSlide content-type override exists for a part  */
const ensureNotesOverride = async (zip: JSZip, notesPart: string): Promise<void> => {
    const ct = await partText(zip, "[Content_Types].xml")
    if (!ct.includes(notesPart))
        zip.file("[Content_Types].xml", ct.replace("</Types>",
            `<Override PartName="/${notesPart}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>`))
}

/**  write speaker notes: create or replace the slide's notesSlide part  */
const setNotes = async (post: Post, slidePart: string, slideRels: Document, text: string): Promise<void> => {
    const base = path.posix.basename(slidePart, ".xml")
    const notesPart = `ppt/notesSlides/notesSlide-pptc-${base}.xml`
    const escaped = text.split("\n").map((line) =>
        `<a:p><a:r><a:t>${xmlEscape(line)}</a:t></a:r></a:p>`).join("")
    post.zip.file(notesPart,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<p:notes xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
        + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        + " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
        + "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        + "<p:grpSpPr/>"
        + "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Notes Placeholder\"/>"
        + "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>"
        + "<p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr></p:nvSpPr>"
        + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${escaped}</p:txBody></p:sp>`
        + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>")
    post.zip.file(`ppt/notesSlides/_rels/${path.posix.basename(notesPart)}.rels`,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster\" Target=\"../notesMasters/notesMaster1.xml\"/>"
        + `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${path.posix.basename(slidePart)}"/>`
        + "</Relationships>")
    await ensureNotesOverride(post.zip, notesPart)
    /*  drop an existing notes rel, then point the slide at the new part  */
    for (const rel of elements(slideRels, "Relationship"))
        if ((rel.getAttribute("Type") ?? "").endsWith("/notesSlide"))
            rel.parentNode?.removeChild(rel)
    addRel(slideRels, nextRid(post),
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
        `../notesSlides/${path.posix.basename(notesPart)}`)
}

/**  wire hyperlink marker attributes to real external relationships  */
const wireHyperlinks = (post: Post, slide: Document, slideRels: Document): void => {
    for (const hlink of elements(slide, "a:hlinkClick")) {
        const url = hlink.getAttribute(HLINK_ATTR)
        if (url === null)
            continue
        const rid = nextRid(post)
        addRel(slideRels, rid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", url, true)
        hlink.removeAttribute(HLINK_ATTR)
        hlink.setAttributeNS(NS_R, "r:id", rid)
    }
}

/**  ensure every slide owns its own notesSlide. The pptc notes part name is
     derived from the slide part basename, but pptx-automizer renumbers slide
     parts across applies, so a stale notes rel can survive on a slide whose
     basename no longer matches -- leaving two slides pointing at one notes
     part that carries a single back-reference. PowerPoint repairs that. Any
     shared pptc notes part is cloned so each referencing slide gets its own,
     with a back-reference to it.  */
const dedupeSharedNotes = async (zip: JSZip, slides: string[]): Promise<void> => {
    /*  notes part -> referencing slide parts, in slide order  */
    const refs = new Map<string, string[]>()
    for (const slidePart of slides) {
        const relsText = await zip.file(`ppt/slides/_rels/${path.posix.basename(slidePart)}.rels`)?.async("string")
        if (relsText === undefined)
            continue
        const rel = elements(parseXml(relsText), "Relationship").find((r) =>
            (r.getAttribute("Type") ?? "").endsWith("/notesSlide"))
        if (rel === undefined)
            continue
        const notesPart = path.posix.normalize(path.posix.join("ppt/slides", rel.getAttribute("Target") ?? ""))
        refs.set(notesPart, [...(refs.get(notesPart) ?? []), slidePart])
    }
    const setBackref = async (notesPart: string, slidePart: string): Promise<void> => {
        const relsPart = `ppt/notesSlides/_rels/${path.posix.basename(notesPart)}.rels`
        const doc = parseXml(await partText(zip, relsPart))
        for (const back of elements(doc, "Relationship"))
            if ((back.getAttribute("Type") ?? "").endsWith("/slide"))
                back.setAttribute("Target", `../slides/${path.posix.basename(slidePart)}`)
        zip.file(relsPart, serializeXml(doc))
    }
    for (const [notesPart, owners] of refs) {
        if (owners.length < 2 || !/notesSlide-pptc-/.test(notesPart))
            continue
        /*  the first owner keeps the original part; the rest get private clones  */
        await setBackref(notesPart, owners[0] as string)
        for (let i = 1; i < owners.length; i++) {
            const slidePart = owners[i] as string
            let clone = `ppt/notesSlides/notesSlide-pptc-${path.posix.basename(slidePart, ".xml")}.xml`
            if (clone === notesPart || zip.file(clone) !== null)
                clone = `ppt/notesSlides/notesSlide-pptc-${path.posix.basename(slidePart, ".xml")}-${i}.xml`
            zip.file(clone, await partText(zip, notesPart))
            zip.file(`ppt/notesSlides/_rels/${path.posix.basename(clone)}.rels`,
                await partText(zip, `ppt/notesSlides/_rels/${path.posix.basename(notesPart)}.rels`))
            await setBackref(clone, slidePart)
            const slideRelsPart = `ppt/slides/_rels/${path.posix.basename(slidePart)}.rels`
            const slideRels = parseXml(await partText(zip, slideRelsPart))
            for (const r of elements(slideRels, "Relationship"))
                if ((r.getAttribute("Type") ?? "").endsWith("/notesSlide"))
                    r.setAttribute("Target", `../notesSlides/${path.posix.basename(clone)}`)
            zip.file(slideRelsPart, serializeXml(slideRels))
            await ensureNotesOverride(zip, clone)
        }
    }
}

/**  restore the slide-tags references (p:custDataLst) captured from the
     source slide. pptx-automizer strips the element and its /tags
     relationships as "unsupported" on every re-import, which would silently
     unlink think-cell charts (and other add-in metadata) deck-wide on every
     apply. The list is rebuilt after p:spTree (its CT_CommonSlideData slot)
     with freshly minted relationship ids; references whose target part
     vanished from the package are dropped, so no dead rel can trip the
     self-verify. The tags part keeps its content-type coverage: automizer
     copies [Content_Types].xml from the root template, so the original
     override survives, and a clone carries the original's override along.  */
const restoreSlideTags = async (post: Post, slide: Document, slideRels: Document, refs: SlideTagRef[]): Promise<void> => {
    const cSld = firstElement(slide, "p:cSld")
    const spTree = firstElement(slide, "p:spTree")
    if (cSld === null || spTree === null)
        return
    const old = firstElement(slide, "p:custDataLst")
    if (old !== null)
        old.parentNode?.removeChild(old)
    const lst = slide.createElementNS(NS_P, "p:custDataLst")
    for (const ref of refs) {
        let target = ref.target
        let part = path.posix.normalize(path.posix.join("ppt/slides", target))
        if (post.zip.file(part) === null)
            continue
        /*  a source slide copied within one apply (slide.copy) would leave
            both output slides referencing ONE tags part -- PowerPoint mints
            an independent part per duplicated slide (think-cell tags are
            per-instance), so every re-use gets a private clone: the tags
            sibling of dedupeSharedNotes  */
        if (post.tagsSeen.has(part)) {
            const ext = path.posix.extname(part)
            const stem = part.slice(0, part.length - ext.length)
            let n = 2
            let clone = `${stem}-pptc-${n}${ext}`
            while (post.zip.file(clone) !== null)
                clone = `${stem}-pptc-${++n}${ext}`
            post.zip.file(clone, await (post.zip.file(part) as JSZip.JSZipObject).async("uint8array"))
            const ct = await partText(post.zip, "[Content_Types].xml")
            const override = new RegExp(`<Override PartName="/${part.replace(/[.\\/]/g, "\\$&")}"[^>]*/>`).exec(ct)
            if (override !== null)
                post.zip.file("[Content_Types].xml",
                    ct.replace("</Types>", `${override[0].replace(part, clone)}</Types>`))
            target = path.posix.join(path.posix.dirname(target), path.posix.basename(clone))
            part = clone
        }
        post.tagsSeen.add(part)
        const rid = nextRid(post)
        addRel(slideRels, rid, ref.relType, target)
        const el = slide.createElementNS(NS_P, ref.kind)
        el.setAttributeNS(NS_R, "r:id", rid)
        lst.appendChild(el)
    }
    if (lst.childNodes.length > 0)
        cSld.insertBefore(lst, spTree.nextSibling)
}

/**  read a slide's title text (title/ctrTitle placeholder), "" if none  */
const slideTitle = (slide: Document): string => {
    for (const sp of elements(slide, "p:sp")) {
        const type = firstElement(sp, "p:ph")?.getAttribute("type") ?? ""
        if (type === "title" || type === "ctrTitle")
            return elements(sp, "a:t").map((n) => n.textContent ?? "").join("")
    }
    return ""
}

/**  escape text for an XML text node  */
const xmlEscape = (s: string): string =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")

/**  keep docProps/app.xml in sync with the real slides. automizer writes
     this extended-properties part once and does NOT grow its slide list on
     later applies, so a deck that gained slides ends up declaring fewer
     slides than it has -- which makes PowerPoint demand a repair on open.
     The leading TitlesOfParts entries (fonts, theme, OLE servers) and every
     HeadingPairs group except the last (the slide-title group) are kept; the
     slide count, the slide-title HeadingPairs count and the title list are
     rewritten from the actual slides.  */
const syncAppProperties = async (zip: JSZip, slides: string[]): Promise<void> => {
    const part = "docProps/app.xml"
    const f = zip.file(part)
    if (f === null)
        return
    let xml = await f.async("string")
    const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(xml)
    const tp = /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(xml)
    if (hp === null || tp === null)
        return
    const hpBody = hp[1] ?? ""
    const tpBody = tp[1] ?? ""
    const n = slides.length
    const titles: string[] = []
    for (const slidePart of slides)
        titles.push(slideTitle(parseXml(await partText(zip, slidePart))))
    /*  the last HeadingPairs group is the slide titles; everything before it
        (fonts/theme/OLE) sums to the count of leading TitlesOfParts entries  */
    const counts = [...hpBody.matchAll(/<vt:i4>(\d+)<\/vt:i4>/g)].map((m) => Number(m[1]))
    const leading = counts.slice(0, -1).reduce((a, b) => a + b, 0)
    let seen = 0
    const newHp = hpBody.replace(/<vt:i4>\d+<\/vt:i4>/g, (m) =>
        ++seen === counts.length ? `<vt:i4>${n}</vt:i4>` : m)
    const kept = [...tpBody.matchAll(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g)]
        .map((m) => m[1] as string).slice(0, leading)
    const lpstr = [...kept, ...titles.map(xmlEscape)].map((e) => `<vt:lpstr>${e}</vt:lpstr>`).join("")
    xml = xml
        .replace(/<Slides>\d+<\/Slides>/, `<Slides>${n}</Slides>`)
        .replace(hp[0], `<HeadingPairs>${newHp}</HeadingPairs>`)
        .replace(tp[0], `<TitlesOfParts><vt:vector size="${leading + n}" baseType="lpstr">${lpstr}</vt:vector></TitlesOfParts>`)
    zip.file(part, xml)
}

/**  update document core properties  */
const setProps = async (zip: JSZip, props: Record<string, string>): Promise<void> => {
    const part = "docProps/core.xml"
    const text = await zip.file(part)?.async("string")
    if (text === undefined)
        return
    const doc = parseXml(text)
    const tags: Record<string, [string, string]> = {
        title: ["dc:title", "http://purl.org/dc/elements/1.1/"],
        author: ["dc:creator", "http://purl.org/dc/elements/1.1/"],
        subject: ["dc:subject", "http://purl.org/dc/elements/1.1/"],
        keywords: ["cp:keywords", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"],
        category: ["cp:category", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"],
        comments: ["dc:description", "http://purl.org/dc/elements/1.1/"]
    }
    for (const [key, value] of Object.entries(props)) {
        const spec = tags[key]
        if (spec === undefined)
            continue
        let el = firstElement(doc, spec[0])
        if (el === null) {
            el = doc.createElementNS(spec[1], spec[0])
            doc.documentElement?.appendChild(el)
        }
        while (el.firstChild !== null)
            el.removeChild(el.firstChild)
        el.appendChild(doc.createTextNode(value))
    }
    zip.file(part, serializeXml(doc))
}

/**  package-level identifiers for the custom document-properties part  */
const CUSTOM_PROPS_PART = "docProps/custom.xml"
const CUSTOM_PROPS_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
const CUSTOM_PROPS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
const NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
const NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
/**  the well-known fmtid every custom document property carries  */
const CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

/**  create or patch docProps/custom.xml with a name→value patch, ensuring
     the content-type override and the package relationship exist. Existing
     custom properties not named in the patch are preserved; PowerPoint keeps
     these across edit/save, so they make the deck self-describing.  */
const setCustomProps = async (zip: JSZip, patch: Record<string, string>): Promise<void> => {
    /*  merge: existing properties first, then the patch overrides/adds  */
    const props = new Map<string, string>()
    const existing = await zip.file(CUSTOM_PROPS_PART)?.async("string")
    if (existing !== undefined)
        for (const p of elements(parseXml(existing), "property")) {
            const vt = firstElement(p, "vt:lpwstr")
            props.set(p.getAttribute("name") ?? "", vt?.textContent ?? "")
        }
    for (const [k, v] of Object.entries(patch))
        props.set(k, v)

    /*  pids are sequential and MUST start at 2 (pid 1 is reserved)  */
    let pid = 2
    let body = ""
    for (const [name, value] of props)
        body += `<property fmtid="${CUSTOM_FMTID}" pid="${pid++}" name="${xmlEscape(name)}">`
            + `<vt:lpwstr>${xmlEscape(value)}</vt:lpwstr></property>`
    zip.file(CUSTOM_PROPS_PART,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
        + `<Properties xmlns="${NS_CUSTOM}" xmlns:vt="${NS_VT}">${body}</Properties>`)

    /*  content-type override (idempotent)  */
    const ct = await partText(zip, "[Content_Types].xml")
    if (!ct.includes(`PartName="/${CUSTOM_PROPS_PART}"`))
        zip.file("[Content_Types].xml", ct.replace("</Types>",
            `<Override PartName="/${CUSTOM_PROPS_PART}" ContentType="${CUSTOM_PROPS_CT}"/></Types>`))

    /*  package relationship from the root .rels (idempotent)  */
    const rootRels = parseXml(await partText(zip, "_rels/.rels"))
    const present = elements(rootRels, "Relationship")
        .some((r) => (r.getAttribute("Type") ?? "") === CUSTOM_PROPS_REL)
    if (!present) {
        const ids = new Set(elements(rootRels, "Relationship").map((r) => r.getAttribute("Id")))
        let n = 1
        while (ids.has(`rId${n}`))
            n++
        addRel(rootRels, `rId${n}`, CUSTOM_PROPS_REL, CUSTOM_PROPS_PART)
        zip.file("_rels/.rels", serializeXml(rootRels))
    }
}

/**
 *  Run the complete post-pass over automizer's output.
 *
 *  @param bytes - the .pptx produced by the automizer pass
 *  @param work - per-slide work items in final slide order
 *  @param props - document core-property patch, null for none
 *  @param customProps - custom document-property patch, null for none
 *  @returns the final .pptx bytes
 */
export const postProcess = async (
    bytes: Buffer,
    work: PostSlideWork[],
    props: Record<string, string> | null,
    customProps: Record<string, string> | null = null
): Promise<Buffer> => {
    const zip = await JSZip.loadAsync(bytes)
    /*  drop parts under the reserved OPC folder [trash]/ FIRST -- before the
        rid scan below reads stale deleted content. PowerPoint's incremental
        save and OneDrive/SharePoint co-authoring park logically deleted parts
        there, without relationship or content type. Carrying them into the
        output would resurrect deleted content and trip the every-part-needs-
        a-content-type self-check (engine/verify.ts ignores the folder for
        the same reason). PowerPoint ignores these parts too.  */
    for (const f of Object.keys(zip.files))
        if (f.startsWith("[trash]/"))
            zip.remove(f)
    /*  the rid counter must clear every rIdPptc left by EARLIER applies,
        or re-wiring (e.g. refilled hyperlinks) collides with old ids  */
    let maxRid = 9000
    for (const relsPart of Object.keys(zip.files).filter((f) => f.endsWith(".rels"))) {
        const text = await zip.file(relsPart)?.async("string") ?? ""
        for (const m of text.matchAll(/"rIdPptc(\d+)"/g))
            maxRid = Math.max(maxRid, Number(m[1]))
    }
    const post: Post = { zip, rid: maxRid + 1, tagsSeen: new Set() }
    const slides = await referencedSlides(zip)
    await gcParts(zip, slides)

    let mediaSeq = 1
    for (let i = 0; i < slides.length; i++) {
        const slidePart = slides[i] as string
        const job = work[i]
        const relsPart = `ppt/slides/_rels/${path.posix.basename(slidePart)}.rels`
        const slide = parseXml(await partText(zip, slidePart))
        const slideRels = parseXml(await partText(zip, relsPart))

        if (job !== undefined) {
            if (job.background !== null)
                setBackground(slide, job.background)
            if (job.footer !== null)
                await setFooter(zip, slidePart, slide, job.footer)
            for (const img of job.images) {
                const ext = path.extname(img.path).slice(1).toLowerCase() || "png"
                const mediaPart = `ppt/media/pptcImage${mediaSeq++}.${ext}`
                zip.file(mediaPart, readFileSync(img.path))
                await ensureImageDefault(zip, ext)
                const rid = nextRid(post)
                addRel(slideRels, rid,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    `../media/${path.posix.basename(mediaPart)}`)
                if (!fillPicturePlaceholder(slide, rid, img.phIdx, img.frame))
                    throw new PptcError("E_ADDR_NOTFOUND",
                        `slide ${i}: no empty picture placeholder idx ${img.phIdx}`)
            }
            if (job.notes !== null)
                await setNotes(post, slidePart, slideRels, job.notes)
            wireHyperlinks(post, slide, slideRels)
            setSlideVisibility(slide, job.hidden)
            if (job.slideTags.length > 0)
                await restoreSlideTags(post, slide, slideRels, job.slideTags)
        }
        const idsChanged = uniquifyShapeIds(slide)
        const relsChanged = pruneUnusedSlideRels(slide, slideRels)

        if (job !== undefined || idsChanged || relsChanged) {
            zip.file(slidePart, serializeXml(slide))
            zip.file(relsPart, serializeXml(slideRels))
        }
    }
    if (props !== null)
        await setProps(zip, props)
    if (customProps !== null)
        await setCustomProps(zip, customProps)
    /*  final sweeps: orphan assets, stale section refs and stale or
        duplicate content-type overrides accumulate across applies  */
    await dedupeSharedNotes(zip, slides)
    await gcAssets(zip)
    await pruneSectionRefs(zip)
    await cleanContentTypes(zip)
    await syncAppProperties(zip, slides)
    return await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
}
