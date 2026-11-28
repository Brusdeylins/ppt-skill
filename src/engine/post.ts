/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/post: the package-wide steps of an apply, after every slide is in
**  place: sweep unreferenced notes and stale presentation references, apply
**  document properties, collect unreachable assets, sync sections, content
**  types and app properties. Slide-level edits live in engine/inplace; the
**  package-level rules these steps satisfy are asserted by engine/verify.
*/

import type JSZip from "jszip"
import { createHash } from "node:crypto"
import type { Document, Element } from "@xmldom/xmldom"
import { appendRel, cleanContentTypes, findPart, gcAssets, optText, partText, readRels, removeParts, tryRels } from "./parts.js"
import { elements, firstElement, parseXml, serializeXml, utf8Declared, xmlEscape, xmlStorable } from "./xml.js"

/**  sync user-created sections with the slide list. Sections are
     contiguous runs of the slide order, so the slides that kept their
     relative order (the longest increasing run of their old positions)
     keep their section; removed slides drop out; new, copied and moved
     slides join the section of their predecessor (the first section when
     they lead the deck) -- as PowerPoint files them  */
const syncSectionRefs = async (zip: JSZip): Promise<void> => {
    const part = "ppt/presentation.xml"
    const xml = await partText(zip, part)
    if (!xml.includes("<p14:sectionLst"))
        return
    const pres = parseXml(xml)
    const sections = elements(pres, "p14:section")
    const lists = sections.map((sec) => firstElement(sec, "p14:sldIdLst"))
    if (sections.length === 0 || lists.some((l) => l === null))
        return
    const sectionOf = new Map<string, number>()
    const oldPos = new Map<string, number>()
    sections.forEach((sec, k) => {
        for (const ref of elements(sec, "p14:sldId")) {
            const id = ref.getAttribute("id") ?? ""
            if (!sectionOf.has(id)) {
                sectionOf.set(id, k)
                oldPos.set(id, oldPos.size)
            }
        }
    })
    const order = elements(pres, "p:sldId").map((s) => s.getAttribute("id") ?? "")

    /*  longest increasing run of old positions (O(n^2), decks are small)  */
    const kept = order.filter((id) => oldPos.has(id))
    const len = kept.map(() => 1)
    const prev = kept.map(() => -1)
    for (let i = 0; i < kept.length; i++)
        for (let j = 0; j < i; j++)
            if ((oldPos.get(kept[j] as string) as number) < (oldPos.get(kept[i] as string) as number)
                && (len[j] as number) + 1 > (len[i] as number)) {
                len[i] = (len[j] as number) + 1
                prev[i] = j
            }
    const stable = new Set<string>()
    for (let i = len.indexOf(Math.max(0, ...len)); i >= 0; i = prev[i] as number)
        stable.add(kept[i] as string)

    const members: string[][] = sections.map(() => [])
    let current = 0
    for (const id of order) {
        if (stable.has(id))
            current = sectionOf.get(id) as number
        members[current]?.push(id)
    }
    for (const [ k, list ] of lists.entries()) {
        const l = list as Element
        for (const ref of elements(l, "p14:sldId"))
            l.removeChild(ref)
        for (const id of members[k] as string[]) {
            const ref = pres.createElementNS(l.namespaceURI, "p14:sldId")
            ref.setAttribute("id", id)
            l.appendChild(ref)
        }
    }
    zip.file(part, serializeXml(pres))
}

/**  the extension of the presentation that carries the sections  */
const SECTION_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"

/**  the PowerPoint 2010 namespace of `p14:sectionLst`  */
const NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

/**  one section as the engine writes it  */
export interface SectionOut {
    /**  section name  */
    name: string
    /**  its id (GUID in braces), null for a new section  */
    id: string | null
    /**  the sldIds of its slides in order  */
    slideIds: number[]
}

/**  a section id for a new section: a GUID derived from the name and the
     ids in use, so an apply stays deterministic  */
const sectionGuid = (name: string, used: Set<string>): string => {
    for (let salt = 0; ; salt++) {
        const h = createHash("sha1").update(`${name}\u0000${[ ...used ].sort().join(",")}\u0000${salt}`).digest("hex").toUpperCase()
        const guid = `{${h.slice(0, 8)}-${h.slice(8, 12)}-4${h.slice(13, 16)}-8${h.slice(17, 20)}-${h.slice(20, 32)}}`
        if (!used.has(guid))
            return guid
    }
}

/**  write the planned sections into the presentation: the section list
     is replaced as a whole (kept sections keep their id), an empty list
     removes it with its extension  */
const writeSections = async (zip: JSZip, sections: SectionOut[]): Promise<void> => {
    const part = "ppt/presentation.xml"
    const pres = parseXml(await partText(zip, part))
    const root = pres.documentElement
    if (root === null)
        return
    let extLst = elements(pres, "p:extLst").find((el) => el.parentNode === root) ?? null
    let ext = extLst === null ? null
        : elements(extLst, "p:ext").find((el) => el.getAttribute("uri") === SECTION_EXT) ?? null
    if (sections.length === 0) {
        if (ext === null)
            return
        extLst?.removeChild(ext)
        if (extLst !== null && elements(extLst, "p:ext").length === 0)
            root.removeChild(extLst)
        zip.file(part, serializeXml(pres))
        return
    }
    if (extLst === null) {
        extLst = pres.createElementNS(root.namespaceURI, "p:extLst")
        root.appendChild(extLst)
    }
    if (ext === null) {
        ext = pres.createElementNS(root.namespaceURI, "p:ext")
        ext.setAttribute("uri", SECTION_EXT)
        extLst.insertBefore(ext, extLst.firstChild)
    }
    while (ext.firstChild !== null)
        ext.removeChild(ext.firstChild)
    const lst = pres.createElementNS(NS_P14, "p14:sectionLst")
    const used = new Set(sections.flatMap((s) => s.id === null ? [] : [ s.id ]))
    for (const s of sections) {
        const sec = pres.createElementNS(NS_P14, "p14:section")
        const id = s.id ?? sectionGuid(s.name, used)
        used.add(id)
        sec.setAttribute("name", s.name)
        sec.setAttribute("id", id)
        const ids = pres.createElementNS(NS_P14, "p14:sldIdLst")
        for (const slideId of s.slideIds) {
            const ref = pres.createElementNS(NS_P14, "p14:sldId")
            ref.setAttribute("id", String(slideId))
            ids.appendChild(ref)
        }
        sec.appendChild(ids)
        lst.appendChild(sec)
    }
    ext.appendChild(lst)
    zip.file(part, serializeXml(pres))
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

/**  keep docProps/app.xml in sync with the real slides: a deck declaring
     fewer slides than it has makes PowerPoint demand a repair on open.
     Every HeadingPairs group but the slide-title group (the last one, or
     the one before a trailing custom-show group) and its TitlesOfParts
     entries are kept; the slide count, the slide-title count and titles are
     rewritten from the actual slides (element content only, edited as text)  */
const syncAppProperties = async (zip: JSZip, slides: string[]): Promise<void> => {
    const part = findPart(zip, "docProps/app.xml")
    if (part === null)
        return
    let xml = await partText(zip, part)
    const n = slides.length
    const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(xml)
    const tp = /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(xml)
    if (hp === null || tp === null) {
        zip.file(part, utf8Declared(xml.replace(/<Slides>\d+<\/Slides>/, () => `<Slides>${n}</Slides>`)))
        return
    }
    const hpBody = hp[1] ?? ""
    const tpBody = tp[1] ?? ""
    const titles: string[] = []
    for (const slidePart of slides)
        titles.push(slideTitle(parseXml(await partText(zip, slidePart))))
    /*  the slide-title group: the last one, unless custom shows follow it  */
    const shows = elements(parseXml(await partText(zip, "ppt/presentation.xml")), "p:custShow").length
    const counts = [ ...hpBody.matchAll(/<vt:i4>(\d+)<\/vt:i4>/g) ].map((m) => Number(m[1]))
    const declared = Number(/<Slides>(\d+)<\/Slides>/.exec(xml)?.[1] ?? "-1")
    const at = shows > 0 && counts.length >= 2 && counts[counts.length - 1] === shows && counts[counts.length - 2] === declared
        ? counts.length - 2 : counts.length - 1
    const leading = counts.slice(0, at).reduce((a, b) => a + b, 0)
    const old = counts[at] ?? 0
    let seen = 0
    const newHp = hpBody.replace(/<vt:i4>\d+<\/vt:i4>/g, (m) =>
        seen++ === at ? `<vt:i4>${n}</vt:i4>` : m)
    const entries = [ ...tpBody.matchAll(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g) ].map((m) => m[1] as string)
    const kept = [ ...entries.slice(0, leading), ...titles.map((t) => xmlEscape(xmlStorable(t))), ...entries.slice(leading + old) ]
    const lpstr = kept.map((e) => `<vt:lpstr>${e}</vt:lpstr>`).join("")

    /*  replacer functions: a title like "US$'000" must not read as a
        replacement pattern  */
    xml = xml
        .replace(/<Slides>\d+<\/Slides>/, () => `<Slides>${n}</Slides>`)
        .replace(hp[0], () => `<HeadingPairs>${newHp}</HeadingPairs>`)
        .replace(tp[0], () => `<TitlesOfParts><vt:vector size="${kept.length}" baseType="lpstr">${lpstr}</vt:vector></TitlesOfParts>`)
    zip.file(part, utf8Declared(xml))
}

/**  update document core properties  */
const setProps = async (zip: JSZip, props: Record<string, string>): Promise<void> => {
    const part = findPart(zip, "docProps/core.xml")
    if (part === null)
        return
    const text = await partText(zip, part)
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
     the content-type override and the package relationship exist. Only
     the named properties change (a string value, `vt:lpwstr`, attributes
     like a link target kept); every other property stays exactly as it
     is, whatever its type. PowerPoint keeps custom properties across
     edit/save, so they make the deck self-describing.  */
const setCustomProps = async (zip: JSZip, patch: Record<string, string>): Promise<void> => {
    const existing = await optText(zip, CUSTOM_PROPS_PART)
    const doc = parseXml(existing ?? "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
        + `<Properties xmlns="${NS_CUSTOM}" xmlns:vt="${NS_VT}"></Properties>`)
    const root = doc.documentElement
    if (root === null)
        return
    const props = elements(doc, "property")

    /*  pids are unique and MUST start at 2 (pid 1 is reserved)  */
    let pid = Math.max(1, ...props.map((p) => Number(p.getAttribute("pid") ?? "0")).filter(Number.isFinite)) + 1
    for (const [ name, value ] of Object.entries(patch)) {
        let prop = props.find((p) => p.getAttribute("name") === name)
        if (prop === undefined) {
            prop = doc.createElementNS(root.namespaceURI, "property")
            prop.setAttribute("fmtid", CUSTOM_FMTID)
            prop.setAttribute("pid", String(pid++))
            prop.setAttribute("name", name)
            root.appendChild(prop)
            props.push(prop)
        }
        while (prop.firstChild !== null)
            prop.removeChild(prop.firstChild)
        const vt = doc.createElementNS(NS_VT, "vt:lpwstr")
        vt.appendChild(doc.createTextNode(value))
        prop.appendChild(vt)
    }
    zip.file(CUSTOM_PROPS_PART, serializeXml(doc))

    /*  content-type override (idempotent)  */
    const ct = parseXml(await partText(zip, "[Content_Types].xml"))
    if (!elements(ct, "Override").some((o) => (o.getAttribute("PartName") ?? "").toLowerCase() === `/${CUSTOM_PROPS_PART}`.toLowerCase())) {
        const override = ct.createElementNS(ct.documentElement?.namespaceURI ?? null, "Override")
        override.setAttribute("PartName", `/${CUSTOM_PROPS_PART}`)
        override.setAttribute("ContentType", CUSTOM_PROPS_CT)
        ct.documentElement?.appendChild(override)
        zip.file("[Content_Types].xml", serializeXml(ct))
    }

    /*  package relationship from the root .rels (idempotent)  */
    const rootRels = parseXml(await partText(zip, "_rels/.rels"))
    const present = elements(rootRels, "Relationship")
        .some((r) => (r.getAttribute("Type") ?? "") === CUSTOM_PROPS_REL)
    if (!present) {
        const ids = new Set(elements(rootRels, "Relationship").map((r) => r.getAttribute("Id")))
        let n = 1
        while (ids.has(`rId${n}`))
            n++
        appendRel(rootRels, `rId${n}`, CUSTOM_PROPS_REL, CUSTOM_PROPS_PART)
        zip.file("_rels/.rels", serializeXml(rootRels))
    }
}

/**  drop notes parts no slide references, presentation relationships
     whose target vanished or that no sldId uses, and custom-show entries
     whose relationship is gone  */
const sweepSlideRefs = async (zip: JSZip, slides: string[]): Promise<void> => {
    /*  a slide with unreadable relationships may own any notes part  */
    const notes = new Set<string>()
    let known = true
    for (const slide of slides) {
        const rels = await tryRels(zip, slide)
        known &&= rels !== null
        for (const r of rels ?? [])
            if (r.kind === "notesSlide" && r.target !== null)
                notes.add(r.target)
    }
    if (known)
        await removeParts(zip, Object.keys(zip.files).filter((f) =>
            /^ppt\/notesSlides\/[^/]+\.xml$/.test(f) && !notes.has(f)))

    const presRelsPart = "ppt/_rels/presentation.xml.rels"
    const presRels = parseXml(await partText(zip, presRelsPart))
    const pres = parseXml(await partText(zip, "ppt/presentation.xml"))
    const liveSlideRids = new Set(elements(pres, "p:sldId").map((s) => s.getAttribute("r:id") ?? ""))
    let pruned = false
    for (const r of await readRels(zip, "ppt/presentation.xml")) {
        if (r.target === null || (zip.file(r.target) !== null && !(r.kind === "slide" && !liveSlideRids.has(r.id))))
            continue
        for (const rel of elements(presRels, "Relationship"))
            if (rel.getAttribute("Id") === r.id)
                rel.parentNode?.removeChild(rel)
        pruned = true
    }
    if (pruned)
        zip.file(presRelsPart, serializeXml(presRels))
    const rids = new Set(elements(presRels, "Relationship").map((r) => r.getAttribute("Id") ?? ""))
    let dropped = false

    /*  p:sld only occurs in the slide lists of custom shows  */
    for (const sld of elements(pres, "p:sld"))
        if (!rids.has(sld.getAttribute("r:id") ?? "")) {
            sld.parentNode?.removeChild(sld)
            dropped = true
        }
    if (dropped)
        zip.file("ppt/presentation.xml", serializeXml(pres))
}

/**
 *  The package-wide steps of an apply, after every slide is in place:
 *  sweep unreferenced notes and stale presentation references, document
 *  properties, unreachable assets, sections, content types, app properties.
 *
 *  @param zip - the edited package
 *  @param slides - output slide parts in final order
 *  @param props - document core-property patch, null for none
 *  @param customProps - custom document-property patch, null for none
 *  @param sections - the planned sections, null to keep the deck's in
 *                    step with the slide list
 */
export const finishPackage = async (
    zip: JSZip,
    slides: string[],
    props: Record<string, string> | null,
    customProps: Record<string, string> | null,
    sections: SectionOut[] | null = null
): Promise<void> => {
    await sweepSlideRefs(zip, slides)
    if (props !== null)
        await setProps(zip, props)
    if (customProps !== null)
        await setCustomProps(zip, customProps)
    await gcAssets(zip)
    if (sections === null)
        await syncSectionRefs(zip)
    else
        await writeSections(zip, sections)
    await cleanContentTypes(zip)
    await syncAppProperties(zip, slides)
}
