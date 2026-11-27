/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/post: the package-wide steps of an apply, after every slide is in
**  place: sweep unreferenced notes and stale presentation references, apply
**  document properties, collect unreachable assets, sync sections, content
**  types and app properties. Slide-level edits live in engine/inplace.
**
**  LOCKSTEP: every repair trigger this module and engine/inplace/slides.ts
**  fix is asserted read-only by engine/verify.ts -- when a trigger is added
**  or changed here, teach the verifier the same rule (and vice versa), or
**  self-verify turns green fixes into false E_INTEGRITY failures.
*/

import type JSZip from "jszip"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { partText, cleanContentTypes, resolveTarget } from "./parts.js"
import { NS_REL, doubleQuoted, elements, firstElement, parseXml, serializeXml } from "./xml.js"

/**  remove a list of parts plus their rels and content-type overrides  */
export const removeParts = async (zip: JSZip, parts: string[]): Promise<void> => {
    if (parts.length === 0)
        return
    let ct = doubleQuoted(await partText(zip, "[Content_Types].xml"))
    for (const part of parts) {
        zip.remove(part)
        zip.remove(`${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`)
        ct = ct.replace(new RegExp(`<Override [^>]*PartName="/${part.replace(/[.\\/]/g, "\\$&")}"[^>]*/>`, "g"), "")
    }
    zip.file("[Content_Types].xml", ct)
}

/**  sync user-created sections with the slide list. Sections are
     contiguous runs of the slide order, so the slides that kept their
     relative order (the longest increasing run of their old positions)
     keep their section; removed slides drop out; new, copied and moved
     slides join the section of their predecessor (the first section when
     they lead the deck) -- as PowerPoint files them  */
export const syncSectionRefs = async (zip: JSZip): Promise<void> => {
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

/**  parts a slide carries (charts with their drawings and theme
     overrides, SmartArt, embeddings, media, comments, tags, ink and
     ActiveX controls);
     unreachable ones (e.g. of a removed slide) are garbage.
     Mirrored by the orphan check in engine/verify.ts.  */
export const ASSET_PART = /^ppt\/(charts|embeddings|media|diagrams|drawings|comments|tags|ink|activeX)\/|^ppt\/theme\/themeOverride[^/]*\.xml$/

/**  GC assets unreachable from any live part  */
export const gcAssets = async (zip: JSZip): Promise<void> => {
    const isAsset = (f: string): boolean =>
        ASSET_PART.test(f) && !f.endsWith("/") && !f.includes("/_rels/")
    const assets = Object.keys(zip.files).filter(isAsset)
    if (assets.length === 0)
        return
    const live = new Set<string>()
    const queue = Object.keys(zip.files).filter((f) =>
        f.endsWith(".rels") && !isAsset(f.replace(/_rels\/([^/]+)\.rels$/, "$1")))
    while (queue.length > 0) {
        const relsPart = queue.pop() as string
        const relsText = await zip.file(relsPart)?.async("string")
        if (relsText === undefined)
            continue
        const base = path.posix.dirname(path.posix.dirname(relsPart))
        for (const rel of elements(parseXml(relsText), "Relationship")) {
            if (rel.getAttribute("TargetMode") === "External")
                continue
            const target = resolveTarget(base, rel.getAttribute("Target") ?? "")
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

/**  content-type defaults for image media  */
export const IMAGE_TYPES: Record<string, string> = {
    png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
    gif: "image/gif", svg: "image/svg+xml", webp: "image/webp",
    bmp: "image/bmp", tif: "image/tiff", tiff: "image/tiff",
    emf: "image/x-emf", wmf: "image/x-wmf"
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

/**  keep docProps/app.xml in sync with the real slides: a deck declaring
     fewer slides than it has makes PowerPoint demand a repair on open.
     The leading TitlesOfParts entries (fonts, theme, OLE servers) and every
     HeadingPairs group except the last (the slide-title group) are kept; the
     slide count, the slide-title HeadingPairs count and the title list are
     rewritten from the actual slides.  */
export const syncAppProperties = async (zip: JSZip, slides: string[]): Promise<void> => {
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
export const setProps = async (zip: JSZip, props: Record<string, string>): Promise<void> => {
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
export const setCustomProps = async (zip: JSZip, patch: Record<string, string>): Promise<void> => {
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

/**  drop notes parts no slide references, presentation relationships
     whose target vanished or that no sldId uses, and custom-show entries
     whose relationship is gone  */
const sweepSlideRefs = async (zip: JSZip, slides: string[]): Promise<void> => {
    const notes = new Set<string>()
    for (const slide of slides) {
        const rels = doubleQuoted(await zip.file(`ppt/slides/_rels/${path.posix.basename(slide)}.rels`)?.async("string") ?? "")
        for (const m of rels.matchAll(/<Relationship\b[^>]*>/g))
            if (/Type="[^"]*\/notesSlide"/.test(m[0]))
                notes.add(resolveTarget("ppt/slides", /Target="([^"]+)"/.exec(m[0])?.[1] ?? ""))
    }
    await removeParts(zip, Object.keys(zip.files).filter((f) =>
        /^ppt\/notesSlides\/[^/]+\.xml$/.test(f) && !notes.has(f)))

    const presRelsPart = "ppt/_rels/presentation.xml.rels"
    const presRels = parseXml(await partText(zip, presRelsPart))
    const presXml = doubleQuoted(await partText(zip, "ppt/presentation.xml"))
    const liveSlideRids = new Set([ ...presXml.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g) ].map((m) => m[1]))
    let pruned = false
    for (const rel of elements(presRels, "Relationship")) {
        if (rel.getAttribute("TargetMode") === "External")
            continue
        const target = resolveTarget("ppt", rel.getAttribute("Target") ?? "")
        if (zip.file(target) === null
            || ((rel.getAttribute("Type") ?? "").endsWith("/slide") && !liveSlideRids.has(rel.getAttribute("Id") ?? ""))) {
            rel.parentNode?.removeChild(rel)
            pruned = true
        }
    }
    if (pruned)
        zip.file(presRelsPart, serializeXml(presRels))
    if (!presXml.includes("<p:custShow"))
        return
    const rids = new Set(elements(presRels, "Relationship").map((r) => r.getAttribute("Id") ?? ""))
    const pres = parseXml(presXml)
    let dropped = false
    for (const show of elements(pres, "p:custShow"))
        for (const sld of elements(show, "p:sld"))
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
 */
export const finishPackage = async (
    zip: JSZip,
    slides: string[],
    props: Record<string, string> | null,
    customProps: Record<string, string> | null
): Promise<void> => {
    await sweepSlideRefs(zip, slides)
    if (props !== null)
        await setProps(zip, props)
    if (customProps !== null)
        await setCustomProps(zip, customProps)
    await gcAssets(zip)
    await syncSectionRefs(zip)
    await cleanContentTypes(zip)
    await syncAppProperties(zip, slides)
}
