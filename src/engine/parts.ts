/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/parts: the OPC package layer below every other engine module --
**  part names, relationships (read through the DOM, never by regex), the
**  relationship-type taxonomy, content types, part removal and asset
**  collection, plus the allocator that hands out every fresh part name and
**  id of an apply.
*/

import type JSZip from "jszip"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../infra/errors.js"
import { NS_REL, decodeXml, elements, parseXml, serializeXml } from "./xml.js"

/**  OfficeDocument relationship type base URI  */
export const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/**  the relationship kinds (the last segment of the type URI) by role --
     the one taxonomy the engine, the content guard and the verifier share  */
export const RELS = {
    /**  leave a slide's own content: layout, master, other slides, notes
         master, theme  */
    foreign:   [ "slideLayout", "slideMaster", "slide", "notesMaster", "theme" ],
    /**  targets a clone inside the same package shares: layouts, masters,
         other slides, immutable media and external content  */
    shared:    [ "slideLayout", "slideMaster", "notesMaster", "slide", "image", "media", "video", "audio", "hyperlink", "theme" ],
    /**  carried without being referenced by id in the owner's XML (the
         SmartArt drawing is referenced from the SmartArt data part)  */
    implicit:  [ "slideLayout", "notesSlide", "diagramDrawing", "comments", "themeOverride", "vmlDrawing", "slideUpdateInfo" ],
    /**  never carry slide content of their own (or are checked separately):
         left out of the content guard's fingerprint  */
    unprinted: [ "slideLayout", "notesSlide", "hyperlink", "slide" ]
} as const

/**  the kind of a relationship type: the last segment of its URI
     (`.../relationships/slide` is "slide")  */
const relKind = (type: string): string =>
    type.slice(type.lastIndexOf("/") + 1)

/**
 *  Whether a relationship is of one of a role's kinds.
 *
 *  @param rel - the relationship
 *  @param kinds - the role, e.g. `RELS.shared`
 *  @returns true when it matches
 */
export const relIs = (rel: Pick<Rel, "kind">, kinds: readonly string[]): boolean =>
    kinds.includes(rel.kind)

/**  an empty relationships document  */
export const EMPTY_RELS = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    + `<Relationships xmlns="${NS_REL}"/>`

/**  resolve a relationship target to a part name; absolute targets
     (`/ppt/media/image1.png`) are valid OPC and kept on untouched parts  */
const resolveTarget = (base: string, target: string): string =>
    target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join(base, target))

/**
 *  The rels part of a part.
 *
 *  @param part - part name ("" for the package)
 *  @returns the name of its relationships part
 */
export const relsPartOf = (part: string): string =>
    part === "" ? "_rels/.rels" : `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`

/**
 *  The part a rels part belongs to ("" for the package rels `_rels/.rels`).
 *
 *  @param relsPart - the rels part name
 *  @returns the owning part
 */
export const relsOwner = (relsPart: string): string =>
    relsPart === "_rels/.rels" ? "" : relsPart.replace(/(^|\/)_rels\/([^/]+)\.rels$/, "$1$2")

/**
 *  The actual name of a part, looked up case-insensitively as OPC demands
 *  (`../media/photo1.png` names `ppt/media/Photo1.PNG`).
 *
 *  @param zip - the archive
 *  @param name - the part name as referenced
 *  @returns the name in the archive, null when there is no such part
 */
export const findPart = (zip: JSZip, name: string): string | null => {
    if (zip.file(name) !== null)
        return name
    const lower = name.toLowerCase()
    return Object.keys(zip.files).find((f) => !f.endsWith("/") && f.toLowerCase() === lower) ?? null
}

/**  a target with percent-escapes resolved (`my%20image.png`), unchanged
     when it carries none or invalid ones  */
const unescapeIri = (target: string): string => {
    try {
        return decodeURIComponent(target)
    }
    catch {
        return target
    }
}

/**  one relationship of a part  */
export interface Rel {
    /**  the Relationship element  */
    el: Element
    /**  its id  */
    id: string
    /**  its full type URI  */
    type: string
    /**  its kind, e.g. "slide" (see `relKind`)  */
    kind: string
    /**  the target part (case resolved when an archive is given), null for
         an external target  */
    target: string | null
}

/**
 *  The relationships of a part, targets resolved to part names.
 *
 *  @param owner - the owning part ("" for the package rels)
 *  @param rels - its rels DOM
 *  @param zip - the archive, to resolve target names case-insensitively
 *  @returns the relationships in document order
 */
export const relTargets = (owner: string, rels: Document, zip?: JSZip): Rel[] =>
    elements(rels, "Relationship").map((el) => {
        let target: string | null = null
        if (el.getAttribute("TargetMode") !== "External") {
            target = resolveTarget(path.posix.dirname(owner), el.getAttribute("Target") ?? "")
            if (zip !== undefined)
                target = findPart(zip, target) ?? findPart(zip, unescapeIri(target)) ?? target
        }
        const type = el.getAttribute("Type") ?? ""
        return { el, id: el.getAttribute("Id") ?? "", type, kind: relKind(type), target }
    })

/**
 *  Relationships by id. A duplicate id resolves to its first relationship
 *  everywhere -- reader, engine, repairs and verify agree.
 *
 *  @param rels - the relationships of a part
 *  @returns the first relationship of every id
 */
export const relMap = (rels: Rel[]): Map<string, Rel> => {
    const map = new Map<string, Rel>()
    for (const r of rels)
        if (!map.has(r.id))
            map.set(r.id, r)
    return map
}

/**
 *  Read the relationships of a part from an archive.
 *
 *  @param zip - the archive
 *  @param owner - the owning part
 *  @returns its relationships, empty when it has none
 */
export const readRels = async (zip: JSZip, owner: string): Promise<Rel[]> => {
    const relsPart = findPart(zip, relsPartOf(owner))
    return relsPart === null ? [] : relTargets(owner, parseXml(await partText(zip, relsPart)), zip)
}

/**
 *  Append a relationship to a rels document.
 *
 *  @param rels - the rels DOM
 *  @param id - the relationship id
 *  @param type - full relationship type URI
 *  @param target - target, relative to the owner (or a URL)
 *  @param external - whether the target is external
 */
export const appendRel = (rels: Document, id: string, type: string, target: string, external = false): void => {
    const rel = rels.createElementNS(NS_REL, "Relationship")
    rel.setAttribute("Id", id)
    rel.setAttribute("Type", type)
    rel.setAttribute("Target", target)
    if (external)
        rel.setAttribute("TargetMode", "External")
    rels.documentElement?.appendChild(rel)
}

/**  inflated content per archive entry: every check of an apply reads
     a part's bytes once  */
const inflated = new WeakMap<JSZip.JSZipObject, Promise<Buffer>>()

/**
 *  The bytes of an archive part, inflated once per archive entry.
 *
 *  @param zip - the archive
 *  @param part - the part name
 *  @returns its bytes, null when the part does not exist
 */
export const bytesOf = async (zip: JSZip, part: string): Promise<Buffer | null> => {
    const f = zip.file(findPart(zip, part) ?? part)
    if (f === null)
        return null
    let bytes = inflated.get(f)
    if (bytes === undefined) {
        bytes = f.async("nodebuffer")
        inflated.set(f, bytes)
    }
    return await bytes
}

/**  escape a string for a literal RegExp  */
const reEscape = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")

/**
 *  Hands out the fresh part names and ids of one apply. Every family
 *  continues above the highest number in use in the whole package -- never a
 *  count, never slide-local -- so a new part can never overwrite an existing
 *  one, and a removed slide's id is never re-used.
 */
export class PartAllocator {
    private constructor(
        private readonly names: Set<string>,
        private rid: number,
        private sldId: number,
        private readonly sldIds: Set<number>
    ) {}

    /**
     *  Scan a package once.
     *
     *  @param zip - the package
     *  @param sldIds - slide ids of the input deck (also those removed now)
     *  @returns the allocator
     */
    static async create(zip: JSZip, sldIds: number[]): Promise<PartAllocator> {
        let rid = 9000
        const names = new Set(Object.keys(zip.files).map((f) => f.toLowerCase()))
        for (const relsPart of Object.keys(zip.files).filter((f) => f.endsWith(".rels"))) {
            const text = await partText(zip, relsPart)
            for (const m of text.matchAll(/["']rIdPptc(\d+)["']/g))
                rid = Math.max(rid, Number(m[1]))

            /*  a dead relationship's target is reserved too: a fresh part
                under its name would silently bring the relationship back  */
            try {
                for (const r of relTargets(relsOwner(relsPart), parseXml(text)))
                    if (r.target !== null)
                        names.add(r.target.toLowerCase())
            }
            catch {
                /*  malformed rels are reported by verify  */
            }
        }
        return new PartAllocator(names, rid + 1, Math.max(255, ...sldIds), new Set(sldIds))
    }

    /**
     *  A fresh part name `<dir>/<stem><n><ext>`, n above every number of the
     *  family (any extension, any case -- OPC part names are case-insensitive)
     *  in use.
     *
     *  @param dir - folder, e.g. "ppt/slides"
     *  @param stem - name stem, e.g. "slide"
     *  @param ext - extension including the dot, e.g. ".xml"
     *  @returns the reserved part name
     */
    part(dir: string, stem: string, ext: string): string {
        const re = new RegExp(`^${reEscape(dir)}/${reEscape(stem)}(\\d+)\\.[^/]+$`, "i")
        let n = 0
        for (const name of this.names)
            n = Math.max(n, Number(re.exec(name)?.[1] ?? "0"))
        const name = `${dir}/${stem}${n + 1}${ext}`
        this.names.add(name.toLowerCase())
        return name
    }

    /**
     *  A fresh part name of the same family as an existing part (media are
     *  always named `image<n>`).
     *
     *  @param part - the part to name a sibling of
     *  @returns the reserved part name
     */
    like(part: string): string {
        const dir = path.posix.dirname(part)
        const ext = path.posix.extname(part)
        const stem = dir === "ppt/media" ? "image" : path.posix.basename(part, ext).replace(/\d+$/, "")
        return this.part(dir, stem === "" ? "part" : stem, ext)
    }

    /**
     *  A fresh relationship id, unique in every rels part of the package.
     *
     *  @returns the id
     */
    relId(): string {
        return `rIdPptc${this.rid++}`
    }

    /**
     *  A fresh slide id above every id of the input deck; ids above the
     *  schema maximum wrap to the lowest free one.
     *
     *  @returns the id
     */
    slideId(): number {
        this.sldId = this.sldId < 2147483647 ? this.sldId + 1 : 256
        while (this.sldIds.has(this.sldId))
            this.sldId++
        this.sldIds.add(this.sldId)
        return this.sldId
    }

    /**
     *  A counter of fresh shape ids for one slide.
     *
     *  @param slide - the slide document
     *  @returns a function returning the next free `p:cNvPr` id
     */
    static shapeIds(slide: Document): () => number {
        let max = 0
        for (const el of elements(slide, "p:cNvPr"))
            max = Math.max(max, Number(el.getAttribute("id") ?? "0") || 0)
        return () => ++max
    }
}

/**
 *  Read an archive part as a string; the part must exist.
 *
 *  @param zip - the archive
 *  @param part - part path inside the archive
 *  @returns the part content
 *  @throws PptcError E_ENGINE when the part is missing
 */
export const partText = async (zip: JSZip, part: string): Promise<string> => {
    const f = zip.file(findPart(zip, part) ?? part)
    if (f === null)
        throw new PptcError("E_ENGINE", `the deck is missing part ${part}`)
    return decodeXml(await f.async("uint8array"))
}

/**
 *  Read an archive part as a string, when it exists.
 *
 *  @param zip - the archive
 *  @param part - part path inside the archive (any case)
 *  @returns the decoded content, null when there is no such part
 */
export const optText = async (zip: JSZip, part: string): Promise<string | null> =>
    findPart(zip, part) === null ? null : await partText(zip, part)

/**
 *  Drop content-type overrides that are duplicates or whose part vanished.
 *
 *  @param zip - the archive to clean in place
 */
export const cleanContentTypes = async (zip: JSZip): Promise<void> => {
    const ct = parseXml(await partText(zip, "[Content_Types].xml"))
    const seen = new Set<string>()
    let changed = false
    for (const o of elements(ct, "Override")) {
        const part = (o.getAttribute("PartName") ?? "").replace(/^\//, "")
        if (findPart(zip, part) === null || seen.has(part.toLowerCase())) {
            o.parentNode?.removeChild(o)
            changed = true
        }
        else
            seen.add(part.toLowerCase())
    }
    if (changed)
        zip.file("[Content_Types].xml", serializeXml(ct))
}

/**
 *  Rewrite absolute relationship targets (`Target="/ppt/slides/slide1.xml"`,
 *  written by some generators) to the relative form, for read-only
 *  analysis. External targets stay as they are.
 *
 *  @param zip - archive to normalize in place
 *  @returns whether any target was rewritten
 */
export const normalizeRelTargets = async (zip: JSZip): Promise<boolean> => {
    let changed = false
    for (const relsPart of Object.keys(zip.files).filter((f) => f.endsWith(".rels"))) {
        const text = await partText(zip, relsPart)
        if (!/Target=["']\//.test(text))
            continue
        const base = path.posix.dirname(path.posix.dirname(relsPart)).replace(/^\.$/, "")
        const doc = parseXml(text)
        let rewritten = false
        for (const rel of elements(doc, "Relationship")) {
            const target = rel.getAttribute("Target") ?? ""
            if (rel.getAttribute("TargetMode") !== "External" && target.startsWith("/")) {
                rel.setAttribute("Target", path.posix.relative(base, target.slice(1)))
                rewritten = true
            }
        }
        if (rewritten) {
            zip.file(relsPart, serializeXml(doc))
            changed = true
        }
    }
    return changed
}

/**  content-type defaults for image media  */
export const IMAGE_TYPES: Record<string, string> = {
    png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
    gif: "image/gif", svg: "image/svg+xml", webp: "image/webp",
    bmp: "image/bmp", tif: "image/tiff", tiff: "image/tiff",
    emf: "image/x-emf", wmf: "image/x-wmf"
}

/**  parts a slide carries (charts with their drawings and theme
     overrides, SmartArt, embeddings, media, comments, tags, ink and
     ActiveX controls); unreachable ones (e.g. of a removed slide) are
     garbage -- collected by `gcAssets`, reported by engine/verify  */
export const ASSET_PART = /^ppt\/(charts|embeddings|media|diagrams|drawings|comments|tags|ink|activeX)\/|^ppt\/theme\/themeOverride[^/]*\.xml$/

/**
 *  Remove parts together with their rels and content-type overrides.
 *
 *  @param zip - the archive to edit in place
 *  @param parts - the parts to remove
 */
export const removeParts = async (zip: JSZip, parts: string[]): Promise<void> => {
    if (parts.length === 0)
        return
    const doomed = new Set(parts.map((p) => `/${p.toLowerCase()}`))
    for (const part of parts) {
        zip.remove(part)
        zip.remove(findPart(zip, relsPartOf(part)) ?? relsPartOf(part))
    }
    const ct = parseXml(await partText(zip, "[Content_Types].xml"))
    for (const o of elements(ct, "Override"))
        if (doomed.has((o.getAttribute("PartName") ?? "").toLowerCase()))
            o.parentNode?.removeChild(o)
    zip.file("[Content_Types].xml", serializeXml(ct))
}

/**
 *  Remove the assets no live part reaches (e.g. those of removed slides).
 *
 *  @param zip - the archive to edit in place
 */
export const gcAssets = async (zip: JSZip): Promise<void> => {
    const isAsset = (f: string): boolean =>
        ASSET_PART.test(f) && !f.endsWith("/") && !f.includes("/_rels/")
    const assets = Object.keys(zip.files).filter(isAsset)
    if (assets.length === 0)
        return
    const live = new Set<string>()
    const queue = Object.keys(zip.files).filter((f) => f.endsWith(".rels") && !isAsset(relsOwner(f)))
        .map(relsOwner)
    while (queue.length > 0) {
        for (const { target } of await readRels(zip, queue.pop() as string))
            if (target !== null && isAsset(target) && !live.has(target)) {
                live.add(target)
                queue.push(target)
            }
    }
    await removeParts(zip, assets.filter((a) => !live.has(a)))
}

/**
 *  The layout parts of a deck in slide-master order -- the order the
 *  reader (and so every layout index) uses; part names need not follow it.
 *
 *  @param zip - the archive
 *  @returns the layout part names
 */
export const layoutOrder = async (zip: JSZip): Promise<string[]> => {
    const presRels = relMap(await readRels(zip, "ppt/presentation.xml"))
    const pres = findPart(zip, "ppt/presentation.xml") === null ? null : parseXml(await partText(zip, "ppt/presentation.xml"))
    const listed = (pres === null ? [] : elements(pres, "p:sldMasterId"))
        .map((m) => presRels.get(m.getAttribute("r:id") ?? "")?.target ?? null)
        .filter((part): part is string => part !== null && findPart(zip, part) !== null)
    const masters = listed.length > 0 ? listed
        : Object.keys(zip.files).filter((f) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f)).sort()
    const order: string[] = []
    for (const master of masters) {
        const rels = relMap(await readRels(zip, master))
        for (const layoutId of elements(parseXml(await partText(zip, master)), "p:sldLayoutId")) {
            const rel = rels.get(layoutId.getAttribute("r:id") ?? "")
            if (rel !== undefined && rel.target !== null && rel.kind === "slideLayout")
                order.push(rel.target)
        }
    }
    return order
}

/**  one entry of a deck's slide list  */
export interface SlideEntry {
    /**  its p:sldId element  */
    el: Element
    /**  its slide id  */
    id: number
    /**  its slide part, null when it resolves to no slide part  */
    part: string | null
}

/**
 *  The slide list of a deck in order -- the one reading the reader, the
 *  engine, the content guard and verify share.
 *
 *  @param zip - the archive (to tell existing slide parts)
 *  @param pres - the presentation DOM
 *  @param presRels - the presentation's relationships
 *  @returns the sldId entries in order
 */
export const slideList = (zip: JSZip, pres: Document, presRels: Rel[]): SlideEntry[] => {
    const rels = relMap(presRels)
    return elements(pres, "p:sldId").map((el) => {
        const rel = rels.get(el.getAttribute("r:id") ?? "")
        const part = rel !== undefined && rel.kind === "slide" && rel.target !== null && findPart(zip, rel.target) !== null
            ? rel.target : null
        return { el, id: Number(el.getAttribute("id")), part }
    })
}
