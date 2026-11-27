/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/parts: neutral zip-package helpers shared by the write-path
**  (engine, post.ts) and the seed builder (seed.ts), plus the allocator that
**  hands out every fresh part name and id of an apply.
*/

import type JSZip from "jszip"
import path from "node:path"
import type { Document } from "@xmldom/xmldom"
import { PptcError } from "../infra/errors.js"
import { elements, parseXml, serializeXml } from "./xml.js"

/**
 *  Resolve a relationship target to a part name. Absolute targets
 *  (`/ppt/media/image1.png`) are valid OPC and kept on untouched parts.
 *
 *  @param base - folder of the part owning the relationship
 *  @param target - the relationship's Target attribute
 *  @returns the part name inside the archive
 */
export const resolveTarget = (base: string, target: string): string =>
    target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join(base, target))

/**
 *  The rels part of a part.
 *
 *  @param part - part name
 *  @returns the name of its relationships part
 */
export const relsPartOf = (part: string): string =>
    `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`

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
        for (const relsPart of Object.keys(zip.files).filter((f) => f.endsWith(".rels"))) {
            const text = await zip.file(relsPart)?.async("string") ?? ""
            for (const m of text.matchAll(/"rIdPptc(\d+)"/g))
                rid = Math.max(rid, Number(m[1]))
        }
        return new PartAllocator(new Set(Object.keys(zip.files)), rid + 1, Math.max(255, ...sldIds), new Set(sldIds))
    }

    /**
     *  A fresh part name `<dir>/<stem><n><ext>`, n above every number of the
     *  family (any extension) in use.
     *
     *  @param dir - folder, e.g. "ppt/slides"
     *  @param stem - name stem, e.g. "slide"
     *  @param ext - extension including the dot, e.g. ".xml"
     *  @returns the reserved part name
     */
    part(dir: string, stem: string, ext: string): string {
        const re = new RegExp(`^${reEscape(dir)}/${reEscape(stem)}(\\d+)\\.[^/]+$`)
        let n = 0
        for (const name of this.names)
            n = Math.max(n, Number(re.exec(name)?.[1] ?? "0"))
        const name = `${dir}/${stem}${n + 1}${ext}`
        this.names.add(name)
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
    const f = zip.file(part)
    if (f === null)
        throw new PptcError("E_ENGINE", `output archive is missing part ${part}`)
    return (await f.async("string")).replace(/^\uFEFF/, "")
}

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
        if (zip.file(part) === null || seen.has(part)) {
            o.parentNode?.removeChild(o)
            changed = true
        }
        else
            seen.add(part)
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
