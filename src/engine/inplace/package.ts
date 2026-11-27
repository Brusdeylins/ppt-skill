/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/package: part-level access to the deck package that the
**  in-place engine edits. Parts are read and written individually; nothing
**  the plan does not touch is ever re-serialized. Also clones the rel graph
**  of a part (slide.copy, transplanted generator parts) under fresh names.
*/

import JSZip from "jszip"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PartAllocator, partText, relsPartOf, resolveTarget } from "../parts.js"
import { IMAGE_TYPES } from "../post.js"
import { NS_REL, elements, parseXml, serializeXml } from "../xml.js"

/**  relationship types whose target a clone inside the same package shares:
     layouts, masters, other slides, immutable media and external content  */
const SHARED_RELS = [ "/slideLayout", "/slideMaster", "/notesMaster", "/slide", "/image", "/media", "/video", "/audio", "/hyperlink", "/theme" ]

/**  an empty relationships document  */
const EMPTY_RELS = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    + `<Relationships xmlns="${NS_REL}"/>`

/**  the deck package of one apply  */
export class Package {
    private constructor(
        /**  the archive being edited  */
        readonly zip: JSZip,
        /**  fresh names and ids  */
        readonly alloc: PartAllocator
    ) {}

    /**
     *  Open the package from the deck bytes.
     *
     *  @param bytes - the deck file as on disk
     *  @param sldIds - slide ids of the input deck
     *  @returns the package
     */
    static async open(bytes: Buffer, sldIds: number[]): Promise<Package> {
        const zip = await JSZip.loadAsync(bytes)

        /*  the reserved OPC folder [trash]/ holds logically deleted parts
            (PowerPoint incremental save, co-authoring): no content type, no
            relationship -- carrying them over would resurrect deleted content  */
        for (const f of Object.keys(zip.files))
            if (f.startsWith("[trash]/"))
                zip.remove(f)
        return new Package(zip, await PartAllocator.create(zip, sldIds))
    }

    /**
     *  Whether a part exists.
     *
     *  @param part - part name
     *  @returns true when present
     */
    has(part: string): boolean {
        return this.zip.file(part) !== null
    }

    /**
     *  Text of a part (byte-order mark dropped); the part must exist.
     *
     *  @param part - part name
     *  @returns its content
     */
    async text(part: string): Promise<string> {
        return await partText(this.zip, part)
    }

    /**
     *  Parsed XML of a part; the part must exist.
     *
     *  @param part - part name
     *  @returns its DOM
     */
    async xml(part: string): Promise<Document> {
        return parseXml(await this.text(part))
    }

    /**
     *  The relationships of a part, an empty list when it has none. Absolute
     *  targets are rewritten to the relative form, as the rels part is about
     *  to be written anyway.
     *
     *  @param part - the owning part
     *  @returns the rels DOM
     */
    async rels(part: string): Promise<Document> {
        const relsPart = relsPartOf(part)
        const doc = parseXml(this.has(relsPart) ? await this.text(relsPart) : EMPTY_RELS)
        const base = path.posix.dirname(part)
        for (const rel of elements(doc, "Relationship")) {
            const target = rel.getAttribute("Target") ?? ""
            if (rel.getAttribute("TargetMode") !== "External" && target.startsWith("/"))
                rel.setAttribute("Target", path.posix.relative(base, target.slice(1)))
        }
        return doc
    }

    /**
     *  Write a part.
     *
     *  @param part - part name
     *  @param data - XML document or raw content
     */
    write(part: string, data: Document | string | Uint8Array): void {
        this.zip.file(part, typeof data === "string" || data instanceof Uint8Array ? data : serializeXml(data))
    }

    /**
     *  Write the relationships of a part.
     *
     *  @param part - the owning part
     *  @param rels - the rels DOM
     */
    writeRels(part: string, rels: Document): void {
        this.write(relsPartOf(part), rels)
    }

    /**
     *  Append a relationship with a fresh id.
     *
     *  @param rels - the rels DOM
     *  @param type - full relationship type URI
     *  @param target - target, relative to the owner (or a URL)
     *  @param external - whether the target is external
     *  @returns the new relationship id
     */
    addRel(rels: Document, type: string, target: string, external = false): string {
        const id = this.alloc.relId()
        const rel = rels.createElementNS(NS_REL, "Relationship")
        rel.setAttribute("Id", id)
        rel.setAttribute("Type", type)
        rel.setAttribute("Target", target)
        if (external)
            rel.setAttribute("TargetMode", "External")
        rels.documentElement?.appendChild(rel)
        return id
    }

    /**
     *  Register a content-type override for a part.
     *
     *  @param part - the part
     *  @param contentType - its content type
     */
    async typeAs(part: string, contentType: string): Promise<void> {
        this.write("[Content_Types].xml", (await this.text("[Content_Types].xml"))
            .replace("</Types>", `<Override PartName="/${part}" ContentType="${contentType}"/></Types>`))
    }

    /**
     *  Give a part a content type: the override it has in the source
     *  package, else a default for its extension.
     *
     *  @param part - the (new) part
     *  @param sourceCt - [Content_Types].xml of the package the content comes from
     *  @param sourcePart - the part it was copied from
     */
    async typePart(part: string, sourceCt: string, sourcePart: string): Promise<void> {
        const override = new RegExp(`<Override [^>]*PartName="/${sourcePart.replace(/[.\\/]/g, "\\$&")}"[^>]*/>`).exec(sourceCt)
        if (override !== null) {
            await this.typeAs(part, /ContentType="([^"]+)"/.exec(override[0])?.[1] ?? "")
            return
        }
        const ct = await this.text("[Content_Types].xml")
        const ext = path.posix.extname(part).slice(1).toLowerCase()
        if (new RegExp(`<Default [^>]*Extension="${ext}"`, "i").test(ct))
            return
        const type = new RegExp(`<Default [^>]*Extension="${ext}"[^>]*ContentType="([^"]+)"`, "i").exec(sourceCt)?.[1]
            ?? IMAGE_TYPES[ext] ?? "application/octet-stream"
        this.write("[Content_Types].xml", ct.replace(/<Default\b/, `<Default Extension="${ext}" ContentType="${type}"/><Default`))
    }

    /**
     *  Copy a part and everything it relates to under fresh names. Inside
     *  one package, layouts, masters, other slides and media stay shared;
     *  from another package (the element generator) everything is copied.
     *
     *  @param from - the package holding the part
     *  @param part - the part to copy
     *  @param done - parts already copied in this operation (source to copy);
     *                a relationship to one of them points at its copy
     *  @returns the name of the copy
     */
    async clonePart(from: Package, part: string, done: Map<string, string> = new Map()): Promise<string> {
        const copied = done.get(part)
        if (copied !== undefined)
            return copied
        const copy = this.alloc.like(part)
        done.set(part, copy)
        this.write(copy, await (from.zip.file(part) as JSZip.JSZipObject).async("uint8array"))
        await this.typePart(copy, await from.text("[Content_Types].xml"), part)
        if (!from.has(relsPartOf(part)))
            return copy
        const rels = await from.rels(part)
        await this.relinkClone(from, part, copy, rels, done)
        this.writeRels(copy, rels)
        return copy
    }

    /**
     *  Re-point the relationships of a copied part: targets copied in the
     *  same operation point at their copy, shared ones keep their part,
     *  everything else is copied recursively.
     *
     *  @param from - the package holding the original
     *  @param part - the original part
     *  @param copy - the name of the copy
     *  @param rels - the relationships to re-point (in place)
     *  @param done - parts already copied in this operation
     */
    async relinkClone(from: Package, part: string, copy: string, rels: Document, done: Map<string, string>): Promise<void> {
        for (const rel of elements(rels, "Relationship")) {
            if (rel.getAttribute("TargetMode") === "External")
                continue
            const target = resolveTarget(path.posix.dirname(part), rel.getAttribute("Target") ?? "")
            const type = rel.getAttribute("Type") ?? ""
            let next = done.get(target)
            if (next === undefined && from === this && SHARED_RELS.some((k) => type.endsWith(k)))
                next = target
            if (next === undefined && !from.has(target)) {
                rel.parentNode?.removeChild(rel)
                continue
            }
            next ??= await this.clonePart(from, target, done)
            rel.setAttribute("Target", path.posix.relative(path.posix.dirname(copy), next))
        }
    }

    /**
     *  Remove parts together with their rels and content-type overrides.
     *
     *  @param parts - the parts to remove
     */
    async remove(parts: string[]): Promise<void> {
        if (parts.length === 0)
            return
        let ct = await this.text("[Content_Types].xml")
        for (const part of parts) {
            this.zip.remove(part)
            this.zip.remove(relsPartOf(part))
            ct = ct.replace(new RegExp(`<Override PartName="/${part.replace(/[.\\/]/g, "\\$&")}"[^>]*/>`, "g"), "")
        }
        this.write("[Content_Types].xml", ct)
    }

    /**
     *  Targets of the relationships of a part, resolved to part names.
     *
     *  @param part - the owning part
     *  @param rels - its rels DOM
     *  @returns relationship elements with their resolved target
     */
    static targets(part: string, rels: Document): { rel: Element, type: string, target: string | null }[] {
        return elements(rels, "Relationship").map((rel) => ({
            rel,
            type: rel.getAttribute("Type") ?? "",
            target: rel.getAttribute("TargetMode") === "External" ? null
                : resolveTarget(path.posix.dirname(part), rel.getAttribute("Target") ?? "")
        }))
    }
}
