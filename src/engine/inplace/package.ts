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
import type { Document } from "@xmldom/xmldom"
import {
    EMPTY_RELS, IMAGE_TYPES, PartAllocator, RELS, appendRel, findPart, partText, relIs, relTargets,
    relsPartOf, removeParts, type Rel
} from "../parts.js"
import { elements, parseXml, serializeXml, xmlEscape } from "../xml.js"

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
     *  Whether a part exists (part names are case-insensitive).
     *
     *  @param part - part name
     *  @returns true when present
     */
    has(part: string): boolean {
        return findPart(this.zip, part) !== null
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
        appendRel(rels, id, type, target, external)
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
            .replace("</Types>", () => `<Override PartName="/${xmlEscape(part, true)}" ContentType="${xmlEscape(contentType, true)}"/></Types>`))
    }

    /**
     *  Give a part a content type: the override it has in the source
     *  package, else a default for its extension.
     *
     *  @param part - the (new) part
     *  @param sourceCtText - [Content_Types].xml of the package the content comes from
     *  @param sourcePart - the part it was copied from
     */
    async typePart(part: string, sourceCtText: string, sourcePart: string): Promise<void> {
        const sourceCt = sourceCtText === "" ? null : parseXml(sourceCtText)
        const override = sourceCt === null ? undefined : elements(sourceCt, "Override").find((o) =>
            (o.getAttribute("PartName") ?? "").toLowerCase() === `/${sourcePart.toLowerCase()}`)
        if (override !== undefined) {
            await this.typeAs(part, override.getAttribute("ContentType") ?? "")
            return
        }
        const ext = path.posix.extname(part).slice(1).toLowerCase()
        const defaultOf = (ct: Document | null): string | null => ct === null ? null : elements(ct, "Default")
            .find((d) => (d.getAttribute("Extension") ?? "").toLowerCase() === ext)?.getAttribute("ContentType") ?? null
        const ct = parseXml(await this.text("[Content_Types].xml"))
        if (defaultOf(ct) !== null)
            return
        const def = ct.createElementNS(ct.documentElement?.namespaceURI ?? null, "Default")
        def.setAttribute("Extension", ext)
        def.setAttribute("ContentType", defaultOf(sourceCt) ?? IMAGE_TYPES[ext] ?? "application/octet-stream")
        ct.documentElement?.insertBefore(def, ct.documentElement.firstChild)
        this.write("[Content_Types].xml", ct)
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
     *  Re-point the relationships of a copied part: dead targets go, targets
     *  copied in the same operation point at their copy, shared ones keep
     *  their part, everything else is copied recursively.
     *
     *  @param from - the package holding the original
     *  @param part - the original part
     *  @param copy - the name of the copy
     *  @param rels - the relationships to re-point (in place)
     *  @param done - parts already copied in this operation
     */
    private async relinkClone(from: Package, part: string, copy: string, rels: Document, done: Map<string, string>): Promise<void> {
        for (const { el, type, target } of from.targets(part, rels)) {
            if (target === null)
                continue
            let next = done.get(target)
            if (next === undefined && !from.has(target)) {
                el.parentNode?.removeChild(el)
                continue
            }
            if (next === undefined && from === this && relIs(type, RELS.shared))
                next = target
            next ??= await this.clonePart(from, target, done)
            el.setAttribute("Target", path.posix.relative(path.posix.dirname(copy), next))
        }
    }

    /**
     *  Remove parts together with their rels and content-type overrides.
     *
     *  @param parts - the parts to remove
     */
    async remove(parts: string[]): Promise<void> {
        await removeParts(this.zip, parts)
    }

    /**
     *  The relationships of a part, targets resolved to the part names of
     *  this package (case-insensitively).
     *
     *  @param part - the owning part
     *  @param rels - its rels DOM
     *  @returns the relationships in document order
     */
    targets(part: string, rels: Document): Rel[] {
        return relTargets(part, rels, this.zip)
    }
}
