/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/parts: neutral zip-package helpers shared by the write-path
**  (post.ts) and the seed builder (seed.ts). Kept separate so the seed
**  builder does not depend on the write module just to reuse one utility.
*/

import type JSZip from "jszip"
import path from "node:path"
import { PptcError } from "../infra/errors.js"
import { elements, parseXml, serializeXml } from "./xml.js"

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
 *  written by some generators) to the relative form the engine and
 *  pptx-automizer resolve. External targets stay as they are.
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
