/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  test/util/fingerprint: an engine-agnostic fingerprint of a deck, per
**  slide in order: id, title, layout, visibility, notes, section, shapes
**  (generated names with their UUID suffix masked), the content of related
**  images, charts, embedded objects, tags and comments, and link targets.
**  Part names and relationship ids are left out -- they are the engine's
**  business, the content is the user's.
*/

import JSZip from "jszip"
import path from "node:path"
import { readFileSync } from "node:fs"
import { createHash } from "node:crypto"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { parseXml, serializeXml } from "../../src/engine/xml.js"

/**  short content hash  */
const hash = (data: Uint8Array | string): string =>
    createHash("sha1").update(data).digest("hex").slice(0, 10)

/**  mask the UUID suffix of generated shape names  */
const maskName = (name: string): string =>
    name.replace(/-?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/, "-*")

/**  resolve a relationship target to a part name  */
const resolve = (base: string, target: string): string =>
    target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join(base, target))

/**  XML in canonical form: parsed and re-serialized, whitespace between
     tags collapsed  */
const canon = (xml: string): string =>
    serializeXml(parseXml(xml.trim())).replace(/>\s+</g, "><")

/**  summary of a chart: chart types, series names and cached values  */
const chartSummary = (xml: string): string =>
    [ ...xml.matchAll(/<c:(\w+Chart)>/g) ].map((m) => m[1]).join("+")
        + ":" + [ ...xml.matchAll(/<c:v>([^<]*)<\/c:v>/g) ].map((m) => m[1]).join("|")

/**  related content of one part, recursing into charts and SmartArt  */
const related = async (zip: JSZip, part: string, depth: number): Promise<string[]> => {
    const rels = await zip.file(`${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`)?.async("string")
    if (rels === undefined)
        return []
    const out: string[] = []
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const type = (/Type="([^"]+)"/.exec(m[0])?.[1] ?? "").replace(/^.*\//, "")
        const target = /Target="([^"]+)"/.exec(m[0])?.[1] ?? ""
        if ([ "slideLayout", "notesSlide", "slide", "notesMaster", "slideMaster", "theme" ].includes(type))
            continue
        if (m[0].includes("TargetMode=\"External\"")) {
            out.push(`${type}:ext:${target}`)
            continue
        }
        const file = zip.file(resolve(path.posix.dirname(part), target))
        if (file === null) {
            out.push(`${type}:missing`)
            continue
        }
        const bytes = await file.async("uint8array")
        const name = resolve(path.posix.dirname(part), target)
        if (type === "chart")
            out.push(`chart:${chartSummary(Buffer.from(bytes).toString("utf8"))}`)
        else if (/\.xml$/.test(name))
            out.push(`${type}:${hash(Buffer.from(bytes).toString("utf8").replace(/\br:(id|embed|link|dm|lo|qs|cs)="[^"]*"/g, "").replace(/\brelId="[^"]*"/g, ""))}`)
        else if (bytes[0] === 0x50 && bytes[1] === 0x4B) {
            /*  an embedded package (workbook): its entries, not the zip
                container, whose timestamps differ per generation  */
            const inner = await JSZip.loadAsync(bytes)
            const entries: string[] = []
            for (const f of Object.keys(inner.files).filter((n) => !n.endsWith("/")).sort())
                entries.push(`${f}:${canon((await (inner.file(f) as JSZip.JSZipObject).async("string"))
                    .replace(/(<dcterms:(?:created|modified)[^>]*>)[^<]*/g, "$1"))}`)
            out.push(`${type}:pkg:${hash(entries.join("\n"))}`)
        }
        else
            out.push(`${type}:${hash(bytes)}`)
        if (depth > 0)
            out.push(...(await related(zip, name, depth - 1)).map((r) => `  ${type}>${r}`))
    }
    return out.sort()
}

/**  the fingerprint of one slide  */
export interface SlidePrint {
    id: number
    title: string | null
    layout: string
    hidden: boolean
    notes: string | null
    section: string | null
    shapes: string[]
    related: string[]
}

/**
 *  Fingerprint a deck file.
 *
 *  @param deck - path of the .pptx
 *  @returns one print per slide, in slide order
 */
export const fingerprint = async (deck: string): Promise<SlidePrint[]> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const sectionOf = new Map<number, string>()
    for (const m of pres.matchAll(/<p14:section name="([^"]*)"[^>]*>([\s\S]*?)<\/p14:section>/g))
        for (const id of (m[2] as string).matchAll(/<p14:sldId id="(\d+)"/g))
            sectionOf.set(Number(id[1]), m[1] as string)
    const out: SlidePrint[] = []
    for (const s of state.slides)
        out.push({
            id: s.id,
            title: s.title,
            layout: s.layoutName,
            hidden: s.hidden,
            notes: s.notes,
            section: sectionOf.get(s.id) ?? null,
            shapes: s.shapes.map((sh) => [ maskName(sh.name), sh.type, sh.placeholderIdx ?? "",
                JSON.stringify(sh.frame), sh.text ?? "", JSON.stringify(sh.table ?? null) ].join("|")),
            related: await related(zip, s.part, 1)
        })
    return out
}

/**
 *  The custom and core properties of a deck, for comparisons.
 *
 *  @param deck - path of the .pptx
 *  @returns core.xml text content and the custom properties
 */
export const propsOf = async (deck: string): Promise<{ core: string, custom: Record<string, string> }> => {
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const core = await zip.file("docProps/core.xml")?.async("string") ?? ""
    const state = await readDeckState(await DeckArchive.open(deck))
    return {
        core: [ ...core.matchAll(/<(dc:title|dc:creator|dc:subject|cp:keywords|cp:category|dc:description)>([^<]*)</g) ]
            .map((m) => `${m[1]}=${m[2]}`).sort().join(";"),
        custom: state.customProps
    }
}
