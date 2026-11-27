/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/guard: semantic self-check of an apply. verify.ts only proves the
**  output opens without a repair prompt; this module proves that every slide
**  still carries what the plan did NOT ask to change. Each apply rebuilds
**  the whole deck, so a part-naming or id collision in the engine can move
**  notes, images, charts or embedded objects to another slide while the
**  package stays perfectly valid. A fingerprint of each kept slide (notes
**  text plus the type and content hash of every related part) and of every
**  layout/master is compared before and after; a mismatch becomes a finding
**  and the apply fails atomically with E_INTEGRITY.
*/

import JSZip from "jszip"
import path from "node:path"
import { createHash } from "node:crypto"
import { parseXml, serializeXml } from "./xml.js"

/**  what the plan allows to change on one output slide  */
export interface SlideExpect {
    /**  source slide part in the input deck, null for slides from the seed  */
    source: string | null
    /**  sldId and title of the output slide, for readable findings  */
    label: string
    /**  whether this apply sets the slide's notes  */
    notes: boolean
    /**  related parts: "same" (untouched), "grow" (elements added) or "any"  */
    assets: "same" | "grow" | "any"
}

/**  rel types that never carry slide content of their own (or are
     checked separately), excluded from the asset fingerprint  */
const SKIP_RELS = [ "/slideLayout", "/notesSlide", "/hyperlink", "/slide" ]

/**  rel types a slide carries without referencing them by id in its XML  */
export const IMPLICIT_RELS = [ "/comments", "/themeOverride", "/vmlDrawing", "/slideUpdateInfo" ]

/**  content hash of a part. XML is hashed in a canonical form -- parsed
     and re-serialized, so quoting, entity and empty-element spellings do
     not matter (automizer re-serializes masters on every apply) -- with
     relationship ids and layout ids blanked, since the engine re-mints
     them without changing content  */
const digest = async (zip: JSZip, part: string): Promise<string> => {
    const file = zip.file(part)
    if (file === null)
        return "missing"
    const bytes = await file.async("nodebuffer")
    let data: Buffer | string = bytes
    if (/\.(xml|rels)$/.test(part)) {
        let xml = bytes.toString("utf8").replace(/^\uFEFF?(<\?xml[^>]*\?>)?\s*/, "")
        try {
            xml = serializeXml(parseXml(xml))
        }
        catch {
            /*  malformed XML is hashed as is (verify.ts reports it)  */
        }
        data = xml
            .replace(/>\s+</g, "><")
            .replace(/\br:(id|embed|link|pict|dm|lo|qs|cs)="[^"]*"/g, "r:$1=\"\"")
            .replace(/\brelId="[^"]*"/g, "relId=\"\"")
            .replace(/(<p:sldLayoutId\b[^>]*?\bid=")\d+"/g, "$1\"")
    }
    return createHash("sha1").update(data).digest("hex").slice(0, 12)
}

/**  type-tagged hashes of the parts a part relates to (charts one level
     deeper, so a swapped workbook or chart style is caught as well)  */
const relatedDigests = async (zip: JSZip, part: string, xml: string | null, depth: number): Promise<string[]> => {
    const relsPart = `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`
    const rels = await zip.file(relsPart)?.async("string")
    if (rels === undefined)
        return []
    const out: string[] = []
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const rel = m[0]
        if (rel.includes("TargetMode=\"External\""))
            continue
        const type = /Type="([^"]+)"/.exec(rel)?.[1] ?? ""
        if (SKIP_RELS.some((k) => type.endsWith(k)))
            continue
        const id = /Id="([^"]+)"/.exec(rel)?.[1] ?? ""
        if (xml !== null && !xml.includes(`"${id}"`) && !IMPLICIT_RELS.some((k) => type.endsWith(k)))
            continue
        const target = path.posix.normalize(path.posix.join(path.posix.dirname(part), /Target="([^"]+)"/.exec(rel)?.[1] ?? ""))
        let hash = await digest(zip, target)
        if (depth > 0)
            hash += `[${(await relatedDigests(zip, target, null, depth - 1)).join(",")}]`
        out.push(`${type.slice(type.lastIndexOf("/") + 1)}:${hash}`)
    }
    return out.sort()
}

/**  all a:t text of the slide's notes part, null when it has none  */
const notesText = async (zip: JSZip, slidePart: string): Promise<string | null> => {
    const rels = await zip.file(`ppt/slides/_rels/${path.posix.basename(slidePart)}.rels`)?.async("string") ?? ""
    const rel = [ ...rels.matchAll(/<Relationship\b[^>]*>/g) ].map((m) => m[0])
        .find((r) => /Type="[^"]*\/notesSlide"/.test(r))
    if (rel === undefined)
        return null
    const part = path.posix.normalize(path.posix.join("ppt/slides", /Target="([^"]+)"/.exec(rel)?.[1] ?? ""))
    const xml = await zip.file(part)?.async("string")
    return xml === undefined ? null
        : [ ...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g) ].map((t) => t[1]).join("\n")
}

/**  fingerprint of one slide: notes text and related-part hashes  */
const slidePrint = async (zip: JSZip, slidePart: string): Promise<{ notes: string | null, assets: string[] }> => {
    const xml = await zip.file(slidePart)?.async("string") ?? ""
    return { notes: await notesText(zip, slidePart), assets: await relatedDigests(zip, slidePart, xml, 1) }
}

/**  slide part paths in presentation order  */
const slideParts = async (zip: JSZip): Promise<string[]> => {
    const pres = await zip.file("ppt/presentation.xml")?.async("string") ?? ""
    const rels = await zip.file("ppt/_rels/presentation.xml.rels")?.async("string") ?? ""
    const target = new Map([ ...rels.matchAll(/<Relationship\b[^>]*>/g) ].map((m) =>
        [ /Id="([^"]+)"/.exec(m[0])?.[1] ?? "", /Target="([^"]+)"/.exec(m[0])?.[1] ?? "" ]))
    return [ ...pres.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g) ].map((m) =>
        path.posix.normalize(path.posix.join("ppt", target.get(m[1] as string) ?? "")))
}

/**  remove the entries of `b` from `a` (multiset difference)  */
const minus = (a: string[], b: string[]): string[] => {
    const rest = [ ...a ]
    for (const x of b) {
        const i = rest.indexOf(x)
        if (i >= 0)
            rest.splice(i, 1)
    }
    return rest
}

/**  split a related-part entry `type:hash[child,...]` into its parts  */
const parseEntry = (entry: string): { base: string, children: string[] } => {
    const i = entry.indexOf("[")
    return { base: entry.slice(0, i), children: entry.slice(i + 1, -1).split(",").filter((c) => c !== "") }
}

/**  outcome of the content comparison: `blocked` findings mean another
     slide's content would appear (an overwrite or swap) and fail the
     apply; `lost` items are content that would merely vanish (a known
     engine limitation) and are reported as warnings  */
export interface ContentCheck {
    blocked: string[]
    lost: { slide: number, what: string }[]
}

/**
 *  Compare the content of an apply's output against its input.
 *
 *  @param before - the input deck
 *  @param after - the output deck
 *  @param expects - one expectation per output slide, in output order
 *  @returns blocking findings and mere losses; both empty when nothing
 *           changed that the plan kept
 */
export const verifyContent = async (before: JSZip, after: JSZip, expects: SlideExpect[]): Promise<ContentCheck> => {
    const blocked: string[] = []
    const lost: { slide: number, what: string }[] = []
    const outParts = await slideParts(after)
    for (const [ i, exp ] of expects.entries()) {
        const outPart = outParts[i]
        if (exp.source === null || outPart === undefined)
            continue
        const was = await slidePrint(before, exp.source)
        const now = await slidePrint(after, outPart)
        if (!exp.notes && was.notes !== now.notes) {
            if (now.notes === null || now.notes === "")
                lost.push({ slide: i, what: "speaker notes" })
            else
                blocked.push(`slide ${exp.label}: speaker notes would change although no op set them`)
        }
        if (exp.assets === "any")
            continue

        /*  parts that vanished are a loss; a part that appeared or changed
            is foreign content -- unless it is a vanished part that only
            lost some of its own children (e.g. a picture inside SmartArt)  */
        const gone = minus(was.assets, now.assets)
        const added = exp.assets === "same" ? minus(now.assets, was.assets) : []
        for (const entry of added) {
            const a = parseEntry(entry)
            const k = gone.findIndex((g) => parseEntry(g).base === a.base
                && minus(a.children, parseEntry(g).children).length === 0)
            if (k < 0) {
                blocked.push(`slide ${exp.label}: images, charts or embedded objects would change `
                    + `although no op touched them (${was.assets.join(" ")} -> ${now.assets.join(" ")})`)
                break
            }
            const g = parseEntry(gone[k] as string)
            for (const child of minus(g.children, a.children))
                lost.push({ slide: i, what: `${child.slice(0, child.indexOf(":"))} inside ${a.base.slice(0, a.base.indexOf(":"))}` })
            gone.splice(k, 1)
        }
        for (const entry of gone)
            lost.push({ slide: i, what: entry.slice(0, entry.indexOf(":")) })
    }

    /*  layouts, masters and themes are never changed by an apply  */
    for (const part of Object.keys(before.files).filter((f) =>
        /^ppt\/(slideLayouts|slideMasters|theme)\/[^/]+\.xml$/.test(f) && after.file(f) !== null)) {
        const xmlBefore = await before.file(part)?.async("string") ?? ""
        const xmlAfter = await after.file(part)?.async("string") ?? ""
        const was = `${await digest(before, part)} ${(await relatedDigests(before, part, xmlBefore, 1)).join(" ")}`
        const now = `${await digest(after, part)} ${(await relatedDigests(after, part, xmlAfter, 1)).join(" ")}`
        if (was !== now)
            blocked.push(`${part}: layout/master content would change`)
    }
    return { blocked, lost }
}
