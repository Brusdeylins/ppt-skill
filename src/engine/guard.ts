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

/**  content hash of a part; for XML the declaration, byte-order mark,
     whitespace between tags, relationship ids and layout ids are dropped
     first, since the engine re-serializes and re-mints them on every apply
     without changing content  */
const digest = async (zip: JSZip, part: string): Promise<string> => {
    const file = zip.file(part)
    if (file === null)
        return "missing"
    const bytes = await file.async("nodebuffer")
    const data = /\.(xml|rels)$/.test(part)
        ? bytes.toString("utf8")
            .replace(/^\uFEFF?(<\?xml[^>]*\?>)?\s*/, "")
            .replace(/>\s+</g, "><")
            .replace(/\br:(id|embed|link|pict|dm|lo|qs|cs)="[^"]*"/g, "r:$1=\"\"")
            .replace(/(<p:sldLayoutId\b[^>]*?\bid=")\d+"/g, "$1\"")
        : bytes
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

/**  is every entry of `inner` contained in `outer` (as multisets)  */
const contains = (outer: string[], inner: string[]): boolean => {
    const rest = [ ...outer ]
    for (const x of inner) {
        const i = rest.indexOf(x)
        if (i < 0)
            return false
        rest.splice(i, 1)
    }
    return true
}

/**
 *  Compare the content of an apply's output against its input.
 *
 *  @param before - the input deck
 *  @param after - the output deck
 *  @param expects - one expectation per output slide, in output order
 *  @returns findings; empty when nothing changed that the plan kept
 */
export const verifyContent = async (before: JSZip, after: JSZip, expects: SlideExpect[]): Promise<string[]> => {
    const findings: string[] = []
    const outParts = await slideParts(after)
    for (const [ i, exp ] of expects.entries()) {
        const outPart = outParts[i]
        if (exp.source === null || outPart === undefined)
            continue
        const was = await slidePrint(before, exp.source)
        const now = await slidePrint(after, outPart)
        if (!exp.notes && was.notes !== now.notes)
            findings.push(`slide ${exp.label}: speaker notes would change although no op set them`)
        if ((exp.assets === "same" && was.assets.join() !== now.assets.join())
            || (exp.assets === "grow" && !contains(now.assets, was.assets)))
            findings.push(`slide ${exp.label}: images, charts or embedded objects would change `
                + `although no op touched them (${was.assets.join(" ")} -> ${now.assets.join(" ")})`)
    }

    /*  layouts, masters and themes are never changed by an apply  */
    for (const part of Object.keys(before.files).filter((f) =>
        /^ppt\/(slideLayouts|slideMasters|theme)\/[^/]+\.xml$/.test(f) && after.file(f) !== null)) {
        const xmlBefore = await before.file(part)?.async("string") ?? ""
        const xmlAfter = await after.file(part)?.async("string") ?? ""
        const was = `${await digest(before, part)} ${(await relatedDigests(before, part, xmlBefore, 1)).join(" ")}`
        const now = `${await digest(after, part)} ${(await relatedDigests(after, part, xmlAfter, 1)).join(" ")}`
        if (was !== now)
            findings.push(`${part}: layout/master content would change`)
    }
    return findings
}
