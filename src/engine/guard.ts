/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/guard: semantic self-checks of an apply. verify.ts only proves the
**  output opens without a repair prompt; this module proves that every
**  touched slide still carries what the plan did NOT ask to change, and
**  (byte guard) that no part of an untouched slide changed at all. A part-naming or id slip
**  while copying, transplanting or repairing could move notes, images,
**  charts or embedded objects to another slide while the package stays
**  perfectly valid. A fingerprint of each such slide (notes text plus the
**  type and content hash of every related part) and of every layout/master
**  is compared before and after; a mismatch becomes a finding and the apply
**  fails atomically with E_INTEGRITY.
*/

import JSZip from "jszip"
import { createHash } from "node:crypto"
import { RELS, bytesOf, optText, readRels, relIs, slideList } from "./parts.js"
import { decodeXml, elements, parseXml, serializeXml } from "./xml.js"

/**  what the plan allows to change on one output slide  */
export interface SlideExpect {
    /**  source slide part in the input deck; null for new slides and for
         untouched ones (the engine's byte guard covers those)  */
    source: string | null
    /**  sldId and title of the output slide, for readable findings  */
    label: string
    /**  whether this apply sets the slide's notes  */
    notes: boolean
    /**  related parts: "same" (untouched), "grow" (elements added) or "any"  */
    assets: "same" | "grow" | "any"
    /**  for "grow": what the planned elements add -- images by content
         hash (`digestBytes` of the file), plus generated PNG fallbacks of
         SVG images and charts by count  */
    grow?: { images: string[], fallbacks: number, charts: number }
    /**  whether an op replaces a whole text body or removes a table row,
         or a text.replace removed a picture-filled run -- which may drop an
         inline picture (a picture bullet, a cell fill) legitimately  */
    textReplaced?: boolean
    /**  images an op put in place of others (el.set image) or brought in
         with layout content (footer, slide number, date): the new image
         (`to`, content hash) must come, the old ones (`from`) may go --
         `to` null: a removed footer-band placeholder's images may go  */
    replaced?: { from: string[], to: string | null }[]
    /**  relationship ids of the source slide whose content a named repair
         removed (W_REPAIRED) -- left out of its fingerprint  */
    released?: string[]
    /**  parts of the slide's content a repair changed (W_REPAIRED): in the
         input (`before`, source names) and the output (`after`) -- they
         count by kind, and the relationships their repair released
         (`released`) leave their children  */
    repaired?: { before: string[], after: string[], released: string[] }
}

/**  a part's text in canonical form (parsed and re-serialized), so that
     attribute quoting and entity spellings never matter to a regex  */
const canonText = async (zip: JSZip, part: string): Promise<string | null> => {
    const text = await optText(zip, part)
    if (text === null)
        return null
    try {
        return serializeXml(parseXml(text))
    }
    catch {
        return text
    }
}

/**  content hash of raw bytes, as `digest` computes it for binary parts  */
export const digestBytes = (bytes: Buffer): string =>
    createHash("sha1").update(bytes).digest("hex").slice(0, 12)

/**  content hashes per archive: layouts and shared parts are hashed once  */
const digests = new WeakMap<JSZip, Map<string, Promise<string>>>()

/**  content hash of a part. XML is hashed in a canonical form -- parsed
     and re-serialized, so quoting, entity and empty-element spellings do
     not matter (a touched part is re-serialized) -- with relationship ids
     and layout ids blanked, since a copied slide carries them under other
     names, and slide-jump links left out (links to removed slides go)  */
const digest = async (zip: JSZip, part: string): Promise<string> => {
    let memo = digests.get(zip)
    if (memo === undefined) {
        memo = new Map()
        digests.set(zip, memo)
    }
    let hash = memo.get(part)
    if (hash === undefined) {
        hash = computeDigest(zip, part)
        memo.set(part, hash)
    }
    return await hash
}
const computeDigest = async (zip: JSZip, part: string): Promise<string> => {
    const bytes = await bytesOf(zip, part)
    if (bytes === null)
        return "missing"
    let data: Buffer | string = bytes
    if (/\.(xml|rels)$/.test(part)) {
        let xml = decodeXml(bytes).replace(/^(<\?xml[^>]*\?>)?\s*/, "")
        try {
            const doc = parseXml(xml)

            /*  slide-jump links to removed slides are dropped by the engine  */
            for (const tag of [ "a:hlinkClick", "a:hlinkHover" ])
                for (const el of elements(doc, tag))
                    if ((el.getAttribute("action") ?? "").includes("hlinksldjump"))
                        el.parentNode?.removeChild(el)
            xml = serializeXml(doc)
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
const relatedDigests = async (zip: JSZip, part: string, xml: string | null, depth: number,
    skip: string[] = [], repaired: string[] = [], released: string[] = []): Promise<string[]> => {
    const out: string[] = []
    const seen = new Set<string>(skip)
    for (const rel of await readRels(zip, part)) {
        const { el, id, kind, target } = rel

        /*  a duplicate id resolves to its first relationship (as repaired)  */
        if (relIs(rel, RELS.unprinted) || seen.has(id))
            continue
        seen.add(id)
        if (xml !== null && !xml.includes(`"${id}"`) && !relIs(rel, RELS.implicit))
            continue

        /*  a linked file (picture, OLE object, video) counts by its target  */
        if (target === null) {
            out.push(`${kind}:ext:${el.getAttribute("Target") ?? ""}`)
            continue
        }

        /*  a dead target carries no content that could be lost; a repaired
            one counts by its kind and its children  */
        if (zip.file(target) === null)
            continue
        const fixed = repaired.includes(target)
        let hash = fixed ? "repaired" : await digest(zip, target)
        if (depth > 0)
            hash += `[${(await relatedDigests(zip, target, null, depth - 1, fixed ? released : [], repaired)).join(",")}]`
        out.push(`${kind}:${hash}`)
    }
    return out.sort()
}

/**  all a:t text of the slide's notes part, null when it has none  */
const notesText = async (zip: JSZip, slidePart: string): Promise<string | null> => {
    const notes = (await readRels(zip, slidePart)).find((r) => r.kind === "notesSlide")?.target ?? null
    const xml = notes === null ? null : await canonText(zip, notes)
    return xml === null ? null
        : [ ...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g) ].map((t) => t[1]).join("\n")
}

/**  fingerprint of one slide: notes text and related-part hashes  */
const slidePrint = async (zip: JSZip, slidePart: string, skip: string[] = [],
    repaired: string[] = [], released: string[] = []): Promise<{ notes: string | null, assets: string[] }> => {
    const xml = await canonText(zip, slidePart) ?? ""
    return { notes: await notesText(zip, slidePart), assets: await relatedDigests(zip, slidePart, xml, 1, skip, repaired, released) }
}

/**  slide part paths in presentation order  */
const slideParts = async (zip: JSZip): Promise<string[]> => {
    const pres = await canonText(zip, "ppt/presentation.xml")
    return pres === null ? []
        : slideList(zip, parseXml(pres), await readRels(zip, "ppt/presentation.xml")).map((s) => s.part ?? "")
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
     slide's content would appear (an overwrite or swap); `lost` items are
     content that would vanish. Both fail the apply.  */
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
 *  @param repaired - package parts the engine repaired (W_REPAIRED), among
 *                    them layouts, masters and themes
 *                    exempt from the layout check
 *  @returns blocking findings and mere losses; both empty when nothing
 *           changed that the plan kept
 */
export const verifyContent = async (before: JSZip, after: JSZip, expects: SlideExpect[],
    repaired: Set<string> = new Set()): Promise<ContentCheck> => {
    const blocked: string[] = []
    const lost: { slide: number, what: string }[] = []
    const outParts = await slideParts(after)
    for (const [ i, exp ] of expects.entries()) {
        const outPart = outParts[i]
        if (exp.source === null || outPart === undefined)
            continue
        const was = await slidePrint(before, exp.source, exp.released, exp.repaired?.before, exp.repaired?.released)
        const now = await slidePrint(after, outPart, [], exp.repaired?.after, exp.repaired?.released)
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
            lost some of its own children (e.g. a picture inside SmartArt).
            Replaced images count as a multiset: the new ones (`to`) join
            what the slide had and must come, the old ones (`from`) may go
            -- so an image moved between pictures or two exchanged images
            are checked in any op order  */
        const image = (hash: string): string => `image:${hash}[]`
        const replaced = exp.replaced ?? []
        const comes = replaced.flatMap((swap) => swap.to === null ? [] : [ image(swap.to) ])
        const expected = [ ...was.assets, ...comes ]
        const added = minus(now.assets, expected)
        let gone = minus(expected, now.assets)
        gone = minus(gone, replaced.flatMap((swap) => swap.from.map(image)))
        for (const entry of minus(comes, minus(comes, gone))) {
            blocked.push(`slide ${exp.label}: the image an op put in place (${entry.slice(6, -2)}) is missing`)
            gone.splice(gone.indexOf(entry), 1)
        }

        /*  "grow": the planned elements account for their images (by
            content), SVG fallbacks and charts; anything beyond is foreign  */
        if (exp.assets === "grow" && exp.grow !== undefined) {
            const images = [ ...exp.grow.images ]
            let fallbacks = exp.grow.fallbacks
            let charts = exp.grow.charts
            for (let j = added.length - 1; j >= 0; j--) {
                const entry = added[j] as string
                const planned = images.indexOf(entry.replace(/^image:([0-9a-f]+)\[\]$/, "$1"))
                if (entry.startsWith("image:") && planned >= 0)
                    images.splice(planned, 1)
                else if (entry.startsWith("image:") && fallbacks > 0)
                    fallbacks--
                else if (entry.startsWith("chart:") && charts > 0)
                    charts--
                else
                    continue
                added.splice(j, 1)
            }
        }
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
        for (const entry of gone) {
            const type = entry.slice(0, entry.indexOf(":"))
            if (!(exp.textReplaced === true && type === "image"))
                lost.push({ slide: i, what: type })
        }
    }

    /*  layouts, masters and themes are never changed by an apply  */
    for (const part of Object.keys(before.files).filter((f) =>
        /^ppt\/(slideLayouts|slideMasters|theme)\/[^/]+\.xml$/.test(f) && after.file(f) !== null && !repaired.has(f))) {
        const xmlBefore = await canonText(before, part) ?? ""
        const xmlAfter = await canonText(after, part) ?? ""
        const was = `${await digest(before, part)} ${(await relatedDigests(before, part, xmlBefore, 1)).join(" ")}`
        const now = `${await digest(after, part)} ${(await relatedDigests(after, part, xmlAfter, 1)).join(" ")}`
        if (was !== now)
            blocked.push(`${part}: layout/master content would change`)
    }
    return { blocked, lost }
}

/**
 *  The byte guard: every part of an untouched slide keeps its exact bytes
 *  (it is never written at all).
 *
 *  @param zip - the edited package
 *  @param untouched - the parts of untouched slides with their input entry
 *  @param after - the output deck
 *  @returns findings, empty when every such part is unchanged
 */
export const byteGuard = async (zip: JSZip, untouched: Map<string, JSZip.JSZipObject | null>, after: JSZip): Promise<string[]> => {
    const changed: string[] = []
    for (const [ name, file ] of untouched) {
        const now = await bytesOf(after, name)
        if (zip.file(name) !== file || file === null || now === null || !now.equals(await bytesOf(zip, name) as Buffer))
            changed.push(`${name}: part of an untouched slide would change`)
    }
    return changed
}
