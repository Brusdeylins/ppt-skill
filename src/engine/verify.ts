/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/verify: static validation of a written deck against every known
**  PowerPoint "repair" trigger. This SHIPS with the engine (and therefore
**  with the skill bundle), so `apply` self-checks its own output on any
**  machine -- a structurally fine but repair-prompting file is never written.
**  The triggers of single parts come from engine/defects, the rule set the
**  engine repairs its input by; the package-level rules (slide list, content
**  types, notes ownership, app properties, orphans) live here.
*/

import JSZip from "jszip"
import { readFileSync } from "node:fs"
import { DOMParser, type Document, type Element } from "@xmldom/xmldom"
import { notesOwner, partDefects, type PartView } from "./defects.js"
import { ASSET_PART, bytesOf, findPart, relTargets, relsOwner, type Rel } from "./parts.js"
import { elements } from "./xml.js"

/**  the element children of an element, by name  */
const childNames = (el: Element): string[] => {
    const names: string[] = []
    for (let n = el.firstChild; n !== null; n = n.nextSibling)
        if (n.nodeType === 1)
            names.push(n.nodeName)
    return names
}

/**
 *  Validate an opened archive for the known PowerPoint repair triggers.
 *
 *  @param zip - the loaded .pptx archive
 *  @returns list of human-readable findings; empty when the deck is clean
 */
export const verifyArchive = async (zip: JSZip): Promise<string[]> => {
    /*  without [Content_Types].xml there is no OPC package to check (an
        interleaved "piece" save stores it fragmented) -- report a single
        finding instead of crashing the part-based checks below  */
    if (zip.file("[Content_Types].xml") === null)
        return [ "[Content_Types].xml: part missing (not a flat OPC package)" ]
    const findings: string[] = []

    /*  parts under the reserved OPC folder [trash]/ are logically deleted
        (left by PowerPoint incremental save / OneDrive co-authoring), carry
        no content type by design and are ignored by PowerPoint -- exclude
        them from every check (the engine drops them from the output)  */
    const names = new Set(Object.keys(zip.files)
        .filter((n) => !n.endsWith("/") && !n.startsWith("[trash]/")))
    const exists = (part: string): boolean => names.has(part)

    /*  1. every XML part must be well-formed; each is parsed once for all
        rules below  */
    const docs = new Map<string, Document>()
    for (const n of names)
        if (n.endsWith(".xml") || n.endsWith(".rels"))
            try {
                const parser = new DOMParser({
                    onError: (_level, msg): void => { throw new Error(msg) }
                })
                const text = (await (zip.file(n) as JSZip.JSZipObject).async("string")).replace(/^\uFEFF/, "")
                docs.set(n, parser.parseFromString(text, "application/xml"))
            }
            catch (err) {
                findings.push(`${n}: malformed XML (${String(err).slice(0, 80)})`)
            }
    const rels = new Map<string, Rel[]>()
    for (const [ n, doc ] of docs)
        if (n.endsWith(".rels"))
            rels.set(relsOwner(n), relTargets(relsOwner(n), doc, zip))

    /*  2. the triggers of every single part: duplicate or dead
        relationships, references to missing relationships, duplicate shape
        ids on slides  */
    const owners = new Set([ ...rels.keys(), ...[ ...docs.keys() ].filter((n) => n.endsWith(".xml")) ])
    for (const part of owners) {
        const view: PartView = { part: part === "" ? "_rels/.rels" : part, doc: docs.get(part) ?? null, rels: rels.get(part) ?? [] }
        findings.push(...partDefects(view, exists))
    }

    /*  3. content-type overrides: no dead parts, no duplicates; every part
        covered by a default or an override  */
    const ct = docs.get("[Content_Types].xml")
    if (ct !== undefined) {
        const overrides = elements(ct, "Override").map((o) => (o.getAttribute("PartName") ?? "").replace(/^\//, ""))
        const seen = new Set<string>()
        for (const o of overrides) {
            if (findPart(zip, o) === null)
                findings.push(`[Content_Types].xml: override for missing part /${o}`)
            if (seen.has(o.toLowerCase()))
                findings.push(`[Content_Types].xml: duplicate override /${o}`)
            seen.add(o.toLowerCase())
        }
        const defaults = new Set(elements(ct, "Default").map((d) => (d.getAttribute("Extension") ?? "").toLowerCase()))
        for (const n of names)
            if (n !== "[Content_Types].xml" && !seen.has(n.toLowerCase())
                && !defaults.has(n.split(".").pop()?.toLowerCase() ?? ""))
                findings.push(`${n}: no content type (neither default nor override)`)
    }

    /*  4. the slide list: every sldId resolves to its own slide part, every
        slide relationship and slide part is in the list  */
    const pres = docs.get("ppt/presentation.xml")
    const presRels = new Map((rels.get("ppt/presentation.xml") ?? []).map((r) => [ r.id, r ]))
    const sldIds = pres === undefined ? [] : elements(pres, "p:sldId")
        .map((el) => ({ id: el.getAttribute("id") ?? "", rid: el.getAttribute("r:id") ?? "" }))
    const slides: string[] = []
    for (const { rid } of sldIds) {
        const target = presRels.get(rid)?.target ?? null
        if (target === null || !exists(target) || !(presRels.get(rid)?.type ?? "").endsWith("/slide"))
            findings.push(`presentation.xml: sldId ${rid} does not resolve to a slide part`)
        else if (slides.includes(target))
            findings.push(`presentation.xml: sldId ${rid} lists slide ${target} a second time`)
        else
            slides.push(target)
    }
    const listed = new Set(sldIds.map((s) => s.rid))
    for (const r of presRels.values())
        if (r.type.endsWith("/slide") && !listed.has(r.id))
            findings.push(`presentation.xml.rels: dangling slide relationship ${r.id} (no sldId references it)`)
    for (const n of names)
        if (/^ppt\/slides\/[^/]+\.xml$/.test(n) && !slides.includes(n))
            findings.push(`${n}: slide part not in the slide list`)

    /*  CT_Presentation is an ordered sequence: master id lists must precede
        sldIdLst, sldSz must precede notesSz  */
    const presRoot = pres?.documentElement ?? null
    const seq = presRoot === null ? [] : childNames(presRoot)
    const before = (a: string, b: string): boolean =>
        !seq.includes(a) || !seq.includes(b) || seq.indexOf(a) < seq.indexOf(b)
    for (const lst of [ "p:notesMasterIdLst", "p:handoutMasterIdLst" ])
        if (!before(lst, "p:sldIdLst"))
            findings.push(`presentation.xml: ${lst.slice(2)} appears after sldIdLst (schema order violated)`)
    if (!before("p:sldSz", "p:notesSz"))
        findings.push("presentation.xml: sldSz appears after notesSz (schema order violated)")

    /*  section slide references must resolve to existing sldIds  */
    const realIds = new Set(sldIds.map((s) => s.id))
    for (const ref of pres === undefined ? [] : elements(pres, "p14:sldId"))
        if (!realIds.has(ref.getAttribute("id") ?? ""))
            findings.push(`presentation.xml: section references removed slide id ${ref.getAttribute("id") ?? ""}`)

    /*  5. docProps/app.xml must declare the real slide count and a matching
        title list: a deck claiming fewer slides than it has makes PowerPoint
        demand a repair on open  */
    if (names.has("docProps/app.xml")) {
        const app = await (zip.file("docProps/app.xml") as JSZip.JSZipObject).async("string")
        const declared = /<Slides>(\d+)<\/Slides>/.exec(app)
        if (declared !== null && Number(declared[1]) !== sldIds.length)
            findings.push(`docProps/app.xml: declares ${declared[1]} slides, deck has ${sldIds.length}`)
        const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(app)
        const tp = /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(app)
        if (hp !== null && tp !== null) {
            const sum = [ ...(hp[1] ?? "").matchAll(/<vt:i4>(\d+)<\/vt:i4>/g) ].reduce((a, m) => a + Number(m[1]), 0)
            const entries = [ ...(tp[1] ?? "").matchAll(/<vt:lpstr>/g) ].length
            if (sum !== entries)
                findings.push(`docProps/app.xml: HeadingPairs sum ${sum} != TitlesOfParts ${entries}`)
        }
    }

    /*  6. a notes part belongs to exactly one slide, the one its
        back-reference names  */
    const notesOf = new Map<string, string[]>()
    for (const slide of [ ...names ].filter((n) => /^ppt\/slides\/[^/]+\.xml$/.test(n)))
        for (const r of rels.get(slide) ?? [])
            if (r.type.endsWith("/notesSlide") && r.target !== null && exists(r.target)) {
                notesOf.set(r.target, [ ...notesOf.get(r.target) ?? [], slide ])
                const owner = notesOwner({ part: r.target, doc: docs.get(r.target) ?? null, rels: rels.get(r.target) ?? [] })
                if (owner !== slide)
                    findings.push(`${slide}: its notes part ${r.target} belongs to ${owner ?? "no slide"}`)
            }
    for (const [ part, users ] of notesOf)
        if (users.length > 1)
            findings.push(`notesSlide ${part} is referenced by ${users.length} slides (must be 1:1)`)

    /*  7. charts, SmartArt, embeddings and media must be reachable from some
        part (orphans accumulate over applies and bloat or break the file)  */
    const live = new Set([ ...rels.values() ].flat().map((r) => r.target))
    for (const n of names)
        if (ASSET_PART.test(n) && !n.includes("/_rels/") && !live.has(n))
            findings.push(`${n}: orphan asset (not referenced from any part)`)

    return findings
}

/**
 *  Validate an opened deck, optionally against the deck it was made from.
 *
 *  @param zip - the deck
 *  @param before - optional pre-apply archive to detect overwritten media
 *  @returns list of findings; empty when clean
 */
export const verifyZip = async (zip: JSZip, before?: JSZip): Promise<string[]> => {
    const findings = await verifyArchive(zip)

    /*  with the pre-apply archive: a media file kept under its name must keep
        its bytes, or another slide silently shows the new image  */
    if (before !== undefined)
        for (const n of Object.keys(before.files).filter((f) => f.startsWith("ppt/media/") && !f.endsWith("/"))) {
            const now = await bytesOf(zip, n)
            if (now !== null && !now.equals(await bytesOf(before, n) as Buffer))
                findings.push(`${n}: existing media file overwritten with different content`)
        }
    return findings
}

/**
 *  Validate a deck given its bytes.
 *
 *  @param bytes - the .pptx contents
 *  @returns list of findings; empty when clean
 */
export const verifyBytes = async (bytes: Buffer): Promise<string[]> =>
    await verifyZip(await JSZip.loadAsync(bytes))

/**
 *  Validate a deck file on disk.
 *
 *  @param file - path of the .pptx to check
 *  @returns list of findings; empty when clean
 */
export const verifyFile = async (file: string): Promise<string[]> =>
    await verifyBytes(readFileSync(file))
