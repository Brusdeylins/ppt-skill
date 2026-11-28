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
import type { Document, Element } from "@xmldom/xmldom"
import { finding, notesOwner, partDefects, type Finding } from "./defects.js"
import { ASSET_PART, bytesOf, findPart, partText, relTargets, relsOwner, slideList, type Rel } from "./parts.js"
import { decodeXml, elements, parseXml } from "./xml.js"

/**  the element children of an element, by name  */
const childNames = (el: Element): string[] => {
    const names: string[] = []
    for (let n = el.firstChild; n !== null; n = n.nextSibling)
        if (n.nodeType === 1)
            names.push(n.nodeName)
    return names
}

/**  a deck as the rules see it: every XML part parsed once  */
interface Deck {
    zip: JSZip
    /**  the part names ([trash]/ excluded)  */
    names: Set<string>
    /**  the parsed XML parts  */
    docs: Map<string, Document>
    /**  the relationships by owning part ("" for the package)  */
    rels: Map<string, Rel[]>
}

/**  the slide parts of the deck  */
const slideNames = (deck: Deck): string[] =>
    [ ...deck.names ].filter((n) => /^ppt\/slides\/[^/]+\.xml$/.test(n))

/**  every XML part must be well-formed (parses the deck for the rules)  */
const parseDeck = async (zip: JSZip, findings: Finding[]): Promise<Deck> => {
    /*  parts under the reserved OPC folder [trash]/ are logically deleted
        (PowerPoint incremental save, co-authoring), carry no content type
        by design and are ignored by PowerPoint -- and by every rule  */
    const names = new Set(Object.keys(zip.files).filter((n) => !n.endsWith("/") && !n.startsWith("[trash]/")))
    const docs = new Map<string, Document>()
    for (const n of names)
        if (n.endsWith(".xml") || n.endsWith(".rels"))
            try {
                docs.set(n, parseXml(decodeXml(await (zip.file(n) as JSZip.JSZipObject).async("uint8array")), true))
            }
            catch (err) {
                findings.push(finding("malformed", n, `malformed XML (${String(err).slice(0, 80)})`))
            }
    const rels = new Map<string, Rel[]>()
    for (const [ n, doc ] of docs)
        if (n.endsWith(".rels"))
            rels.set(relsOwner(n), relTargets(relsOwner(n), doc, zip))
    return { zip, names, docs, rels }
}

/**  the triggers of every single part (engine/defects): duplicate or dead
     relationships, references to missing relationships, duplicate shape
     ids on slides  */
const partRules = (deck: Deck): Finding[] => {
    const owners = new Set([ ...deck.rels.keys(), ...[ ...deck.docs.keys() ].filter((n) => n.endsWith(".xml")) ])
    return [ ...owners ].flatMap((part) => partDefects({
        part: part === "" ? "_rels/.rels" : part, doc: deck.docs.get(part) ?? null, rels: deck.rels.get(part) ?? []
    }, (p) => deck.names.has(p)))
}

/**  content-type overrides: no dead parts, no duplicates; every part
     covered by a default or an override  */
const contentTypeRules = (deck: Deck): Finding[] => {
    const ct = deck.docs.get("[Content_Types].xml")
    if (ct === undefined)
        return []
    const findings: Finding[] = []
    const seen = new Set<string>()
    for (const o of elements(ct, "Override").map((el) => (el.getAttribute("PartName") ?? "").replace(/^\//, ""))) {
        if (findPart(deck.zip, o) === null)
            findings.push(finding("ct-dead", "[Content_Types].xml", `override for missing part /${o}`))
        if (seen.has(o.toLowerCase()))
            findings.push(finding("ct-dup", "[Content_Types].xml", `duplicate override /${o}`))
        seen.add(o.toLowerCase())
    }
    const defaults = new Set(elements(ct, "Default").map((d) => (d.getAttribute("Extension") ?? "").toLowerCase()))
    for (const n of deck.names)
        if (n !== "[Content_Types].xml" && !seen.has(n.toLowerCase()) && !defaults.has(n.split(".").pop()?.toLowerCase() ?? ""))
            findings.push(finding("ct-none", n, "no content type (neither default nor override)"))
    return findings
}

/**  the slide list: every sldId unique and resolving to its own slide part,
     every slide relationship and slide part in the list; the schema order
     of the presentation; sections referencing existing slides  */
const slideListRules = (deck: Deck): Finding[] => {
    const pres = deck.docs.get("ppt/presentation.xml")
    const root = pres?.documentElement ?? null
    if (pres === undefined || root === null)
        return []
    const findings: Finding[] = []
    const presRels = deck.rels.get("ppt/presentation.xml") ?? []
    const sldIds = slideList(deck.zip, pres, presRels).map((s) => ({ ...s, rid: s.el.getAttribute("r:id") ?? "", id: String(s.id) }))
    const slides: string[] = []
    const ids = new Set<string>()
    for (const { id, rid, part } of sldIds) {
        if (part === null || !deck.names.has(part))
            findings.push(finding("sldid-unresolved", "ppt/presentation.xml", `sldId ${rid} does not resolve to a slide part`))
        else if (slides.includes(part))
            findings.push(finding("sldid-twice", "ppt/presentation.xml", `sldId ${rid} lists slide ${part} a second time`))
        else
            slides.push(part)
        if (ids.has(id))
            findings.push(finding("sldid-dup", "ppt/presentation.xml", `sldId id ${id} is used twice`))
        ids.add(id)
    }
    const listed = new Set(sldIds.map((s) => s.rid))
    for (const r of presRels)
        if (r.kind === "slide" && !listed.has(r.id))
            findings.push(finding("slide-rel-dangling", "ppt/_rels/presentation.xml.rels", `dangling slide relationship ${r.id} (no sldId references it)`))
    for (const n of slideNames(deck))
        if (!slides.includes(n))
            findings.push(finding("slide-unlisted", n, "slide part not in the slide list"))

    /*  CT_Presentation is an ordered sequence: master id lists precede
        sldIdLst, sldSz precedes notesSz  */
    const seq = childNames(root)
    const before = (a: string, b: string): boolean =>
        !seq.includes(a) || !seq.includes(b) || seq.indexOf(a) < seq.indexOf(b)
    for (const lst of [ "p:notesMasterIdLst", "p:handoutMasterIdLst" ])
        if (!before(lst, "p:sldIdLst"))
            findings.push(finding("pres-order", "ppt/presentation.xml", `${lst.slice(2)} appears after sldIdLst (schema order violated)`))
    if (!before("p:sldSz", "p:notesSz"))
        findings.push(finding("pres-order", "ppt/presentation.xml", "sldSz appears after notesSz (schema order violated)"))
    for (const ref of elements(pres, "p14:sldId"))
        if (!ids.has(ref.getAttribute("id") ?? ""))
            findings.push(finding("section-ref", "ppt/presentation.xml", `section references removed slide id ${ref.getAttribute("id") ?? ""}`))
    return findings
}

/**  docProps/app.xml declares the real slide count and a matching title
     list: a deck claiming fewer slides than it has makes PowerPoint demand
     a repair (element content only, read as text)  */
const appRules = async (deck: Deck): Promise<Finding[]> => {
    if (!deck.names.has("docProps/app.xml"))
        return []
    const findings: Finding[] = []
    const app = await partText(deck.zip, "docProps/app.xml")
    const slideCount = elements(deck.docs.get("ppt/presentation.xml") ?? parseXml("<x/>"), "p:sldId").length
    const declared = /<Slides>(\d+)<\/Slides>/.exec(app)
    if (declared !== null && Number(declared[1]) !== slideCount)
        findings.push(finding("app-slides", "docProps/app.xml", `declares ${declared[1]} slides, deck has ${slideCount}`))
    const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(app)
    const tp = /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(app)
    if (hp !== null && tp !== null) {
        const sum = [ ...(hp[1] ?? "").matchAll(/<vt:i4>(\d+)<\/vt:i4>/g) ].reduce((a, m) => a + Number(m[1]), 0)
        const entries = [ ...(tp[1] ?? "").matchAll(/<vt:lpstr>/g) ].length
        if (sum !== entries)
            findings.push(finding("app-titles", "docProps/app.xml", `HeadingPairs sum ${sum} != TitlesOfParts ${entries}`))
    }
    return findings
}

/**  a notes part belongs to exactly one slide, the one it names  */
const notesRules = (deck: Deck): Finding[] => {
    const findings: Finding[] = []
    const users = new Map<string, string[]>()
    for (const slide of slideNames(deck))
        for (const r of deck.rels.get(slide) ?? [])
            if (r.kind === "notesSlide" && r.target !== null && deck.names.has(r.target)) {
                users.set(r.target, [ ...users.get(r.target) ?? [], slide ])
                const owner = notesOwner({ part: r.target, doc: deck.docs.get(r.target) ?? null, rels: deck.rels.get(r.target) ?? [] })
                if (owner !== slide)
                    findings.push(finding("notes-owner", slide, `its notes part ${r.target} belongs to ${owner ?? "no slide"}`))
            }
    for (const [ part, slides ] of users)
        if (slides.length > 1)
            findings.push(finding("notes-shared", part, `referenced by ${slides.length} slides (must be 1:1)`))
    return findings
}

/**  charts, SmartArt, embeddings and media are reachable from some part
     (orphans accumulate over applies and bloat or break the file)  */
const orphanRules = (deck: Deck): Finding[] => {
    const live = new Set([ ...deck.rels.values() ].flat().map((r) => r.target))
    return [ ...deck.names ].filter((n) => ASSET_PART.test(n) && !n.includes("/_rels/") && !live.has(n))
        .map((n) => finding("orphan", n, "orphan asset (not referenced from any part)"))
}

/**
 *  The repair triggers of an opened archive, structured.
 *
 *  @param zip - the loaded .pptx archive
 *  @param before - optional pre-apply archive to detect overwritten media
 *  @returns the findings; empty when the deck is clean
 */
export const findDefects = async (zip: JSZip, before?: JSZip): Promise<Finding[]> => {
    /*  without [Content_Types].xml there is no OPC package to check (an
        interleaved "piece" save stores it fragmented)  */
    if (zip.file("[Content_Types].xml") === null)
        return [ finding("ct-missing", "[Content_Types].xml", "part missing (not a flat OPC package)") ]
    const findings: Finding[] = []
    const deck = await parseDeck(zip, findings)
    findings.push(...partRules(deck), ...contentTypeRules(deck), ...slideListRules(deck),
        ...await appRules(deck), ...notesRules(deck), ...orphanRules(deck))

    /*  with the pre-apply archive: a media file kept under its name must keep
        its bytes, or another slide silently shows the new image  */
    if (before !== undefined)
        for (const n of Object.keys(before.files).filter((f) => f.startsWith("ppt/media/") && !f.endsWith("/"))) {
            const now = await bytesOf(zip, n)
            if (now !== null && !now.equals(await bytesOf(before, n) as Buffer))
                findings.push(finding("media-overwritten", n, "existing media file overwritten with different content"))
        }
    return findings
}

/**
 *  Validate an opened archive for the known PowerPoint repair triggers.
 *
 *  @param zip - the loaded .pptx archive
 *  @returns list of human-readable findings; empty when the deck is clean
 */
export const verifyArchive = async (zip: JSZip): Promise<string[]> =>
    (await findDefects(zip)).map((f) => f.text)

/**
 *  Validate a deck given its bytes.
 *
 *  @param bytes - the .pptx contents
 *  @returns list of findings; empty when clean
 */
export const verifyBytes = async (bytes: Buffer): Promise<string[]> =>
    await verifyArchive(await JSZip.loadAsync(bytes))

/**
 *  Validate a deck file on disk.
 *
 *  @param file - path of the .pptx to check
 *  @returns list of findings; empty when clean
 */
export const verifyFile = async (file: string): Promise<string[]> =>
    await verifyBytes(readFileSync(file))
