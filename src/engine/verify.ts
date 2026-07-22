/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/verify: static validation of a written deck against every known
**  PowerPoint "repair" trigger. This SHIPS with the engine (and therefore
**  with the skill bundle), so `apply` self-checks its own output on any
**  machine -- a structurally fine but repair-prompting file is never written.
**  The post-pass repairs these triggers; verify is the safety net that turns
**  any residual corruption into a clean, atomic failure instead of a broken
**  .pptx the user only discovers in PowerPoint.
**
**  LOCKSTEP: the trigger list here mirrors what engine/post.ts repairs --
**  extend both sides together (see the matching note there).
*/

import JSZip from "jszip"
import path from "node:path"
import { readFileSync } from "node:fs"
import { DOMParser } from "@xmldom/xmldom"

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
        return ["[Content_Types].xml: part missing (not a flat OPC package)"]
    const findings: string[] = []
    /*  parts under the reserved OPC folder [trash]/ are logically deleted
        (left by PowerPoint incremental save / OneDrive co-authoring), carry
        no content type by design and are ignored by PowerPoint -- exclude
        them from every check (the post-pass drops them from the output)  */
    const names = new Set(Object.keys(zip.files)
        .filter((n) => !n.endsWith("/") && !n.startsWith("[trash]/")))
    const text = async (part: string): Promise<string> =>
        await (zip.file(part) as JSZip.JSZipObject).async("string")

    /*  1. every XML part must be well-formed  */
    for (const n of names)
        if (n.endsWith(".xml") || n.endsWith(".rels"))
            try {
                const parser = new DOMParser({
                    onError: (_level, msg): void => { throw new Error(msg) }
                })
                parser.parseFromString(await text(n), "application/xml")
            }
            catch (err) {
                findings.push(`${n}: malformed XML (${String(err).slice(0, 80)})`)
            }

    /*  2. cNvPr shape ids must be unique per slide  */
    for (const n of names)
        if (/^ppt\/slides\/slide\d+\.xml$/.test(n)) {
            const ids = [...(await text(n)).matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1])
            const dup = [...new Set(ids.filter((id) => ids.indexOf(id) !== ids.lastIndexOf(id)))]
            if (dup.length > 0)
                findings.push(`${n}: duplicate cNvPr ids ${dup.join(",")}`)
        }

    /*  3. every internal relationship target must exist; rel ids unique  */
    for (const n of names)
        if (n.endsWith(".rels")) {
            const relIds = [...(await text(n)).matchAll(/<Relationship Id="([^"]+)"/g)].map((m) => m[1])
            const dupIds = [...new Set(relIds.filter((id) => relIds.indexOf(id) !== relIds.lastIndexOf(id)))]
            if (dupIds.length > 0)
                findings.push(`${n}: duplicate relationship ids ${dupIds.join(",")}`)
            const base = path.posix.dirname(path.posix.dirname(n))
            for (const m of (await text(n)).matchAll(/<Relationship [^>]*?\/>/g)) {
                if (m[0].includes("TargetMode=\"External\""))
                    continue
                const target = /Target="([^"]+)"/.exec(m[0])?.[1] ?? ""
                const full = path.posix.normalize(path.posix.join(base, target))
                if (!names.has(full))
                    findings.push(`${n}: dead relationship target ${target}`)
            }
        }

    /*  4. content-type overrides: no dead parts, no duplicates  */
    const ct = await text("[Content_Types].xml")
    const overrides = [...ct.matchAll(/<Override PartName="([^"]+)"/g)].map((m) => m[1] as string)
    for (const o of new Set(overrides)) {
        if (!names.has(o.replace(/^\//, "")))
            findings.push(`[Content_Types].xml: override for missing part ${o}`)
        if (overrides.filter((x) => x === o).length > 1)
            findings.push(`[Content_Types].xml: duplicate override ${o}`)
    }

    /*  5. every part must be covered by a default or override  */
    const defaults = new Set([...ct.matchAll(/<Default Extension="([^"]+)"/g)].map((m) => (m[1] as string).toLowerCase()))
    const overridden = new Set(overrides.map((o) => o.replace(/^\//, "")))
    for (const n of names)
        if (n !== "[Content_Types].xml" && !overridden.has(n)
            && !defaults.has(n.split(".").pop()?.toLowerCase() ?? ""))
            findings.push(`${n}: no content type (neither default nor override)`)

    /*  6. every sldId in the presentation must resolve to a slide part  */
    const pres = await text("ppt/presentation.xml")
    const rels = await text("ppt/_rels/presentation.xml.rels")
    const relMap = new Map([...rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)]
        .map((m) => [m[1] as string, m[2] as string]))
    for (const m of pres.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g)) {
        const target = relMap.get(m[1] as string)
        if (target === undefined || !names.has(path.posix.normalize(path.posix.join("ppt", target))))
            findings.push(`presentation.xml: sldId ${m[1]} does not resolve to a slide part`)
    }

    /*  7. CT_Presentation is an ordered sequence: master id lists must
        precede sldIdLst (violations trigger the repair dialog)  */
    const seq = [...pres.matchAll(/<p:(sldMasterIdLst|notesMasterIdLst|handoutMasterIdLst|sldIdLst|sldSz|notesSz)[ >/]/g)]
        .map((m) => m[1] as string)
    const sldAt = seq.indexOf("sldIdLst")
    if (sldAt >= 0)
        for (const lst of ["notesMasterIdLst", "handoutMasterIdLst"])
            if (seq.includes(lst) && seq.indexOf(lst) > sldAt)
                findings.push(`presentation.xml: ${lst} appears after sldIdLst (schema order violated)`)
    if (seq.includes("sldSz") && seq.includes("notesSz")
        && seq.indexOf("sldSz") > seq.indexOf("notesSz"))
        findings.push("presentation.xml: sldSz appears after notesSz (schema order violated)")

    /*  7a. no dangling /slide relationship: automizer mints a fresh
        "-created" slide rel on every re-import but wires only one into the
        sldIdLst; the rest pile up pointing at parts that still exist, so the
        dead-target check (#3) misses them, and a thousand-strong rels file
        makes PowerPoint demand a repair. Every /slide rel must be live.  */
    const liveSlideRids = new Set([...pres.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g)].map((m) => m[1]))
    for (const m of rels.matchAll(/<Relationship Id="([^"]+)"[^>]*Type="[^"]*\/slide"[^>]*\/>/g))
        if (!liveSlideRids.has(m[1] as string))
            findings.push(`presentation.xml.rels: dangling slide relationship ${m[1]} (no sldId references it)`)

    /*  7b. section slide references must resolve to existing sldIds  */
    const realIds = new Set([...pres.matchAll(/<p:sldId [^>]*\bid="(\d+)"/g)].map((m) => m[1] as string))
    for (const m of pres.matchAll(/<p14:sldId id="(\d+)"/g))
        if (!realIds.has(m[1] as string))
            findings.push(`presentation.xml: section references removed slide id ${m[1]}`)

    /*  7c. docProps/app.xml must declare the real slide count and a matching
        title list. automizer leaves it stale when slides are added on a later
        apply, and a deck claiming fewer slides than it has makes PowerPoint
        demand a repair on open.  */
    if (names.has("docProps/app.xml")) {
        const app = await text("docProps/app.xml")
        const slideCount = [...pres.matchAll(/<p:sldId /g)].length
        const declared = /<Slides>(\d+)<\/Slides>/.exec(app)
        if (declared !== null && Number(declared[1]) !== slideCount)
            findings.push(`docProps/app.xml: declares ${declared[1]} slides, deck has ${slideCount}`)
        const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(app)
        const tp = /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(app)
        if (hp !== null && tp !== null) {
            const sum = [...(hp[1] ?? "").matchAll(/<vt:i4>(\d+)<\/vt:i4>/g)].reduce((a, m) => a + Number(m[1]), 0)
            const entries = [...(tp[1] ?? "").matchAll(/<vt:lpstr>/g)].length
            if (sum !== entries)
                findings.push(`docProps/app.xml: HeadingPairs sum ${sum} != TitlesOfParts ${entries}`)
        }
    }

    /*  7d. a notesSlide must belong to exactly one slide. Two slides
        pointing at one notes part (it carries a single back-reference) is
        invalid and makes PowerPoint demand a repair -- automizer renumbers
        slide parts across applies, which can alias the basename-derived pptc
        notes part name.  */
    const notesRef = new Map<string, string[]>()
    for (const n of names)
        if (/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(n))
            for (const m of (await text(n)).matchAll(/Target="\.\.\/notesSlides\/([^"]+)"/g)) {
                const part = m[1] as string
                notesRef.set(part, [...(notesRef.get(part) ?? []), n])
            }
    for (const [part, owners] of notesRef)
        if (owners.length > 1)
            findings.push(`notesSlide ${part} is referenced by ${owners.length} slides (must be 1:1)`)

    /*  8. charts, embeddings and media must be reachable from some part
        (orphans accumulate over applies and bloat or break the file)  */
    const isAsset = (f: string): boolean =>
        /^ppt\/(charts|embeddings|media)\//.test(f) && !f.includes("/_rels/")
    const live = new Set<string>()
    for (const n of names)
        if (n.endsWith(".rels")) {
            const base = path.posix.dirname(path.posix.dirname(n))
            for (const m of (await text(n)).matchAll(/<Relationship [^>]*?\/>/g)) {
                if (m[0].includes("TargetMode=\"External\""))
                    continue
                const target = /Target="([^"]+)"/.exec(m[0])?.[1] ?? ""
                live.add(path.posix.normalize(path.posix.join(base, target)))
            }
        }
    for (const n of names)
        if (isAsset(n) && !live.has(n))
            findings.push(`${n}: orphan asset (not referenced from any part)`)

    return findings
}

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
