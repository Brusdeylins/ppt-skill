/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/engine: execute a finished MutationPlan by editing the
**  deck package in place. Slides the plan does not touch are never parsed
**  for writing, renamed or re-serialized; new and copied slides get fresh
**  part names and sldIds above everything in use. Every touched slide gets
**  all of its edits in one pass, then the package-wide steps run, the output
**  is checked (repair triggers, content guard, byte guard) and written
**  atomically. A failing session never touches the target file.
*/

import JSZip from "jszip"
import { readFileSync } from "node:fs"
import path from "node:path"
import type { Element } from "@xmldom/xmldom"
import { PptcError, toPptcError } from "../../infra/errors.js"
import type { MutationPlan, SlidePlanEntry } from "../../core/ops/registry.js"
import type { LintWarning } from "../../core/lint.js"
import { atomicWrite } from "../../infra/fs.js"
import { digestBytes, verifyContent, type SlideExpect } from "../guard.js"
import { finishPackage } from "../post.js"
import { normalizeRelTargets, relsPartOf } from "../parts.js"
import type { DeckArchive } from "../reader.js"
import { layoutOrder } from "../seed.js"
import { refFindings, relsFindings, shapeIdFindings, verifyBytes } from "../verify.js"
import { NS_P, NS_R, doubleQuoted, elements, firstElement } from "../xml.js"
import { Package } from "./package.js"
import { createSlide, dropLinksToRemoved, editSlide, repairPackageRels, repairPart } from "./slides.js"
import { generateElements } from "./transplant.js"

/**  OfficeDocument relationship type base URI  */
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/**  relationship types that leave a slide's own content: its layout and
     master, other slides, the notes master and themes  */
const FOREIGN_RELS = [ "/slideLayout", "/slideMaster", "/slide", "/notesMaster", "/theme" ]

/**  result facts of an executed session  */
export interface SessionResult {
    /**  number of output slides  */
    slideCount: number
    /**  plan-order positions of the entries that carried a ref  */
    refIndexes: Record<string, number>
    /**  slides repaired because the input deck carried repair triggers
         on them (W_REPAIRED)  */
    warnings: LintWarning[]
}

/**
 *  Run a function with console output diverted to stderr. stdout belongs
 *  exclusively to the envelope, so anything a library prints (PptxGenJS
 *  warns on odd images) must land on stderr.
 *
 *  @param fn - the function to run shielded
 *  @returns the function's result
 */
const withStdoutShield = async <T>(fn: () => Promise<T>): Promise<T> => {
    const original = { log: console.log, info: console.info, warn: console.warn }
    console.log = (...args: unknown[]): void => { console.error(...args) }
    console.info = (...args: unknown[]): void => { console.error(...args) }
    console.warn = (...args: unknown[]): void => { console.error(...args) }
    try {
        return await fn()
    }
    finally {
        console.log = original.log
        console.info = original.info
        console.warn = original.warn
    }
}

/**  does the plan change the content of the slide (not only its position)  */
const hasEdits = (e: SlidePlanEntry): boolean =>
    e.fills.length > 0 || e.setTexts.length > 0 || e.removeNames.length > 0 || e.elements.length > 0
        || e.notes !== null || e.footer !== null || e.background !== null

/**  the parts that make up a slide's own content: the slide, its rels and
     everything reachable from it except layouts, masters, other slides and
     themes (media included -- shared media must not change either)  */
const closure = async (pkg: Package, slide: string): Promise<string[]> => {
    const parts: string[] = []
    const queue = [ slide ]
    while (queue.length > 0) {
        const part = queue.pop() as string
        if (parts.includes(part) || !pkg.has(part))
            continue
        parts.push(part)
        if (!pkg.has(relsPartOf(part)))
            continue
        parts.push(relsPartOf(part))
        for (const t of Package.targets(part, await pkg.rels(part)))
            if (t.target !== null && !FOREIGN_RELS.some((k) => t.type.endsWith(k)))
                queue.push(t.target)
    }
    return parts
}

/**  do the relationships of any part of a closure link to a removed slide  */
const linksTo = async (pkg: Package, parts: string[], removed: Set<string>): Promise<boolean> => {
    for (const part of parts.filter((p) => !p.endsWith(".rels") && pkg.has(relsPartOf(p))))
        if (Package.targets(part, await pkg.rels(part)).some((t) =>
            t.type.endsWith("/slide") && t.target !== null && removed.has(t.target)))
            return true
    return false
}

/**  the two passes: slide structure and edits, then package-wide steps  */
const executePlan = async (outFile: string, plan: MutationPlan, archive: DeckArchive): Promise<SessionResult> => {
    /*  the package as on disk: untouched parts keep their exact bytes  */
    const presIn = await archive.xml("ppt/presentation.xml")
    const inputIds = elements(presIn, "p:sldId").map((s) => Number(s.getAttribute("id")))
    const pkg = await Package.open(readFileSync(archive.file), inputIds)
    const pres = await pkg.xml("ppt/presentation.xml")
    const presRels = await pkg.rels("ppt/presentation.xml")
    const ridTarget = new Map(Package.targets("ppt/presentation.xml", presRels)
        .map((t) => [ t.rel.getAttribute("Id") ?? "", t.target ?? "" ]))
    const input = elements(pres, "p:sldId").map((el) => ({
        el, id: Number(el.getAttribute("id")), part: ridTarget.get(el.getAttribute("r:id") ?? "") ?? ""
    }))
    const notesMaster = Package.targets("ppt/presentation.xml", presRels)
        .find((t) => t.type.endsWith("/notesMaster"))?.target ?? null
    if (plan.entries.some((e) => e.notes !== null) && (notesMaster === null || !pkg.has(notesMaster)))
        throw new PptcError("E_TEMPLATE", "the deck has no notes master, which speaker notes need -- "
            + "add a note to any slide in PowerPoint once, then retry")
    const layouts = await layoutOrder(archive.zip)

    /*  structure: kept slides stay where they are; copies are cloned from
        the unchanged input, new slides come from the deck's own layouts  */
    const keptIds = new Set(plan.entries.filter((e) => e.virtualId > 0).map((e) => e.virtualId))
    const removed = new Set(input.filter((s) => !keptIds.has(s.id)).map((s) => s.part))
    const parts: string[] = []
    for (const entry of plan.entries) {
        const kept = input.find((s) => s.id === entry.virtualId)
        if (kept !== undefined)
            parts.push(kept.part)
        else if (entry.source.kind === "self")
            parts.push(await pkg.clonePart(pkg, entry.source.part))
        else {
            const layout = layouts[entry.source.layoutIndex]
            if (layout === undefined)
                throw new PptcError("E_TEMPLATE", `the deck has no layout ${entry.source.layoutIndex}`)
            parts.push(await createSlide(pkg, layout))
        }
    }

    /*  which kept slides are touched: edited, re-shown/hidden, linking to a
        removed slide, sharing a notes part, or carrying repair triggers  */
    const names = new Set(Object.keys(archive.zip.files).filter((n) => !n.endsWith("/") && !n.startsWith("[trash]/")))
    const quoted = async (part: string): Promise<string> => doubleQuoted(await archive.text(part) ?? "")
    const touched: boolean[] = []
    const sharedNotes: boolean[] = []
    const repairs: string[][] = []
    const untouched = new Map<string, JSZip.JSZipObject | null>()
    const warnings: LintWarning[] = []
    for (const [ i, entry ] of plan.entries.entries()) {
        const part = parts[i] as string
        touched.push(true)
        sharedNotes.push(false)
        repairs.push([])
        if (entry.virtualId <= 0)
            continue

        /*  repair triggers of the slide and of every part of its own
            content (notes, charts, ...): such a slide is repaired  */
        const own = await closure(pkg, part)
        const xml = await quoted(part)
        const defects = [ ...shapeIdFindings(part, xml), ...refFindings(part, xml, await quoted(relsPartOf(part))) ]
        for (const relsPart of own.filter((p) => p.endsWith(".rels"))) {
            const found = relsFindings(relsPart, await quoted(relsPart), names)
            defects.push(...found)
            const owner = relsPart.replace(/_rels\/([^/]+)\.rels$/, "$1")
            if (found.length > 0 && owner !== part)
                (repairs[i] as string[]).push(owner)
        }

        /*  a notes part whose back-reference names another slide (shared,
            or left over by a removed slide) gets a copy of its own  */
        const notes = Package.targets(part, await pkg.rels(part)).find((t) => t.type.endsWith("/notesSlide"))?.target ?? null
        if (notes !== null && pkg.has(notes)) {
            const notesXml = await pkg.text(notes)
            const backref = Package.targets(notes, await pkg.rels(notes)).find((t) =>
                t.type.endsWith("/slide") && !notesXml.includes(`"${t.rel.getAttribute("Id") ?? ""}"`))?.target ?? part
            if (backref !== part) {
                sharedNotes[i] = true
                defects.push(`notes part ${notes} belongs to another slide`)
            }
        }
        const hidden = /<p:sld\b[^>]*\bshow="(0|false)"/.test(xml)
        const touch = hasEdits(entry) || hidden !== entry.hidden || defects.length > 0 || await linksTo(pkg, own, removed)
        touched[i] = touch
        if (defects.length > 0)
            warnings.push({
                code: "W_REPAIRED",
                slide: { id: entry.virtualId, index: i, title: entry.title },
                message: `the input deck carried PowerPoint repair triggers on this slide; pptc repaired them: ${defects.join("; ")}`
            })
        if (!touch)
            for (const p of own)
                untouched.set(p, pkg.zip.file(p))
    }

    /*  one pass per touched slide, then the repairs of its other parts  */
    const generated = await generateElements(plan.entries)
    for (const [ i, entry ] of plan.entries.entries())
        if (touched[i] === true) {
            await editSlide(pkg, {
                part: parts[i] as string, entry, removed, notesMaster: notesMaster ?? "", generated,
                sharedNotes: sharedNotes[i] === true
            })
            for (const owner of repairs[i] as string[])
                await repairPart(pkg, owner)
        }

    /*  the slide list: kept sldIds in their new order, fresh ids for new
        and copied slides, removed slides and their relationships gone  */
    if (parts.length !== input.length || parts.some((p, i) => p !== input[i]?.part)) {
        let lst = firstElement(pres, "p:sldIdLst")
        if (lst === null) {
            lst = pres.createElementNS(NS_P, "p:sldIdLst")
            pres.documentElement?.insertBefore(lst, firstElement(pres, "p:sldSz"))
        }
        for (const s of input)
            lst.removeChild(s.el)
        for (const [ i, entry ] of plan.entries.entries()) {
            let el: Element | undefined = input.find((s) => s.id === entry.virtualId)?.el
            if (el === undefined) {
                el = pres.createElementNS(NS_P, "p:sldId")
                el.setAttribute("id", String(pkg.alloc.slideId()))
                el.setAttributeNS(NS_R, "r:id", pkg.addRel(presRels, `${REL}/slide`,
                    path.posix.relative("ppt", parts[i] as string)))
            }
            lst.appendChild(el)
        }
        for (const s of input.filter((s) => removed.has(s.part)))
            for (const rel of elements(presRels, "Relationship"))
                if (rel.getAttribute("Id") === s.el.getAttribute("r:id"))
                    rel.parentNode?.removeChild(rel)
        pkg.write("ppt/presentation.xml", pres)
        pkg.writeRels("ppt/presentation.xml", presRels)
    }

    /*  removed slides: their parts go, links to them elsewhere go  */
    await pkg.remove([ ...removed ])
    if (removed.size > 0)
        for (const relsPart of Object.keys(pkg.zip.files).filter((f) =>
            /^ppt\/(slideMasters|slideLayouts|diagrams|notesSlides)\/_rels\/[^/]+\.rels$|^ppt\/_rels\/viewProps\.xml\.rels$/.test(f))) {
            const owner = relsPart.replace(/_rels\/([^/]+)\.rels$/, "$1")
            if (!pkg.has(owner))
                continue
            const rels = await pkg.rels(owner)
            const doc = await pkg.xml(owner)
            if (dropLinksToRemoved(owner, doc, rels, removed)) {
                pkg.writeRels(owner, rels)
                pkg.write(owner, doc)
            }
        }

    /*  repair triggers outside the slides (layouts, masters, presentation):
        dead and duplicate relationships go  */
    for (const relsPart of Object.keys(pkg.zip.files).filter((f) =>
        /^ppt\/(slideLayouts|slideMasters|notesMasters|handoutMasters|theme)\/_rels\/[^/]+\.rels$|^ppt\/_rels\/presentation\.xml\.rels$/.test(f))) {
        const owner = relsPart.replace(/_rels\/([^/]+)\.rels$/, "$1")
        if (!untouched.has(relsPart) && relsFindings(relsPart, doubleQuoted(await pkg.text(relsPart)), new Set(Object.keys(pkg.zip.files))).length > 0
            && await repairPackageRels(pkg, owner))
            warnings.push({
                code: "W_REPAIRED",
                slide: { id: null, index: -1, title: null },
                message: `the input deck carried PowerPoint repair triggers in ${owner}; pptc dropped its dead or duplicate relationships`
            })
    }
    await finishPackage(pkg.zip, parts, plan.props, plan.customProps)
    const finalBytes = await pkg.zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })

    /*  self-checks before writing: repair triggers, the content of every
        touched slide (canonical), and every part of every untouched slide
        (never written at all)  */
    const after = await JSZip.loadAsync(finalBytes)
    const changed: string[] = []
    for (const [ name, file ] of untouched) {
        const now = after.file(name)
        if (pkg.zip.file(name) !== file || file === null || now === null
            || !Buffer.from(await now.async("uint8array")).equals(Buffer.from(await file.async("uint8array"))))
            changed.push(`${name}: part of an untouched slide would change`)
    }
    await normalizeRelTargets(after)
    const expects: SlideExpect[] = plan.entries.map((entry, i) => ({
        source: entry.source.kind === "self" && touched[i] === true ? entry.source.part : null,
        label: `#${i} '${entry.title ?? ""}'`,
        notes: entry.notes !== null,
        assets: entry.removeNames.length > 0 || entry.background !== null || entry.fills.some((f) => f.image !== undefined)
            ? "any" : entry.elements.length > 0 ? "grow" : "same",
        grow: {
            images: entry.elements.flatMap((e) => e.spec.type === "image" ? [ digestBytes(readFileSync(e.spec.path)) ] : []),
            fallbacks: entry.elements.filter((e) => e.spec.type === "image" && /\.svg$/i.test(e.spec.path)).length,
            charts: entry.elements.filter((e) => e.spec.type === "chart").length
        },
        textReplaced: entry.setTexts.length > 0 || entry.fills.some((f) => f.text !== undefined && !f.append)
    }))
    const content = await verifyContent(archive.zip, after, expects)
    const findings = [
        ...await verifyBytes(finalBytes, archive.zip),
        ...content.blocked,
        ...content.lost.map((l) => `slide ${(expects[l.slide] as SlideExpect).label}: ${l.what} would be lost although no op removed it`),
        ...changed
    ]
    if (findings.length > 0)
        throw new PptcError("E_INTEGRITY",
            `the engine produced ${findings.length} integrity violation(s); deck not written`,
            { findings })
    atomicWrite(path.resolve(outFile), finalBytes)

    const refIndexes: Record<string, number> = {}
    for (const [ name, entry ] of plan.refs)
        refIndexes[name] = plan.entries.indexOf(entry)
    return { slideCount: plan.entries.length, refIndexes, warnings }
}

/**
 *  Execute a mutation plan against a deck file.
 *
 *  @param deckFile - the deck to read from
 *  @param outFile - the file to write atomically (may equal deckFile)
 *  @param plan - the finished mutation plan
 *  @param archive - the caller's open archive of deckFile (read-only reuse)
 *  @returns session facts for the result envelope
 *  @throws PptcError E_INTEGRITY when a self-check fails, E_ENGINE on
 *          unexpected failures (target untouched in both cases)
 */
export const runSession = async (
    deckFile: string,
    outFile: string,
    plan: MutationPlan,
    archive: DeckArchive
): Promise<SessionResult> => {
    try {
        if (path.resolve(deckFile) !== archive.file)
            throw new PptcError("E_ENGINE", "the archive does not belong to the deck file")
        return await withStdoutShield(() => executePlan(outFile, plan, archive))
    }
    catch (err) {
        throw toPptcError(err)
    }
}
