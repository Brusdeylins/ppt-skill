/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/engine: execute a finished MutationPlan by editing the
**  deck package in place. Repair triggers of the input outside the slides
**  are repaired first; slides the plan does not touch are never parsed for
**  writing, renamed or re-serialized; new and copied slides get fresh part
**  names and sldIds above everything in use. Every touched slide gets all
**  of its edits in one pass, then the package-wide steps run, the output is
**  checked (repair triggers, content guard, byte guard) and written
**  atomically. A failing session never touches the target file.
*/

import JSZip from "jszip"
import { readFileSync } from "node:fs"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError, toPptcError } from "../../infra/errors.js"
import type { MutationPlan, SlidePlanEntry } from "../../core/ops/registry.js"
import type { LintWarning } from "../../core/lint.js"
import { atomicWrite } from "../../infra/fs.js"
import { notesOwner, partDefects, type PartView } from "../defects.js"
import { byteGuard, digestBytes, verifyContent, type SlideExpect } from "../guard.js"
import { finishPackage } from "../post.js"
import { ASSET_PART, REL, RELS, layoutOrder, readRels, relIs, relMap, relsOwner, relsPartOf } from "../parts.js"
import type { DeckArchive } from "../reader.js"
import { verifyZip } from "../verify.js"
import { NS_P, NS_R, elements, firstElement } from "../xml.js"
import { Package } from "./package.js"
import { repairPart } from "./repair.js"
import { createSlide, dropLinksToRemoved, editSlide } from "./slides.js"
import { generateElements } from "./transplant.js"

/**  is a part outside every slide's own content (presentation, masters,
     layouts, themes, document properties, custom XML, ...)  */
const isPackagePart = (part: string): boolean =>
    !/^ppt\/(slides|notesSlides)\//.test(part) && !ASSET_PART.test(part)

/**  result facts of an executed session  */
export interface SessionResult {
    /**  number of output slides  */
    slideCount: number
    /**  plan-order positions of the entries that carried a ref  */
    refIndexes: Record<string, number>
    /**  repairs of PowerPoint repair triggers the input deck carried
         (W_REPAIRED)  */
    warnings: LintWarning[]
}

/**  one entry of the input slide list  */
interface InputSlide {
    /**  its p:sldId element  */
    el: Element
    /**  its id as the plan knows it  */
    id: number
    /**  its slide part  */
    part: string
}

/**  the input deck's slide list  */
interface Deck {
    /**  the presentation part  */
    pres: Document
    /**  its relationships  */
    presRels: Document
    /**  the input slides in order  */
    input: InputSlide[]
    /**  the notes master, null when the deck has none  */
    notesMaster: string | null
}

/**  one output slide and what the apply does to it  */
interface SlideState {
    /**  its plan entry  */
    entry: SlidePlanEntry
    /**  the output slide part  */
    part: string
    /**  the parts of its own content, the slide first  */
    closure: PartView[]
    /**  whether the apply writes it  */
    touched: boolean
    /**  whether its notes part belongs to another slide  */
    sharedNotes: boolean
    /**  parts of its content other than the slide carrying repair triggers  */
    repairs: string[]
    /**  the repair triggers its content carries  */
    defects: string[]
    /**  whether a repair removed a whole shape (named in W_REPAIRED), so its
         content may legitimately be gone  */
    shapeRemoved: boolean
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

/**  a part with its DOM and relationships, as engine/defects reads it  */
const viewOf = async (pkg: Package, part: string): Promise<PartView> => ({
    part,
    doc:  /\.xml$/.test(part) && pkg.has(part) ? await pkg.xml(part) : null,
    rels: pkg.targets(part, await pkg.rels(part))
})

/**  the parts that make up a slide's own content: the slide and everything
     reachable from it except layouts, masters, other slides and themes
     (media included -- shared media must not change either)  */
const closure = async (pkg: Package, slide: string): Promise<PartView[]> => {
    const views: PartView[] = []
    const queue = [ slide ]
    while (queue.length > 0) {
        const part = queue.shift() as string
        if (views.some((v) => v.part === part) || !pkg.has(part))
            continue
        const view = await viewOf(pkg, part)
        views.push(view)
        for (const r of view.rels)
            if (r.target !== null && !relIs(r, RELS.foreign))
                queue.push(r.target)
    }
    return views
}

/**  a W_REPAIRED warning  */
const repairedWarning = (slide: LintWarning["slide"], message: string): LintWarning =>
    ({ code: "W_REPAIRED", slide, message })

/**
 *  Check what an apply needs from the deck beyond its layouts, before
 *  anything runs (also for --dry-run).
 *
 *  @param plan - the finished mutation plan
 *  @param archive - the open input deck
 *  @throws PptcError E_TEMPLATE when speaker notes are planned on a deck
 *          without notes master
 */
export const preflight = async (plan: MutationPlan, archive: DeckArchive): Promise<void> => {
    if (!plan.entries.some((e) => e.notes !== null))
        return
    const master = (await readRels(archive.zip, "ppt/presentation.xml")).find((r) => r.kind === "notesMaster")?.target ?? null
    if (master === null || archive.zip.file(master) === null)
        throw new PptcError("E_TEMPLATE", "the deck has no notes master, which speaker notes need -- "
            + "add a note to any slide in PowerPoint once, then retry")
}

/**  repair triggers outside the slides (presentation, masters, layouts,
     themes, document properties, ...) are repaired before the engine
     writes anything, so a defect of the input never hides one of the
     engine's own  */
const repairPackage = async (pkg: Package, warnings: LintWarning[], repaired: Set<string>): Promise<void> => {
    const owners = new Set(Object.keys(pkg.zip.files).filter((f) => !f.endsWith("/"))
        .map((f) => f.endsWith(".rels") ? relsOwner(f) : f).filter((f) => f === "" || /\.xml$/.test(f)))
    for (const part of [ ...owners ].filter(isPackagePart)) {
        const defects = partDefects(await viewOf(pkg, part), (p) => pkg.has(p))
        const actions = defects.length > 0 ? await repairPart(pkg, part) : []
        if (actions.length > 0) {
            repaired.add(part)
            warnings.push(repairedWarning({ id: null, index: -1, title: null },
                `the input deck carried PowerPoint repair triggers in ${part === "" ? "_rels/.rels" : part} (${defects.join("; ")}); `
                + `pptc ${actions.join(", ")}`))
        }
    }
}

/**  the input slide list of the (repaired) package  */
const readDeck = async (pkg: Package): Promise<Deck> => {
    const pres = await pkg.xml("ppt/presentation.xml")
    const presRels = await pkg.rels("ppt/presentation.xml")
    const targets = pkg.targets("ppt/presentation.xml", presRels)
    const ridTarget = relMap(targets)
    const input = elements(pres, "p:sldId").map((el) => ({
        el, id: Number(el.getAttribute("id")), part: ridTarget.get(el.getAttribute("r:id") ?? "")?.target ?? ""
    }))
    const notesMaster = targets.find((t) => t.kind === "notesMaster")?.target ?? null
    return { pres, presRels, input, notesMaster }
}

/**  the input slide each plan entry keeps, null for new and copied
     slides -- matched by id and part, so a deck with duplicate slide ids
     or a slide listed twice keeps every entry  */
const matchInput = (plan: MutationPlan, deck: Deck): (InputSlide | null)[] => {
    const used = new Set<InputSlide>()
    return plan.entries.map((entry) => {
        if (entry.virtualId <= 0)
            return null
        const free = deck.input.filter((s) => !used.has(s) && s.id === entry.virtualId)
        const kept = free.find((s) => entry.source.kind !== "self" || s.part === entry.source.part) ?? free[0] ?? null
        if (kept !== null)
            used.add(kept)
        return kept
    })
}

/**  the output slide parts: kept slides stay where they are; copies are
     cloned from the unchanged input, new slides come from the deck's own
     layouts  */
const buildSlides = async (pkg: Package, plan: MutationPlan, matched: (InputSlide | null)[]): Promise<string[]> => {
    const layouts = await layoutOrder(pkg.zip)
    const parts: string[] = []
    for (const [ i, entry ] of plan.entries.entries()) {
        const kept = matched[i] ?? null
        if (kept !== null)
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
    return parts
}

/**  what the apply does to each output slide: a slide is touched when the
     plan edits, shows or hides it, when it links to a removed slide, or
     when its content carries repair triggers (engine/defects) or a notes
     part of another slide; untouched slides are never written  */
const classifySlides = async (pkg: Package, plan: MutationPlan, parts: string[], removed: Set<string>): Promise<SlideState[]> => {
    const exists = (p: string): boolean => pkg.has(p)
    const states: SlideState[] = []
    for (const [ i, entry ] of plan.entries.entries()) {
        const part = parts[i] as string
        const views = await closure(pkg, part)
        const slide = views[0] as PartView
        const defects: string[] = []
        const repairs: string[] = []
        for (const view of views) {
            const found = partDefects(view, exists)
            defects.push(...found)
            if (found.length > 0 && view.part !== part)
                repairs.push(view.part)
        }
        const notes = slide.rels.find((r) => r.kind === "notesSlide")?.target ?? null
        const notesView = views.find((v) => v.part === notes)
        const sharedNotes = notesView !== undefined && notesOwner(notesView) !== part
        if (sharedNotes)
            defects.push(`notes part ${notes ?? ""} belongs to another slide or none`)
        const show = slide.doc?.documentElement?.getAttribute("show") ?? ""
        const links = views.some((v) => v.rels.some((r) => r.kind === "slide" && r.target !== null && removed.has(r.target)))
        const touched = entry.virtualId <= 0 || hasEdits(entry) || (show === "0" || show === "false") !== entry.hidden
            || defects.length > 0 || links
        states.push({ entry, part, closure: views, touched, sharedNotes, repairs, defects, shapeRemoved: false })
    }
    return states
}

/**  the slide list: kept sldIds in their new order, fresh ids for new and
     copied slides, removed slides and their relationships gone  */
const writeSlideList = (pkg: Package, deck: Deck, states: SlideState[], matched: (InputSlide | null)[],
    removed: Set<string>, force: boolean): void => {
    const { pres, presRels, input } = deck
    if (!force && states.length === input.length && states.every((s, i) => matched[i] === input[i]))
        return
    let lst = firstElement(pres, "p:sldIdLst")
    if (lst === null) {
        lst = pres.createElementNS(NS_P, "p:sldIdLst")
        pres.documentElement?.insertBefore(lst, firstElement(pres, "p:sldSz"))
    }
    for (const s of input)
        s.el.parentNode?.removeChild(s.el)
    for (const [ i, { part } ] of states.entries()) {
        let el: Element | undefined = matched[i]?.el
        if (el === undefined) {
            el = pres.createElementNS(NS_P, "p:sldId")
            el.setAttribute("id", String(pkg.alloc.slideId()))
            el.setAttributeNS(NS_R, "r:id", pkg.addRel(presRels, `${REL}/slide`, path.posix.relative("ppt", part)))
        }
        lst.appendChild(el)
    }
    for (const t of pkg.targets("ppt/presentation.xml", presRels))
        if (t.kind === "slide" && t.target !== null && removed.has(t.target))
            t.el.parentNode?.removeChild(t.el)
    pkg.write("ppt/presentation.xml", pres)
    pkg.writeRels("ppt/presentation.xml", presRels)
}

/**  removed slides: their parts go, links to them from parts outside the
     slides go  */
const removeSlides = async (pkg: Package, removed: Set<string>): Promise<void> => {
    if (removed.size === 0)
        return
    await pkg.remove([ ...removed ])
    for (const owner of Object.keys(pkg.zip.files).filter((f) =>
        /^ppt\/(slideMasters|slideLayouts|diagrams|notesSlides)\/_rels\/[^/]+\.rels$|^ppt\/_rels\/viewProps\.xml\.rels$/.test(f)).map(relsOwner)) {
        if (!pkg.has(owner))
            continue
        const rels = await pkg.rels(owner)
        const doc = await pkg.xml(owner)
        if (dropLinksToRemoved(pkg, owner, doc, rels, removed)) {
            pkg.writeRels(owner, rels)
            pkg.write(owner, doc)
        }
    }
}

/**  what the content guard allows to change on each output slide  */
const expectations = (states: SlideState[]): SlideExpect[] =>
    states.map(({ entry, touched, shapeRemoved }, i) => ({
        source: entry.source.kind === "self" && touched ? entry.source.part : null,
        label: `#${i} '${entry.title ?? ""}'`,
        notes: entry.notes !== null,
        assets: shapeRemoved || entry.removeNames.length > 0 || entry.background !== null || entry.fills.some((f) => f.image !== undefined)
            ? "any" : entry.elements.length > 0 ? "grow" : "same",
        grow: {
            images: entry.elements.flatMap((e) => e.spec.type === "image" ? [ digestBytes(readFileSync(e.spec.path)) ] : []),
            fallbacks: entry.elements.filter((e) => e.spec.type === "image" && /\.svg$/i.test(e.spec.path)).length,
            charts: entry.elements.filter((e) => e.spec.type === "chart").length
        },
        textReplaced: entry.setTexts.length > 0 || entry.fills.some((f) => f.text !== undefined && !f.append)
    }))

/**  the passes of an apply: package repairs, slide structure, edits,
     slide list, package-wide steps, self-checks, write  */
const executePlan = async (outFile: string, plan: MutationPlan, archive: DeckArchive): Promise<SessionResult> => {
    await preflight(plan, archive)

    /*  the package as on disk: untouched parts keep their exact bytes  */
    const inputIds = elements(await archive.xml("ppt/presentation.xml"), "p:sldId").map((s) => Number(s.getAttribute("id")))
    const pkg = await Package.open(archive.bytes, inputIds)
    const warnings: LintWarning[] = []
    const repaired = new Set<string>()
    await repairPackage(pkg, warnings, repaired)
    const deck = await readDeck(pkg)

    /*  duplicate slide ids get fresh ones (the plan matches by part)  */
    const ids = new Set<number>()
    let renumbered = false
    for (const s of deck.input) {
        if (ids.has(s.id)) {
            s.el.setAttribute("id", String(pkg.alloc.slideId()))
            renumbered = true
        }
        ids.add(s.id)
    }
    if (renumbered)
        warnings.push(repairedWarning({ id: null, index: -1, title: null },
            "the input deck carried PowerPoint repair triggers in ppt/presentation.xml (duplicate slide ids); pptc gave the duplicates fresh ids"))

    /*  removed slides (a part another, kept entry lists stays), plus slide
        parts outside the slide list (garbage PowerPoint complains about)  */
    const matched = matchInput(plan, deck)
    const keptParts = new Set(matched.flatMap((s) => s === null ? [] : [ s.part ]))
    const removed = new Set(deck.input.filter((s) => !keptParts.has(s.part)).map((s) => s.part))
    for (const orphan of Object.keys(pkg.zip.files).filter((f) =>
        /^ppt\/slides\/[^/]+\.xml$/.test(f) && !deck.input.some((s) => s.part === f))) {
        removed.add(orphan)
        warnings.push(repairedWarning({ id: null, index: -1, title: null },
            `the input deck carried a slide part outside its slide list; pptc removed ${orphan}`))
    }

    const parts = await buildSlides(pkg, plan, matched)
    const states = await classifySlides(pkg, plan, parts, removed)
    const untouched = new Map<string, JSZip.JSZipObject | null>()
    for (const s of states.filter((s) => !s.touched))
        for (const { part } of s.closure)
            for (const p of pkg.has(relsPartOf(part)) ? [ part, relsPartOf(part) ] : [ part ])
                untouched.set(p, pkg.zip.file(p))

    /*  one pass per touched slide, after the repairs of its other parts  */
    const generated = await generateElements(plan.entries)
    for (const [ i, s ] of states.entries()) {
        if (!s.touched)
            continue
        const actions: string[] = []
        for (const owner of s.repairs)
            actions.push(...(await repairPart(pkg, owner)).map((a) => `${a} in ${owner}`))
        actions.push(...await editSlide(pkg, {
            part: s.part, entry: s.entry, removed, notesMaster: deck.notesMaster ?? "", generated, sharedNotes: s.sharedNotes
        }))
        if (s.sharedNotes)
            actions.push("gave the slide its own notes part")
        s.shapeRemoved = actions.some((a) => a.startsWith("removed the shape"))
        if (s.defects.length > 0)
            warnings.push(repairedWarning({ id: s.entry.virtualId > 0 ? s.entry.virtualId : null, index: i, title: s.entry.title },
                `the input deck carried PowerPoint repair triggers on this slide${s.entry.virtualId > 0 ? "" : " (or the slide it was copied from)"} `
                + `(${s.defects.join("; ")}); pptc ${actions.length > 0 ? actions.join(", ") : "repaired them"}`))
    }
    writeSlideList(pkg, deck, states, matched, removed, renumbered)
    await removeSlides(pkg, removed)
    await finishPackage(pkg.zip, parts, plan.props, plan.customProps)
    const finalBytes = await pkg.zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })

    /*  self-checks before writing: repair triggers, the content of every
        touched slide (canonical), every part of every untouched slide  */
    const after = await JSZip.loadAsync(finalBytes)
    const expects = expectations(states)
    const content = await verifyContent(archive.zip, after, expects, repaired)

    /*  repair triggers the input already carried and pptc cannot repair
        stay as they were (W_INPUT_DEFECT); only new ones block the write  */
    let triggers = await verifyZip(after, archive.zip)
    if (triggers.length > 0) {
        const inherited = new Set(await verifyZip(archive.zip))
        const kept = triggers.filter((f) => inherited.has(f))
        triggers = triggers.filter((f) => !inherited.has(f))
        if (kept.length > 0)
            warnings.push({ code: "W_INPUT_DEFECT", slide: { id: null, index: -1, title: null },
                message: `the input deck carries PowerPoint repair triggers pptc cannot repair; they stay as they were: ${kept.join("; ")}` })
    }
    const findings = [
        ...triggers,
        ...content.blocked,
        ...content.lost.map((l) => `slide ${(expects[l.slide] as SlideExpect).label}: ${l.what} would be lost although no op removed it`),
        ...await byteGuard(pkg.zip, untouched, after)
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
 *  Execute a mutation plan against an open deck.
 *
 *  @param outFile - the file to write atomically (may be the deck itself)
 *  @param plan - the finished mutation plan
 *  @param archive - the open input deck the plan was made for
 *  @returns session facts for the result envelope
 *  @throws PptcError E_INTEGRITY when a self-check fails, E_TEMPLATE when
 *          the deck lacks what the plan needs, E_ENGINE on unexpected
 *          failures (target untouched in every case)
 */
export const runSession = async (outFile: string, plan: MutationPlan, archive: DeckArchive): Promise<SessionResult> => {
    try {
        return await withStdoutShield(() => executePlan(outFile, plan, archive))
    }
    catch (err) {
        throw toPptcError(err)
    }
}
