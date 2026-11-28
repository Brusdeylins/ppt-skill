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
import { notesOwner, partDefects, type Finding, type PartView } from "../defects.js"
import { byteGuard, digestBytes, verifyContent, type SlideExpect } from "../guard.js"
import { deckLang } from "../lang.js"
import { finishPackage, type SectionOut } from "../post.js"
import { ASSET_PART, REL, RELS, findPart, layoutOrder, relTarget, readRels, relIs, relsOwner, relsPartOf, slideList } from "../parts.js"
import type { DeckArchive } from "../reader.js"
import { findDefects } from "../verify.js"
import { NS_P, NS_R, elements, firstElement } from "../xml.js"
import { Package } from "./package.js"
import { repairPart } from "./repair.js"
import { createSlide, dropLinksToRemoved, editSlide } from "./slides.js"
import { generateElements } from "./transplant.js"
import { mayMatch } from "../../core/text-replace.js"

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
         (W_REPAIRED), and triggers left as they were (W_INPUT_DEFECT)  */
    warnings: LintWarning[]
    /**  replacements of each text.replace op in the content the output
         slides carried before the apply, by plan-order slide position  */
    replaced: { index: number, op: number, count: number }[]
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

/**  the slide address of a warning about the deck as a whole  */
const DECK_WIDE: LintWarning["slide"] = { id: null, index: -1, title: null }

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
    /**  the same parts in the input (for a copy: those of its source)  */
    sourceRepairs: string[]
    /**  the repair triggers its content carries  */
    defects: Finding[]
    /**  relationship ids whose content a named repair removed (W_REPAIRED)  */
    released: string[]
    /**  the same, released inside the repaired parts of its content  */
    partReleased: string[]
    /**  images the slide pass put in place of others or adopted from the
         layout (content guard)  */
    replacedImages: { from: string[], to: string }[]
    /**  replacements of each text.replace op in its existing content  */
    replaced: { op: number, count: number }[]
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
    e.fills.length > 0 || e.edits.length > 0 || e.removeNames.length > 0 || e.elements.length > 0
        || e.notes !== null || e.footer !== null || e.background !== null || e.slideNumber !== null
        || e.date !== null

/**  whether a text.replace of the plan may find something on a slide (in
     its notes too, when it replaces there): a slide whose text holds no
     possible match stays untouched  */
const replacesMatch = (e: SlidePlanEntry, slide: Document | null, notes: Document | null): boolean => {
    const text = (doc: Document | null): string =>
        doc === null ? "" : elements(doc, "a:t").map((t) => t.textContent ?? "").join("")
    const slideText = e.replaces.length > 0 ? text(slide) : ""
    const notesText = e.replaces.some((r) => r.notes) ? text(notes) : ""
    return e.replaces.some((r) => mayMatch(slideText, r) || (r.notes && mayMatch(notesText, r)))
}

/**  a part with its DOM and relationships, as engine/defects reads it; a
     malformed part has no DOM (verify reports it)  */
const viewOf = async (pkg: Package, part: string): Promise<PartView> => {
    let doc: Document | null = null
    if (/\.xml$/.test(part) && pkg.has(part))
        try {
            doc = await pkg.xml(part)
        }
        catch {
            doc = null
        }
    let rels: PartView["rels"] = []
    try {
        rels = pkg.targets(part, await pkg.rels(part))
    }
    catch {
        doc = null
    }
    return { part, doc, rels }
}

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
    /*  a rels part whose owner is gone is garbage  */
    for (const relsPart of Object.keys(pkg.zip.files).filter((f) => f.endsWith(".rels") && relsOwner(f) !== "" && !pkg.has(relsOwner(f)))) {
        pkg.zip.remove(relsPart)
        warnings.push(repairedWarning(DECK_WIDE, `the input deck carried relationships of a missing part; pptc removed ${relsPart}`))
    }
    const owners = new Set(Object.keys(pkg.zip.files).filter((f) => !f.endsWith("/"))
        .map((f) => f.endsWith(".rels") ? relsOwner(f) : f).filter((f) => f === "" || /\.xml$/.test(f)))
    for (const part of [ ...owners ].filter(isPackagePart)) {
        const view = await viewOf(pkg, part)
        if (view.doc === null && /\.xml$/.test(part))
            continue
        const defects = partDefects(view, (p) => pkg.has(p))
        const { actions } = defects.length > 0 ? await repairPart(pkg, part) : { actions: [] }
        if (actions.length > 0) {
            repaired.add(part)
            warnings.push(repairedWarning(DECK_WIDE,
                `the input deck carried PowerPoint repair triggers in ${part === "" ? "_rels/.rels" : part} `
                + `(${defects.map((d) => d.text).join("; ")}); pptc ${actions.join(", ")}`))
        }
    }
}

/**  the input slide list of the (repaired) package; a sldId resolving to
     no slide part is dropped (the reader shows no slide for it)  */
const readDeck = async (pkg: Package, warnings: LintWarning[]): Promise<Deck & { dropped: boolean }> => {
    const pres = await pkg.xml("ppt/presentation.xml")
    const presRels = await pkg.rels("ppt/presentation.xml")
    const targets = pkg.targets("ppt/presentation.xml", presRels)
    const input: InputSlide[] = []
    let dropped = false
    for (const s of slideList(pkg.zip, pres, targets))
        if (s.part !== null)
            input.push({ el: s.el, id: s.id, part: s.part })
        else {
            s.el.parentNode?.removeChild(s.el)
            for (const t of targets.filter((t) => t.kind === "slide" && t.id === s.el.getAttribute("r:id")))
                t.el.parentNode?.removeChild(t.el)
            dropped = true
            warnings.push(repairedWarning(DECK_WIDE,
                `the input deck listed slide id ${s.id} without a slide part; pptc removed it from the slide list`))
        }
    const notesMaster = targets.find((t) => t.kind === "notesMaster")?.target ?? null
    return { pres, presRels, input, notesMaster, dropped }
}

/**  duplicate slide ids get fresh ones (the plan matches by part)
     @returns whether an id changed  */
const renumberSlideIds = (pkg: Package, deck: Deck, warnings: LintWarning[]): boolean => {
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
        warnings.push(repairedWarning(DECK_WIDE,
            "the input deck carried PowerPoint repair triggers in ppt/presentation.xml (duplicate slide ids); pptc gave the duplicates fresh ids"))
    return renumbered
}

/**  the slide parts that go: those of removed entries (a part another,
     kept entry lists stays), plus slide parts outside the slide list
     (garbage PowerPoint complains about)  */
const removedSlides = (pkg: Package, deck: Deck, matched: (InputSlide | null)[], warnings: LintWarning[]): Set<string> => {
    const keptParts = new Set(matched.flatMap((s) => s === null ? [] : [ s.part ]))
    const removed = new Set(deck.input.filter((s) => !keptParts.has(s.part)).map((s) => s.part))
    for (const orphan of Object.keys(pkg.zip.files).filter((f) =>
        /^ppt\/slides\/[^/]+\.xml$/.test(f) && !deck.input.some((s) => s.part === f))) {
        removed.add(orphan)
        warnings.push(repairedWarning(DECK_WIDE, `the input deck carried a slide part outside its slide list; pptc removed ${orphan}`))
    }
    return removed
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
        const defects: Finding[] = []
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
            defects.push({ rule: "notes-owner", part, text: `${part}: its notes part ${notes ?? ""} belongs to another slide or none` })
        const show = slide.doc?.documentElement?.getAttribute("show") ?? ""
        const links = views.some((v) => v.rels.some((r) => r.kind === "slide" && r.target !== null && removed.has(r.target)))
        const touched = entry.virtualId <= 0 || hasEdits(entry) || (show === "0" || show === "false") !== entry.hidden
            || defects.length > 0 || links || replacesMatch(entry, slide.doc, notesView?.doc ?? null)
        let sourceRepairs = repairs
        if (entry.virtualId <= 0 && entry.source.kind === "self" && repairs.length > 0)
            sourceRepairs = (await closure(pkg, entry.source.part)).slice(1)
                .filter((v) => partDefects(v, exists).length > 0).map((v) => v.part)
        states.push({ entry, part, closure: views, touched, sharedNotes, repairs, sourceRepairs, defects, released: [], partReleased: [],
            replacedImages: [], replaced: [] })
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
            el.setAttributeNS(NS_R, "r:id", pkg.addRel(presRels, `${REL}/slide`, relTarget("ppt/presentation.xml", part)))
        }
        lst.appendChild(el)
    }
    for (const t of pkg.targets("ppt/presentation.xml", presRels))
        if (t.kind === "slide" && t.target !== null && removed.has(t.target))
            t.el.parentNode?.removeChild(t.el)
    pkg.write("ppt/presentation.xml", pres)
    pkg.writeRels("ppt/presentation.xml", presRels)
}

/**  the planned sections with the output sldIds of their slides (the
     slide list is written: its order is the plan's), null when the plan
     leaves the sections to the package step  */
const sectionsOut = (plan: MutationPlan, deck: Deck): SectionOut[] | null => {
    if (plan.sections === null)
        return null
    const ids = elements(deck.pres, "p:sldId").map((s) => Number(s.getAttribute("id")))
    return plan.sections.map((section) => ({
        name: section.name,
        id: section.id,
        slideIds: plan.entries.flatMap((e, i) => e.section === section ? [ ids[i] as number ] : [])
    }))
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
    states.map(({ entry, touched, released, repairs, sourceRepairs, partReleased, replacedImages }, i) => ({
        source: entry.source.kind === "self" && touched ? entry.source.part : null,
        label: `#${i} '${entry.title ?? ""}'`,
        notes: entry.notes !== null || entry.replaces.some((r) => r.notes),
        assets: entry.removeNames.length > 0 || entry.background !== null || entry.fills.some((f) => f.image !== undefined)
            ? "any" : entry.elements.length > 0 ? "grow" : "same",
        grow: {
            images: entry.elements.flatMap((e) => e.spec.type === "image" ? [ digestBytes(readFileSync(e.spec.path)) ] : []),
            fallbacks: entry.elements.filter((e) => e.spec.type === "image" && /\.svg$/i.test(e.spec.path)).length,
            charts: entry.elements.filter((e) => e.spec.type === "chart").length
        },
        textReplaced: entry.edits.some((e) => e.set.text !== undefined || e.set.cells !== undefined
            || e.set.addRows !== undefined || e.set.removeRows !== undefined)
            || entry.fills.some((f) => f.text !== undefined && !f.append) || entry.replaces.length > 0,
        replaced: replacedImages,
        released,
        repaired: { before: sourceRepairs, after: repairs, released: partReleased }
    }))

/**  the parts of untouched slides with their input entry, for the byte
     guard -- except parts a touched slide writes (a slide listed twice, a
     part shared with a repaired slide)  */
const untouchedParts = (pkg: Package, states: SlideState[]): Map<string, JSZip.JSZipObject | null> => {
    const written = new Set(states.filter((s) => s.touched).flatMap((s) => [ s.part, ...s.repairs ]))
    const untouched = new Map<string, JSZip.JSZipObject | null>()
    for (const s of states.filter((s) => !s.touched))
        for (const { part } of s.closure)
            for (const p of [ part, findPart(pkg.zip, relsPartOf(part)) ].filter((n): n is string => n !== null))
                if (!written.has(p) && !written.has(relsOwner(p)))
                    untouched.set(p, pkg.zip.file(p))
    return untouched
}

/**  one pass per touched slide, after the repairs of its other parts;
     every repair is reported (W_REPAIRED) with what it removed  */
const editSlides = async (pkg: Package, plan: MutationPlan, deck: Deck, states: SlideState[], removed: Set<string>,
    warnings: LintWarning[]): Promise<void> => {
    const lang = await deckLang(pkg.zip, plan.customProps)
    const generated = await generateElements(plan.entries, lang)
    const links = new Map([ ...plan.links ].map(([ link, entry ]) => [ link, (states[plan.entries.indexOf(entry)] as SlideState).part ]))
    for (const [ i, s ] of states.entries()) {
        if (!s.touched)
            continue
        const actions: string[] = []
        for (const owner of s.repairs) {
            const repair = await repairPart(pkg, owner)
            actions.push(...repair.actions.map((a) => `${a} in ${owner}`))
            s.partReleased.push(...repair.released)
        }
        const own = await editSlide(pkg, {
            part: s.part, entry: s.entry, removed, notesMaster: deck.notesMaster ?? "", generated, sharedNotes: s.sharedNotes, lang, links
        })
        actions.push(...own.actions)
        s.replacedImages = [ ...own.swaps, ...own.adopted.map((to) => ({ from: [], to })) ]
        s.replaced = own.replaced

        /*  a slide listed twice is edited twice: the first pass released  */
        s.released = [ ...new Set([ ...states.filter((o) => o.part === s.part).flatMap((o) => o.released), ...own.released ]) ]
        if (s.sharedNotes)
            actions.push("gave the slide its own notes part")
        if (s.defects.length > 0 && (actions.length > 0 || !states.slice(0, i).some((o) => o.touched && o.defects.length > 0
            && (o.part === s.part || o.repairs.some((r) => s.repairs.includes(r))))))
            warnings.push(repairedWarning({ id: s.entry.virtualId > 0 ? s.entry.virtualId : null, index: i, title: s.entry.title },
                `the input deck carried PowerPoint repair triggers on this slide${s.entry.virtualId > 0 ? "" : " (or the slide it was copied from)"} `
                + `(${s.defects.map((d) => d.text).join("; ")}); pptc ${actions.length > 0 ? actions.join(", ") : "repaired them"}`))
    }
}

/**  the repair triggers of the output that the apply adds: a trigger the
     input already carried (same rule and part, as often) and pptc cannot
     repair stays as it was and is reported (W_INPUT_DEFECT)  */
const newDefects = async (after: JSZip, before: JSZip, warnings: LintWarning[]): Promise<string[]> => {
    const found = await findDefects(after, before)
    if (found.length === 0)
        return []
    const key = (f: Finding): string => `${f.rule}|${f.part}`
    const inherited = new Map<string, number>()
    for (const f of await findDefects(before))
        inherited.set(key(f), (inherited.get(key(f)) ?? 0) + 1)
    const kept: Finding[] = []
    const added: Finding[] = []
    for (const f of found) {
        const left = inherited.get(key(f)) ?? 0

        /*  a part the input did not have inherits nothing  */
        if (left > 0 && findPart(before, f.part) !== null) {
            inherited.set(key(f), left - 1)
            kept.push(f)
        }
        else
            added.push(f)
    }
    if (kept.length > 0)
        warnings.push({ code: "W_INPUT_DEFECT", slide: DECK_WIDE,
            message: `the input deck carries PowerPoint repair triggers pptc cannot repair; they stay as they were: ${kept.map((f) => f.text).join("; ")}` })
    return added.map((f) => f.text)
}

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
    const deck = await readDeck(pkg, warnings)
    const renumbered = renumberSlideIds(pkg, deck, warnings)
    const matched = matchInput(plan, deck)
    const removed = removedSlides(pkg, deck, matched, warnings)

    /*  structure, classification, edits, slide list, package-wide steps  */
    const parts = await buildSlides(pkg, plan, matched)
    const states = await classifySlides(pkg, plan, parts, removed)
    const untouched = untouchedParts(pkg, states)
    await editSlides(pkg, plan, deck, states, removed, warnings)
    writeSlideList(pkg, deck, states, matched, removed, renumbered || deck.dropped)
    await removeSlides(pkg, removed)
    await finishPackage(pkg.zip, parts, plan.props, plan.customProps, sectionsOut(plan, deck))
    const finalBytes = await pkg.zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })

    /*  self-checks before writing: repair triggers the apply adds, the
        content of every touched slide, every part of every untouched slide  */
    const after = await JSZip.loadAsync(finalBytes)
    const expects = expectations(states)
    const content = await verifyContent(archive.zip, after, expects, repaired)
    const findings = [
        ...await newDefects(after, archive.zip, warnings),
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
    const replaced = states.flatMap((s, index) => s.replaced.map((r) => ({ index, ...r })))
    return { slideCount: plan.entries.length, refIndexes, warnings, replaced }
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
