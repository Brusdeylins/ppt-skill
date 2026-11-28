/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/sections: PowerPoint sections in the plan. Sections are
**  contiguous runs of the slide order. The plan models them only once an
**  op or fill field touches them: from then on every entry names its
**  section, so section starts, renames and removals compose with the
**  structural ops of the same document. Until then the engine keeps the
**  deck's sections in step with the slide list by the same rule the model
**  starts from.
*/

import { PptcError } from "../../infra/errors.js"
import type { SectionRenameOp, SectionRmOp } from "../../schema/ops.js"
import type { OpHandler, PlanContext, PlannedSection, SlidePlanEntry } from "./registry.js"

/**  the name PowerPoint gives the section of the slides before the first
     section started in a deck without sections  */
export const DEFAULT_SECTION = "Default Section"

/**
 *  The sections of the plan, modelled on first use: the deck's sections
 *  with every entry filed as the engine files slides -- those that kept
 *  their relative order (the longest increasing run of their old
 *  positions) keep their section, new, copied and moved ones join the
 *  section of their predecessor (the first section when they lead).
 *
 *  @param ctx - planning context
 *  @returns the planned sections in deck order (changed in place by ops)
 */
export const planSections = (ctx: PlanContext): PlannedSection[] => {
    if (ctx.plan.sections !== null)
        return ctx.plan.sections
    const sections: PlannedSection[] = ctx.deck.sections.map((s) => ({ name: s.name, id: s.id }))
    const sectionOf = new Map<number, number>()
    const oldPos = new Map<number, number>()
    ctx.deck.sections.forEach((sec, k) => {
        for (const id of sec.slideIds)
            if (!sectionOf.has(id)) {
                sectionOf.set(id, k)
                oldPos.set(id, oldPos.size)
            }
    })
    const entries = ctx.plan.entries
    const kept = entries.filter((e) => e.virtualId > 0 && oldPos.has(e.virtualId))
    const pos = (e: SlidePlanEntry): number => oldPos.get(e.virtualId) as number

    /*  longest increasing run of old positions (O(n^2), decks are small)  */
    const len = kept.map(() => 1)
    const prev = kept.map(() => -1)
    for (let i = 0; i < kept.length; i++)
        for (let j = 0; j < i; j++)
            if (pos(kept[j] as SlidePlanEntry) < pos(kept[i] as SlidePlanEntry) && (len[j] as number) + 1 > (len[i] as number)) {
                len[i] = (len[j] as number) + 1
                prev[i] = j
            }
    const stable = new Set<SlidePlanEntry>()
    for (let i = len.indexOf(Math.max(0, ...len)); i >= 0; i = prev[i] as number)
        stable.add(kept[i] as SlidePlanEntry)
    let current = 0
    for (const entry of entries) {
        if (stable.has(entry))
            current = sectionOf.get(entry.virtualId) as number
        entry.section = sections[current] ?? null
    }
    ctx.plan.sections = sections
    return sections
}

/**
 *  File a slide that was added, copied or moved into the section of its
 *  predecessor (the first section when it leads the deck), as PowerPoint
 *  does -- once the plan models sections.
 *
 *  @param ctx - planning context
 *  @param entry - the entry at its new position
 */
export const joinNeighbour = (ctx: PlanContext, entry: SlidePlanEntry): void => {
    const sections = ctx.plan.sections
    if (sections === null)
        return
    const idx = ctx.plan.entries.indexOf(entry)
    entry.section = idx > 0 ? (ctx.plan.entries[idx - 1] as SlidePlanEntry).section : sections[0] ?? null
}

/**  the one section of a name (E_ADDR_NOTFOUND / E_ADDR_AMBIGUOUS)  */
const sectionNamed = (sections: PlannedSection[], name: string): PlannedSection => {
    const hits = sections.filter((s) => s.name === name)
    if (hits.length === 0)
        throw new PptcError("E_ADDR_NOTFOUND", `no section named '${name}'`,
            { sections: sections.map((s) => s.name) })
    if (hits.length > 1)
        throw new PptcError("E_ADDR_AMBIGUOUS", `the deck has ${hits.length} sections named '${name}'`)
    return hits[0] as PlannedSection
}

/**  reject a name another section carries already  */
const checkFree = (sections: PlannedSection[], name: string): void => {
    if (sections.some((s) => s.name === name))
        throw new PptcError("E_SCHEMA", `a section named '${name}' exists already -- section names are unique`,
            { sections: sections.map((s) => s.name) })
}

/**
 *  Start a section at a slide (fill field `section`): the new section runs
 *  from the slide up to the next section start, so a slide inside a
 *  section splits it; a section left without slides goes. In a deck
 *  without sections the slides before it form PowerPoint's "Default
 *  Section". A slide already in a section of that name stays as it is.
 *
 *  @param ctx - planning context
 *  @param entry - the slide that starts the section
 *  @param name - the section name (unique)
 *  @param added - whether the slide is added by the current op (it then
 *                 belongs to no section yet)
 *  @throws PptcError E_SCHEMA when another section carries the name
 */
export const startSection = (ctx: PlanContext, entry: SlidePlanEntry, name: string, added: boolean): void => {
    const sections = planSections(ctx)
    const entries = ctx.plan.entries
    const idx = entries.indexOf(entry)
    if (!added && entry.section?.name === name)
        return
    checkFree(sections, name)
    const fresh: PlannedSection = { name, id: null }
    if (sections.length === 0) {
        const lead: PlannedSection = { name: DEFAULT_SECTION, id: null }
        if (idx > 0) {
            checkFree([ fresh ], DEFAULT_SECTION)
            sections.push(lead)
        }
        sections.push(fresh)
        entries.forEach((e, i) => { e.section = i < idx ? lead : fresh })
        return
    }
    const from = added ? (idx > 0 ? (entries[idx - 1] as SlidePlanEntry).section : null) : entry.section
    sections.splice(from === null ? 0 : sections.indexOf(from) + 1, 0, fresh)
    entry.section = fresh
    for (let j = idx + 1; from !== null && j < entries.length && (entries[j] as SlidePlanEntry).section === from; j++)
        (entries[j] as SlidePlanEntry).section = fresh
    if (from !== null && !entries.some((e) => e.section === from))
        sections.splice(sections.indexOf(from), 1)
}

/**  op handler: section.rename  */
export const sectionRename: OpHandler<SectionRenameOp> = {
    name: "section.rename",
    plan(ctx, op): void {
        const sections = planSections(ctx)
        const section = sectionNamed(sections, op.from)
        if (op.to === op.from)
            return
        checkFree(sections, op.to)
        section.name = op.to
    }
}

/**  op handler: section.rm -- the slides stay and join the previous
     section (the next one when the removed section was the first)  */
export const sectionRm: OpHandler<SectionRmOp> = {
    name: "section.rm",
    plan(ctx, op): void {
        const sections = planSections(ctx)
        const section = sectionNamed(sections, op.name)
        const k = sections.indexOf(section)
        const into = (k > 0 ? sections[k - 1] : sections[k + 1]) ?? null
        for (const entry of ctx.plan.entries)
            if (entry.section === section)
                entry.section = into
        sections.splice(k, 1)
    }
}
