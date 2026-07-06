/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/registry: the op contract and the mutation plan model. Every op
**  is a pure plan transformer: it validates its input against the current
**  virtual deck and rewrites the planned slide list. The engine session later
**  interprets the finished plan in one pass -- ops never touch files.
*/

import { PptcError } from "../../infra/errors.js"
import type { Op } from "../../schema/ops.js"
import type { ElementSpec, RichText } from "../../schema/payloads.js"
import type { DeckState, Layout, TemplateInfo } from "../model.js"
import type { LintWarning } from "../lint.js"
import type { SelectableSlide } from "../selector.js"

/**  one planned placeholder fill on a slide  */
export interface PlannedFill {
    /**  OOXML placeholder idx  */
    phIdx: number
    /**  shape name used to address the placeholder in the engine  */
    phName: string
    /**  text content to write  */
    text?: RichText
    /**  image file to insert into a picture placeholder  */
    image?: string
    /**  geometry of the placeholder (for image insertion), inches  */
    frame?: { x: number, y: number, w: number, h: number }
    /**  append to existing content instead of replacing  */
    append: boolean
}

/**  one element to generate on a slide, optionally named  */
export interface PlannedElement {
    /**  shape name assigned to the generated object, null for default  */
    name: string | null
    /**  validated element payload  */
    spec: ElementSpec
}

/**  the planned future of one output slide  */
export interface SlidePlanEntry {
    /**  where the slide content comes from  */
    source: { kind: "self", part: string } | { kind: "seed", layoutIndex: number }
    /**  virtual id during planning: real sldId for kept slides, negative for new  */
    virtualId: number
    /**  tracked title (for title: selectors within the same ops document)  */
    title: string | null
    /**  layout index the slide uses (-1 when unknown)  */
    layoutIndex: number
    /**  placeholder fills to apply  */
    fills: PlannedFill[]
    /**  speaker notes to set, null = leave untouched  */
    notes: string | null
    /**  footer text to set, null = leave untouched  */
    footer: string | null
    /**  solid background color (RRGGBB), null = leave untouched  */
    background: string | null
    /**  whether the output slide is hidden ("Hide Slide", `show="0"`). Kept
         slides inherit the source slide's state; new slides default visible  */
    hidden: boolean
    /**  elements to generate via PptxGenJS  */
    elements: PlannedElement[]
    /**  existing elements to retext: shape name to new content  */
    setTexts: { name: string, text: RichText }[]
    /**  existing elements to remove, by shape name  */
    removeNames: string[]
}

/**
 *  Create a fresh plan entry with all mutation fields at their defaults.
 *
 *  @param source - where the slide content comes from
 *  @param virtualId - virtual id during planning
 *  @param title - tracked slide title, null when unknown
 *  @param layoutIndex - layout index the slide uses (-1 when unknown)
 *  @param hidden - initial hidden state (kept slides inherit it; defaults to
 *                  false for new slides)
 *  @returns a plan entry with empty work lists
 */
export const newPlanEntry = (
    source: SlidePlanEntry["source"],
    virtualId: number,
    title: string | null,
    layoutIndex: number,
    hidden = false
): SlidePlanEntry => ({
    source,
    virtualId,
    title,
    layoutIndex,
    fills: [],
    notes: null,
    footer: null,
    background: null,
    hidden,
    elements: [],
    setTexts: [],
    removeNames: []
})

/**  the complete mutation plan an apply run executes  */
export interface MutationPlan {
    /**  output slides in final order  */
    entries: SlidePlanEntry[]
    /**  document core-property patch, null when untouched  */
    props: Record<string, string> | null
    /**  custom document-property patch (name→value), null when untouched  */
    customProps: Record<string, string> | null
    /**  lint findings collected during planning  */
    warnings: LintWarning[]
    /**  refs declared by ops, mapped to their entry  */
    refs: Map<string, SlidePlanEntry>
}

/**  shared mutable state all op transformers operate on  */
export interface PlanContext {
    /**  the evolving plan  */
    plan: MutationPlan
    /**  read model of the deck as it exists on disk  */
    deck: DeckState
    /**  layouts of the deck itself (for fills on existing slides)  */
    deckLayouts: Layout[]
    /**  resolved template info, null when no --template was given  */
    template: TemplateInfo | null
    /**  next virtual id for newly created slides (negative, decreasing)  */
    nextVirtualId: number
    /**  font-size floor in pt for the readability lint (0 disables)  */
    minFontPt: number
}

/**  contract every op module fulfills  */
export interface OpHandler<T extends Op = Op> {
    /**  op discriminator, e.g. "slide.fill"  */
    name: T["op"]
    /**
     *  Rewrite the plan according to one op.
     *
     *  @param ctx - shared planning context
     *  @param op - the validated op payload
     */
    plan(ctx: PlanContext, op: T): void
}

/**
 *  View the current entries as selectable slides (selector resolution).
 *
 *  @param ctx - planning context
 *  @returns selectable view in current plan order
 */
export const selectableEntries = (ctx: PlanContext): (SelectableSlide & { entry: SlidePlanEntry })[] =>
    ctx.plan.entries.map((entry, index) => ({
        id: entry.virtualId,
        index,
        title: entry.title,
        entry
    }))

/**
 *  Register a document-local ref against a plan entry, rejecting duplicates.
 *
 *  @param ctx - planning context
 *  @param ref - the ref name, or undefined when the op declares none
 *  @param entry - the plan entry the ref points at
 *  @throws PptcError E_SCHEMA on a duplicate ref
 */
export const registerRef = (ctx: PlanContext, ref: string | undefined, entry: SlidePlanEntry): void => {
    if (ref === undefined)
        return
    if (ctx.plan.refs.has(ref))
        throw new PptcError("E_SCHEMA", `duplicate ref '${ref}' in ops document`)
    ctx.plan.refs.set(ref, entry)
}
