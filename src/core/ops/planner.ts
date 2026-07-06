/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/planner: assemble the op registry and run the plan phase. The
**  planner validates the whole ops document against the current deck state
**  and produces the complete MutationPlan -- before a single byte of the
**  deck is written. Failures here leave the deck untouched by construction.
*/

import { PptcError } from "../../infra/errors.js"
import type { Op, OpsDocument } from "../../schema/ops.js"
import type { DeckState, Layout, TemplateInfo } from "../model.js"
import { newPlanEntry, type MutationPlan, type OpHandler, type PlanContext, type SlidePlanEntry } from "./registry.js"
import { slideAdd } from "./slide-add.js"
import { slideCopy, slideFill, slideMove, slideRm } from "./slide-edit.js"
import { elAdd, elRm, elSet, imgPrompts, metaProps } from "./elements.js"

/**  the op registry: one handler per op name  */
const HANDLERS: ReadonlyMap<string, OpHandler> = new Map(
    ([slideAdd, slideFill, slideRm, slideMove, slideCopy,
        elAdd, elSet, elRm, imgPrompts, metaProps] as OpHandler[])
        .map((h) => [h.name, h]))

/**
 *  Plan an ops document against the current deck state.
 *
 *  @param doc - validated ops document
 *  @param deck - read model of the deck on disk
 *  @param deckLayouts - layouts available inside the deck itself
 *  @param template - resolved template info, null when no --template given
 *  @param minFontPt - readability font-size floor in pt (0 disables), default 11
 *  @returns the complete mutation plan
 *  @throws PptcError E_REV_CONFLICT when expectRev mismatches, or any
 *          planning error of the individual ops (deck untouched)
 */
export const planOps = (
    doc: OpsDocument,
    deck: DeckState,
    deckLayouts: Layout[],
    template: TemplateInfo | null,
    minFontPt: number = 11
): MutationPlan => {
    if (doc.expectRev !== undefined && doc.expectRev !== deck.rev)
        throw new PptcError("E_REV_CONFLICT",
            `deck revision is '${deck.rev}' but ops document expects '${doc.expectRev}' -- re-read with 'pptc state'`,
            { expected: doc.expectRev, actual: deck.rev })

    /*  start from the deck as it exists: every slide is kept unchanged  */
    const entries: SlidePlanEntry[] = deck.slides.map((slide) =>
        newPlanEntry({ kind: "self", part: slide.part }, slide.id, slide.title, slide.layoutIndex, slide.hidden))
    const ctx: PlanContext = {
        plan: { entries, props: null, customProps: null, warnings: [], refs: new Map() },
        deck,
        deckLayouts,
        template,
        nextVirtualId: -1,
        minFontPt
    }

    doc.ops.forEach((op: Op, i: number) => {
        const handler = HANDLERS.get(op.op)
        if (handler === undefined)
            throw new PptcError("E_SCHEMA", `unknown op '${op.op}'`)
        try {
            handler.plan(ctx, op)
        }
        catch (err) {
            if (err instanceof PptcError)
                throw new PptcError(err.code, `op ${i} (${op.op}): ${err.message}`,
                    { failedAt: i, op: op.op, details: err.details })
            throw err
        }
    })
    return ctx.plan
}
