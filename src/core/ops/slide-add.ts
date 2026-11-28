/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/slide-add: plan a new slide based on a template layout,
**  optionally filling it in the same op.
*/

import { PptcError } from "../../infra/errors.js"
import type { SlideAddOp } from "../../schema/ops.js"
import { newPlanEntry, registerRef, type OpHandler, type PlanContext } from "./registry.js"
import { planFill } from "./fill-common.js"
import { joinNeighbour } from "./sections.js"

/**  resolve a layout address (index or exact name). When a template was
     given, resolve against it; otherwise resolve against the deck's OWN
     embedded layouts -- the seed is then derived from the deck itself, so
     adding slides to an existing deck needs no template file.  */
const resolveLayout = (ctx: PlanContext, layout: number | string): number => {
    const layouts = ctx.template?.layouts ?? ctx.deckLayouts
    if (layouts.length === 0)
        throw new PptcError("E_TEMPLATE",
            "'slide.add' needs layouts: this deck carries none -- pass --template <file.potx>")
    if (typeof layout === "number") {
        if (layouts[layout] === undefined)
            throw new PptcError("E_ADDR_NOTFOUND",
                `layout index ${layout} out of range (0-${layouts.length - 1})`,
                { layouts: layouts.map((l) => ({ index: l.index, name: l.name })) })
        return layout
    }
    const hit = layouts.find((l) => l.name === layout)
    if (hit === undefined)
        throw new PptcError("E_ADDR_NOTFOUND", `no layout named '${layout}'`,
            { layouts: layouts.map((l) => ({ index: l.index, name: l.name })) })
    return hit.index
}

/**  op handler: slide.add  */
export const slideAdd: OpHandler<SlideAddOp> = {
    name: "slide.add",
    plan(ctx, op): void {
        const layoutIndex = resolveLayout(ctx, op.layout)
        const entry = newPlanEntry({ kind: "seed", layoutIndex }, ctx.nextVirtualId--, null, layoutIndex)
        const at = Math.min(op.at ?? ctx.plan.entries.length, ctx.plan.entries.length)
        ctx.plan.entries.splice(at, 0, entry)
        registerRef(ctx, op.ref, entry)
        joinNeighbour(ctx, entry)
        planFill(ctx, entry, op, true)
    }
}
