/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/ops/slide-edit: plan transformers for the slide-level edit ops --
**  slide.fill (content), slide.rm, slide.move and slide.copy (structure).
*/

import { PptcError } from "../../infra/errors.js"
import type { SlideFillOp, Op } from "../../schema/ops.js"
import { registerRef, type OpHandler, type SlidePlanEntry } from "./registry.js"
import { planFill, resolveEntry } from "./fill-common.js"

/**  op handler: slide.fill  */
export const slideFill: OpHandler<SlideFillOp> = {
    name: "slide.fill",
    plan(ctx, op): void {
        planFill(ctx, resolveEntry(ctx, op.slide), op)
    }
}

/**  op handler: slide.rm  */
export const slideRm: OpHandler<Extract<Op, { op: "slide.rm" }>> = {
    name: "slide.rm",
    plan(ctx, op): void {
        const entry = resolveEntry(ctx, op.slide)
        ctx.plan.entries.splice(ctx.plan.entries.indexOf(entry), 1)
        for (const [name, ref] of ctx.plan.refs)
            if (ref === entry)
                ctx.plan.refs.delete(name)
    }
}

/**  op handler: slide.move  */
export const slideMove: OpHandler<Extract<Op, { op: "slide.move" }>> = {
    name: "slide.move",
    plan(ctx, op): void {
        const entry = resolveEntry(ctx, op.slide)
        if (op.to >= ctx.plan.entries.length)
            throw new PptcError("E_ADDR_NOTFOUND",
                `move target ${op.to} out of range (0-${ctx.plan.entries.length - 1})`)
        ctx.plan.entries.splice(ctx.plan.entries.indexOf(entry), 1)
        ctx.plan.entries.splice(op.to, 0, entry)
    }
}

/**  op handler: slide.copy  */
export const slideCopy: OpHandler<Extract<Op, { op: "slide.copy" }>> = {
    name: "slide.copy",
    plan(ctx, op): void {
        const source = resolveEntry(ctx, op.slide)
        const copy: SlidePlanEntry = {
            ...source,
            virtualId: ctx.nextVirtualId--,
            fills: structuredClone(source.fills),
            /*  deep-copy planned elements: a later el.set on the copy must
                not patch the original's planned spec through aliasing  */
            elements: source.elements.map((e) => ({ name: e.name, spec: structuredClone(e.spec) })),
            edits: structuredClone(source.edits),
            removeNames: [...source.removeNames],
            replaces: structuredClone(source.replaces),
            replaced: structuredClone(source.replaced)
        }
        ctx.plan.entries.splice(ctx.plan.entries.indexOf(source) + 1, 0, copy)
        registerRef(ctx, op.ref, copy)
    }
}
