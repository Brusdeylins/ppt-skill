/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Contract: `pptc schema <op>` answers for every op the index lists, and
**  the index lists exactly the ops the ops document accepts.
*/

import { describe, expect, it } from "vitest"
import { cmdSchema } from "../../src/commands/schema.js"
import { OpSchema, OP_NAMES } from "../../src/schema/ops.js"

describe("schema command contract", () => {
    it("OP_NAMES lists exactly the members of the op union", () => {
        const union = OpSchema.options.map((s) => s.shape.op.value).sort()
        expect([ ...OP_NAMES ].sort()).toEqual(union)
    })

    it.each([ ...OP_NAMES ])("pptc schema %s returns its JSON Schema", (op) => {
        const out = cmdSchema([ op ]) as { result: { op: string, schema: Record<string, unknown> } }
        expect(out.result.op).toBe(op)
        expect(JSON.stringify(out.result.schema)).toContain(`"${op}"`)
    })

    it("pptc schema document returns the whole ops document", () => {
        const out = cmdSchema([ "document" ]) as { result: { schema: Record<string, unknown> } }
        expect(JSON.stringify(out.result.schema)).toContain("\"ops\"")
    })
})
