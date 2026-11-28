/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  commands/schema: emit the JSON Schemas of the ops vocabulary, generated
**  straight from the Zod definitions -- the same objects that validate every
**  apply run. Nothing is documented twice.
*/

import { z, type ZodType } from "zod"
import { PptcError } from "../infra/errors.js"
import { parse } from "../infra/args.js"
import { OpSchema, OpsDocumentSchema, OP_NAMES } from "../schema/ops.js"

/**  op name to schema mapping, derived from the members of the op union so a
     new op can never miss its `pptc schema <op>` entry  */
const SCHEMAS: Record<string, ZodType> = {
    ...Object.fromEntries(OpSchema.options.map((s) => [ s.shape.op.value, s ])),
    "document": OpsDocumentSchema
}

/**
 *  CLI command `pptc schema [opName]`.
 *
 *  @param argv - raw arguments after the command name
 *  @returns JSON Schema of one op, or the index of all ops
 */
export const cmdSchema = (argv: string[]): Record<string, unknown> => {
    const args = parse(argv, {}, ["op"], 0)
    const target = args.positionals[0] ?? null
    if (target === null)
        return {
            result: {
                ops: [...OP_NAMES],
                hint: "run 'pptc schema <op>' for the full JSON Schema of one op, or 'pptc schema document' for the whole ops document"
            }
        }
    const schema = SCHEMAS[target]
    if (schema === undefined)
        throw new PptcError("E_USAGE", `unknown op '${target}'`, { ops: [...OP_NAMES, "document"] })
    return { result: { op: target, schema: z.toJSONSchema(schema, { io: "input" }) } }
}
