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
import {
    ElAddSchema, ElRmSchema, ElSetSchema, ImgPromptsSchema, MetaPropsSchema,
    OpsDocumentSchema, OP_NAMES, SlideAddSchema, SlideCopySchema,
    SlideFillSchema, SlideMoveSchema, SlideRmSchema, TextReplaceSchema
} from "../schema/ops.js"

/**  op name to schema mapping  */
const SCHEMAS: Record<string, ZodType> = {
    "slide.add": SlideAddSchema,
    "slide.fill": SlideFillSchema,
    "slide.rm": SlideRmSchema,
    "slide.move": SlideMoveSchema,
    "slide.copy": SlideCopySchema,
    "el.add": ElAddSchema,
    "el.rm": ElRmSchema,
    "el.set": ElSetSchema,
    "img.prompts": ImgPromptsSchema,
    "meta.props": MetaPropsSchema,
    "text.replace": TextReplaceSchema,
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
