/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  commands/sugar: flag-based micro-edit commands for small corrections
**  without a JSON file. Each one compiles to exactly one op and runs through
**  the identical validate/plan/apply path -- there is no second write
**  implementation behind these shortcuts.
*/

import { PptcError } from "../infra/errors.js"
import type { Op } from "../schema/ops.js"
import { DeckArchive, readDeckState } from "../engine/reader.js"
import { parse, type Parsed } from "../infra/args.js"
import { executeOps, minFontPtArg, type ExecuteResult } from "./apply.js"

/**  shared flags of all sugar commands  */
const COMMON = {
    "slide": { type: "string" as const },
    "rev": { type: "string" as const },
    "strict": { type: "boolean" as const },
    "dry-run": { type: "boolean" as const },
    "min-font-pt": { type: "string" as const }
}

/**  run sugar-built ops through the standard apply path  */
const runOps = async (args: Parsed, deck: string, ops: Op[]): Promise<ExecuteResult> =>
    await executeOps(deck, { ops }, {
        templatePath: null,
        dryRun: args.flag("dry-run"),
        strict: args.flag("strict"),
        expectRev: args.str("rev"),
        outFile: null,
        minFontPt: minFontPtArg(args)
    })

/**  run a single sugar-built op through the standard apply path  */
const runOne = async (args: Parsed, deck: string, op: Op): Promise<ExecuteResult> =>
    await runOps(args, deck, [op])

/**
 *  CLI command `pptc text <deck> --slide SEL --ph KEY "text" [--append]`.
 *  Compiles to one slide.fill op with plain text.
 *
 *  @param argv - raw arguments after the command name
 *  @returns the apply result payload
 */
export const cmdText = async (argv: string[]): Promise<ExecuteResult> => {
    const args = parse(argv, { ...COMMON, "ph": { type: "string" }, "append": { type: "boolean" } },
        ["deck", "text"])
    return await runOne(args, args.positionals[0] as string, {
        op: "slide.fill",
        slide: args.need("slide"),
        placeholders: {
            [args.str("ph") ?? "title"]: {
                text: args.positionals[1] as string,
                ...(args.flag("append") && { append: true })
            }
        }
    })
}

/**
 *  CLI command `pptc note <deck> --slide SEL "speaker notes"`.
 *
 *  @param argv - raw arguments after the command name
 *  @returns the apply result payload
 */
export const cmdNote = async (argv: string[]): Promise<ExecuteResult> => {
    const args = parse(argv, COMMON, ["deck", "text"])
    return await runOne(args, args.positionals[0] as string, {
        op: "slide.fill",
        slide: args.need("slide"),
        notes: args.positionals[1] as string
    })
}

/**
 *  CLI command `pptc footer <deck> [--slide SEL] "footer text"`.
 *  Without --slide the footer is set on every slide.
 *
 *  @param argv - raw arguments after the command name
 *  @returns the apply result payload
 */
export const cmdFooter = async (argv: string[]): Promise<ExecuteResult> => {
    const args = parse(argv, COMMON, ["deck", "text"])
    const deck = args.positionals[0] as string
    const text = args.positionals[1] as string
    const selector = args.str("slide")
    const ops: Op[] = selector !== null
        ? [{ op: "slide.fill", slide: selector, footer: text }]
        : []
    if (selector === null) {
        /*  one fill op per existing slide, addressed by stable index  */
        const deckState = await readDeckState(await DeckArchive.open(deck))
        for (const slide of deckState.slides)
            ops.push({ op: "slide.fill", slide: `id:${slide.id}`, footer: text })
        if (ops.length === 0)
            throw new PptcError("E_ADDR_NOTFOUND", "deck has no slides")
    }
    return await runOps(args, deck, ops)
}

/**
 *  CLI command `pptc rm <deck> --slide SEL`.
 *
 *  @param argv - raw arguments after the command name
 *  @returns the apply result payload
 */
export const cmdRm = async (argv: string[]): Promise<ExecuteResult> => {
    const args = parse(argv, COMMON, ["deck"])
    return await runOne(args, args.positionals[0] as string,
        { op: "slide.rm", slide: args.need("slide") })
}

/**
 *  CLI command `pptc move <deck> --slide SEL --to N`.
 *
 *  @param argv - raw arguments after the command name
 *  @returns the apply result payload
 */
export const cmdMove = async (argv: string[]): Promise<ExecuteResult> => {
    const args = parse(argv, { ...COMMON, "to": { type: "string" } }, ["deck"])
    const to = Number(args.need("to"))
    if (!Number.isInteger(to) || to < 0)
        throw new PptcError("E_USAGE", "--to expects a non-negative slide position")
    return await runOne(args, args.positionals[0] as string,
        { op: "slide.move", slide: args.need("slide"), to })
}
