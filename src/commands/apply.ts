/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  commands/apply: the single write path. Validates the ops document as a
**  whole, plans every mutation against the current deck state, lints, and
**  only then executes the engine session with one atomic write.
*/

import { PptcError } from "../infra/errors.js"
import { planOps } from "../core/ops/planner.js"
import { OpsDocumentSchema, type OpsDocument } from "../schema/ops.js"
import { DeckArchive, readDeckState, readTemplateInfo } from "../engine/reader.js"
import { runSession } from "../engine/session.js"
import { parseJson, requireFile, resolvePayload } from "../infra/fs.js"
import { parse } from "../infra/args.js"
import type { LintWarning } from "../core/lint.js"

/**  options controlling one ops execution  */
export interface ExecuteOptions {
    /**  template for slide.add ops, null when none given  */
    templatePath: string | null
    /**  validate and plan only, write nothing  */
    dryRun: boolean
    /**  escalate lint warnings to a hard failure (exit 7)  */
    strict: boolean
    /**  expected deck revision (optimistic lock), overrides the document  */
    expectRev: string | null
    /**  output file, defaults to the deck file itself  */
    outFile: string | null
    /**  readability font-size floor in pt (0 disables); defaults to 11  */
    minFontPt?: number
}

/**  result payload of an ops execution  */
export interface ExecuteResult {
    file: string
    rev: { before: string, after: string | null }
    result: {
        applied: number
        slideCount: number
        dryRun: boolean
        slides: Record<string, { id: number, index: number }>
    }
    warnings: LintWarning[]
}

/**
 *  Validate, plan and execute an ops document against a deck.
 *
 *  @param deckFile - the deck to modify
 *  @param rawDoc - unvalidated ops document (already JSON-parsed)
 *  @param opts - execution options
 *  @returns the result payload for the envelope
 *  @throws PptcError on any validation, addressing, lint or engine failure;
 *          the deck file is untouched in every failure case
 */
export const executeOps = async (
    deckFile: string,
    rawDoc: unknown,
    opts: ExecuteOptions
): Promise<ExecuteResult> => {
    requireFile(deckFile, "deck")
    const parsedDoc = OpsDocumentSchema.safeParse(rawDoc)
    if (!parsedDoc.success)
        throw new PptcError("E_SCHEMA", "ops document failed validation",
            { issues: parsedDoc.error.issues })
    const doc: OpsDocument = parsedDoc.data
    if (opts.expectRev !== null)
        doc.expectRev = opts.expectRev

    /*  read current state and resolve layouts  */
    const archive = await DeckArchive.open(deckFile)
    const deck = await readDeckState(archive)
    const deckInfo = await readTemplateInfo(archive)
    let template = null
    if (opts.templatePath !== null) {
        requireFile(opts.templatePath, "template")
        template = await readTemplateInfo(await DeckArchive.open(opts.templatePath))
    }

    /*  plan everything before touching anything  */
    const plan = planOps(doc, deck, deckInfo.layouts, template, opts.minFontPt ?? 11)
    if (opts.strict && plan.warnings.length > 0)
        throw new PptcError("E_LINT",
            `${plan.warnings.length} lint finding(s) under --strict`,
            { warnings: plan.warnings })

    if (opts.dryRun)
        return {
            file: deck.file,
            rev: { before: deck.rev, after: null },
            result: { applied: doc.ops.length, slideCount: plan.entries.length, dryRun: true, slides: {} },
            warnings: plan.warnings
        }

    /*  execute and report the final identities  */
    const outFile = opts.outFile ?? deckFile
    const session = await runSession(deckFile, outFile, plan, opts.templatePath)
    const finalState = await readDeckState(await DeckArchive.open(outFile))
    const slides: Record<string, { id: number, index: number }> = {}
    for (const [ref, index] of Object.entries(session.refIndexes)) {
        const slide = finalState.slides[index]
        if (slide !== undefined)
            slides[ref] = { id: slide.id, index: slide.index }
    }
    return {
        file: finalState.file,
        rev: { before: deck.rev, after: finalState.rev },
        result: { applied: doc.ops.length, slideCount: session.slideCount, dryRun: false, slides },
        warnings: plan.warnings
    }
}

/**
 *  CLI command `pptc apply <deck> (--ops @file | -e <op>) [...]`.
 *
 *  @param argv - raw arguments after the command name
 *  @returns the result payload for the envelope
 */
export const cmdApply = async (argv: string[]): Promise<ExecuteResult> => {
    const args = parse(argv, {
        "ops": { type: "string" },
        "expr": { type: "string", short: "e" },
        "template": { type: "string" },
        "dry-run": { type: "boolean" },
        "strict": { type: "boolean" },
        "rev": { type: "string" },
        "out": { type: "string" },
        "min-font-pt": { type: "string" }
    }, ["deck"])
    const opsArg = args.str("ops")
    const exprArg = args.str("expr")
    if ((opsArg === null) === (exprArg === null))
        throw new PptcError("E_USAGE", "pass exactly one of --ops <@file|-> or -e '<op-json>'")
    const rawDoc = opsArg !== null
        ? parseJson(resolvePayload(opsArg))
        : { ops: [parseJson(exprArg as string)] }
    /*  a bare op list is accepted as shorthand for { ops: [...] }  */
    const doc = Array.isArray(rawDoc) ? { ops: rawDoc } : rawDoc
    const minFontRaw = args.str("min-font-pt")
    const minFontPt = minFontRaw === null ? 11 : Number(minFontRaw)
    if (!Number.isFinite(minFontPt) || minFontPt < 0)
        throw new PptcError("E_USAGE", "--min-font-pt must be a non-negative number (0 disables)")
    return await executeOps(args.positionals[0] as string, doc, {
        templatePath: args.str("template"),
        dryRun: args.flag("dry-run"),
        strict: args.flag("strict"),
        expectRev: args.str("rev"),
        outFile: args.str("out"),
        minFontPt
    })
}
