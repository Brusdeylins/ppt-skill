/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  commands/apply: the single write path. Validates the ops document as a
**  whole, plans every mutation against the current deck state, lints, and
**  only then executes the engine session with one atomic write.
*/

import path from "node:path"
import { PptcError } from "../infra/errors.js"
import { planOps } from "../core/ops/planner.js"
import { OpsDocumentSchema, type OpsDocument } from "../schema/ops.js"
import { DeckArchive, readDeckState, readTemplateInfo } from "../engine/reader.js"
import type { Layout } from "../core/model.js"
import { preflight, runSession } from "../engine/inplace/engine.js"
import { IMAGE_TYPES } from "../engine/parts.js"
import { parseJson, requireFile, resolvePayload } from "../infra/fs.js"
import { parse, type Parsed } from "../infra/args.js"
import { DEFAULT_MIN_FONT_PT, type LintWarning } from "../core/lint.js"

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
    /**  readability font-size floor in pt (0 disables)  */
    minFontPt: number
}

/**
 *  Resolve the shared `--min-font-pt` flag of the write commands.
 *
 *  @param args - parsed command arguments
 *  @returns the floor in pt; the default when the flag is absent
 *  @throws PptcError E_USAGE on a non-numeric or negative value
 */
export const minFontPtArg = (args: Parsed): number => {
    const minFontPt = args.num("min-font-pt") ?? DEFAULT_MIN_FONT_PT
    if (minFontPt < 0)
        throw new PptcError("E_USAGE", "--min-font-pt must be a non-negative number (0 disables)")
    return minFontPt
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
        /**  per text.replace op (not on --dry-run): its position in the
             ops document, the replacements it made in the output slides
             and the slides they are on  */
        replacements?: Replacement[]
    }
    warnings: LintWarning[]
}

/**  the outcome of one text.replace op  */
export interface Replacement {
    /**  position of the op in the ops document  */
    op: number
    /**  replacements in the output deck  */
    count: number
    /**  the output slides with replacements  */
    slides: { id: number, index: number, count: number }[]
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

        /*  a seed slide is bound to the deck's layout of the SAME index, so a
            template of another design would put its placeholders on a foreign
            layout (pptc never imports layouts into an existing deck)  */
        const shape = (layouts: Layout[]): string[] => layouts.map((l) =>
            `${l.name}|${l.placeholders.map((p) => `${p.idx}:${p.kind}`).sort().join(",")}`)
        const deckShape = shape(deckInfo.layouts)
        const tplShape = shape(template.layouts)
        const differs = deckShape.findIndex((l, i) => l !== tplShape[i])
        if (deckShape.length > 0 && (differs >= 0 || deckShape.length !== tplShape.length)) {
            const at = differs >= 0 ? differs : Math.min(deckShape.length, tplShape.length)
            throw new PptcError("E_TEMPLATE",
                "--template does not match the deck's layouts (another design or version); "
                + "omit --template -- the deck carries its own layouts",
                { layoutIndex: at, deck: deckShape[at] ?? null, template: tplShape[at] ?? null })
        }
    }

    /*  plan everything before touching anything  */
    const plan = planOps(doc, deck, deckInfo.layouts, template, opts.minFontPt)
    /*  image files are read by the engine only: check them now, so that
        --dry-run already reports a missing or unsupported file  */
    for (const entry of plan.entries) {
        for (const fill of entry.fills)
            if (fill.image !== undefined) {
                requireFile(fill.image, "image")
                const ext = path.extname(fill.image).slice(1).toLowerCase()
                if (ext !== "" && IMAGE_TYPES[ext] === undefined)
                    throw new PptcError("E_SCHEMA", `unsupported image type '.${ext}' for a picture placeholder`,
                        { supported: Object.keys(IMAGE_TYPES) })
            }
        for (const el of entry.elements)
            if (el.spec.type === "image")
                requireFile(el.spec.path, "image")
        for (const { set } of entry.edits)
            if (set.image !== undefined) {
                requireFile(set.image, "image")
                const ext = path.extname(set.image).slice(1).toLowerCase()
                if (ext !== "" && IMAGE_TYPES[ext] === undefined)
                    throw new PptcError("E_SCHEMA", `unsupported image type '.${ext}' for a picture`,
                        { supported: Object.keys(IMAGE_TYPES) })
            }
    }
    if (opts.strict && plan.warnings.length > 0)
        throw new PptcError("E_LINT",
            `${plan.warnings.length} lint finding(s) under --strict`,
            { warnings: plan.warnings })

    if (opts.dryRun) {
        await preflight(plan, archive)
        return {
            file: deck.file,
            rev: { before: deck.rev, after: null },
            result: { applied: doc.ops.length, slideCount: plan.entries.length, dryRun: true, slides: {} },
            warnings: plan.warnings
        }
    }

    /*  execute and report the final identities  */
    const outFile = opts.outFile ?? deckFile
    const session = await runSession(outFile, plan, archive)
    const finalState = await readDeckState(await DeckArchive.open(outFile))
    const slides: Record<string, { id: number, index: number }> = {}
    for (const [ref, index] of Object.entries(session.refIndexes)) {
        const slide = finalState.slides[index]
        if (slide !== undefined)
            slides[ref] = { id: slide.id, index: slide.index }
    }

    /*  text.replace: planned content counted while planning, the slides'
        existing content by the engine  */
    const replaceOps = doc.ops.flatMap((op, i) => op.op === "text.replace" ? [ i ] : [])
    const replacements = replaceOps.map((op): Replacement => {
        const perSlide = new Map<number, number>()
        const add = (index: number, count: number): void => {
            perSlide.set(index, (perSlide.get(index) ?? 0) + count)
        }
        plan.entries.forEach((entry, index) => {
            for (const r of entry.replaced)
                if (r.op === op)
                    add(index, r.count)
        })
        for (const r of session.replaced)
            if (r.op === op)
                add(r.index, r.count)
        const hits = [ ...perSlide ].filter(([ , count ]) => count > 0).sort((a, b) => a[0] - b[0])
        return {
            op,
            count: hits.reduce((sum, [ , count ]) => sum + count, 0),
            slides: hits.map(([ index, count ]) => ({ id: finalState.slides[index]?.id ?? 0, index, count }))
        }
    })
    return {
        file: finalState.file,
        rev: { before: deck.rev, after: finalState.rev },
        result: { applied: doc.ops.length, slideCount: session.slideCount, dryRun: false, slides,
            ...(replacements.length > 0 && { replacements }) },
        warnings: [ ...plan.warnings, ...session.warnings ]
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
    return await executeOps(args.positionals[0] as string, doc, {
        templatePath: args.str("template"),
        dryRun: args.flag("dry-run"),
        strict: args.flag("strict"),
        expectRev: args.str("rev"),
        outFile: args.str("out"),
        minFontPt: minFontPtArg(args)
    })
}
