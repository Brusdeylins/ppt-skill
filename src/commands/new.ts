/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  commands/new: create a deck from a template -- an empty .pptx carrying the
**  template's masters, layouts and theme. With --ops, the whole deck is built
**  in the same run; a failing ops document leaves no file behind.
*/

import { existsSync, rmSync } from "node:fs"
import path from "node:path"
import { PptcError } from "../infra/errors.js"
import { buildEmptyDeck } from "../engine/seed.js"
import { atomicWrite, cacheDir, parseJson, requireFile, resolvePayload } from "../infra/fs.js"
import { parse } from "../infra/args.js"
import { executeOps, minFontPtArg, type ExecuteResult } from "./apply.js"

/**
 *  CLI command `pptc new <deck> --template <file.potx> [--force] [--ops @file]`.
 *
 *  @param argv - raw arguments after the command name
 *  @returns creation facts, including apply results when --ops was given
 */
export const cmdNew = async (argv: string[]): Promise<Record<string, unknown>> => {
    const args = parse(argv, {
        "template": { type: "string" },
        "force": { type: "boolean" },
        "ops": { type: "string" },
        "strict": { type: "boolean" },
        "min-font-pt": { type: "string" }
    }, ["deck"])
    const deckFile = args.positionals[0] as string
    const templatePath = args.need("template")
    requireFile(templatePath, "template")
    if (existsSync(deckFile) && !args.flag("force"))
        throw new PptcError("E_FILE", `'${deckFile}' already exists -- use --force to overwrite`)

    const opsArg = args.str("ops")
    if (opsArg === null) {
        atomicWrite(path.resolve(deckFile), await buildEmptyDeck(templatePath))
        return { file: path.resolve(deckFile), result: { created: true, slideCount: 0 } }
    }

    /*  build the empty deck in the cache, apply ops, write the target last  */
    const rawDoc = parseJson(resolvePayload(opsArg))
    const doc = Array.isArray(rawDoc) ? { ops: rawDoc } : rawDoc
    const stage = path.join(cacheDir(), `pptc-new-${process.pid}.pptx`)
    try {
        atomicWrite(stage, await buildEmptyDeck(templatePath))
        const result: ExecuteResult = await executeOps(stage, doc, {
            templatePath,
            dryRun: false,
            strict: args.flag("strict"),
            expectRev: null,
            outFile: path.resolve(deckFile),
            minFontPt: minFontPtArg(args)
        })
        return { ...result, result: { ...result.result, created: true } }
    }
    finally {
        rmSync(stage, { force: true })
    }
}
