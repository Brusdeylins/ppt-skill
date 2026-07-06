/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  cli/args: argument parsing on top of node:util parseArgs -- positional
**  extraction with usage errors, typed flag access. No CLI framework.
*/

import { parseArgs, type ParseArgsOptionsConfig } from "node:util"
import { PptcError } from "./errors.js"

/**  parsed command arguments: positionals plus typed flag accessors  */
export interface Parsed {
    /**  positional arguments in order  */
    positionals: string[]
    /**  string flag value or null  */
    str(name: string): string | null
    /**  string flag value, required  */
    need(name: string): string
    /**  numeric flag value or null; E_USAGE when present but not a number  */
    num(name: string): number | null
    /**  boolean flag  */
    flag(name: string): boolean
}

/**
 *  Parse command arguments against a flag specification.
 *
 *  @param argv - raw arguments after the command name
 *  @param spec - parseArgs option config (flag name to type/short)
 *  @param positionalNames - names of expected positionals (for usage errors)
 *  @param required - how many leading positionals are mandatory
 *  @returns typed access to positionals and flags
 *  @throws PptcError E_USAGE on unknown flags or missing positionals
 */
export const parse = (
    argv: string[],
    spec: ParseArgsOptionsConfig,
    positionalNames: string[] = [],
    required: number = positionalNames.length
): Parsed => {
    let values: Record<string, unknown>, positionals: string[]
    try {
        ({ values, positionals } = parseArgs({
            args: argv, options: spec, allowPositionals: true, strict: true
        }))
    }
    catch (err) {
        throw new PptcError("E_USAGE", (err as Error).message)
    }
    if (positionals.length < required)
        throw new PptcError("E_USAGE",
            `missing argument <${positionalNames[positionals.length] ?? "arg"}>`,
            { expected: positionalNames })
    return {
        positionals,
        str: (name) => {
            const v = values[name]
            return typeof v === "string" ? v : null
        },
        need: (name) => {
            const v = values[name]
            if (typeof v !== "string")
                throw new PptcError("E_USAGE", `missing required option --${name}`)
            return v
        },
        num: (name) => {
            const v = values[name]
            if (typeof v !== "string")
                return null

            /*  Number("") is 0, so an empty value must not slip through  */
            const n = v.trim() === "" ? NaN : Number(v)
            if (!Number.isFinite(n))
                throw new PptcError("E_USAGE", `--${name} must be a number`)
            return n
        },
        flag: (name) => values[name] === true
    }
}
