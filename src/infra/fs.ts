/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  infra/fs: file-system primitives -- payload resolution (@file / stdin /
**  inline), atomic writes, the seed cache directory, and content hashing.
*/

import { readFileSync, writeFileSync, renameSync, mkdirSync, existsSync } from "node:fs"
import { tmpdir } from "node:os"
import { createHash } from "node:crypto"
import path from "node:path"
import { PptcError } from "./errors.js"

/**
 *  Resolve a payload argument to its raw string content.
 *  `@path` reads a file, `-` reads stdin, anything else is taken verbatim.
 *
 *  @param arg - raw CLI argument value
 *  @returns the payload text
 *  @throws PptcError E_FILE when a referenced file cannot be read
 */
export const resolvePayload = (arg: string): string => {
    if (arg === "-")
        return readFileSync(0, "utf8")
    if (arg.startsWith("@")) {
        const file = arg.slice(1)
        try {
            return readFileSync(file, "utf8")
        }
        catch {
            throw new PptcError("E_FILE", `cannot read payload file '${file}'`)
        }
    }
    return arg
}

/**
 *  Parse a payload string as JSON with a typed error on failure.
 *
 *  @param text - raw JSON text
 *  @returns the parsed value
 *  @throws PptcError E_JSON with the parser message
 */
export const parseJson = (text: string): unknown => {
    try {
        return JSON.parse(text)
    }
    catch (err) {
        throw new PptcError("E_JSON", `payload is not valid JSON: ${(err as Error).message}`)
    }
}

/**
 *  Write a buffer atomically: write to a sibling temp file, then rename.
 *  The target either keeps its previous content or gets the complete new one.
 *
 *  @param file - target file path
 *  @param data - content to write
 */
export const atomicWrite = (file: string, data: Buffer): void => {
    const tmp = `${file}.tmp-${process.pid}`
    writeFileSync(tmp, data)
    renameSync(tmp, file)
}

/**
 *  Scratch directory (staging of `pptc new`), created on demand.
 *
 *  @returns absolute path of the cache directory
 */
export const cacheDir = (): string => {
    const dir = path.join(tmpdir(), "pptc-cache")
    mkdirSync(dir, { recursive: true })
    return dir
}

/**
 *  Short content hash (SHA-1, 12 hex chars) used for rev tokens and cache keys.
 *
 *  @param parts - byte or string parts hashed in order
 *  @returns 12-character hex digest
 */
export const contentHash = (...parts: (Buffer | string)[]): string => {
    const h = createHash("sha1")
    for (const p of parts)
        h.update(p)
    return h.digest("hex").slice(0, 12)
}

/**
 *  Assert that a file exists, with a typed error otherwise.
 *
 *  @param file - path to check
 *  @param what - human label used in the error message
 *  @throws PptcError E_FILE when the file is missing
 */
export const requireFile = (file: string, what: string): void => {
    if (!existsSync(file))
        throw new PptcError("E_FILE", `${what} not found: '${file}'`)
}
