/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the 'full' read level exposes enough to recreate styled
**  content without unzipping -- per-run formatting (font/size/bold/italic/
**  color) as paragraphs/runs, and a picture's media file name. Plain text
**  stays lean (no paragraphs, the flat `text` suffices).
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import type { ParaInfo, RunInfo, ShapeInfo, SlideInfo } from "../../src/core/model.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "read-rich.pptx")
const PNG = path.join(TMP, "read-rich-pixel.png")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  smallest valid 1x1 PNG  */
const PIXEL = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64")

const shapesOf = async (): Promise<ShapeInfo[]> =>
    ((await readDeckState(await DeckArchive.open(DECK))).slides[0] as SlideInfo).shapes

/**  find a shape by its base name (el.add appends a "-<uuid>" suffix)  */
const byName = (shapes: ShapeInfo[], base: string): ShapeInfo => {
    const s = shapes.find((x) => x.name === base || x.name.startsWith(`${base}-`))
    if (s === undefined)
        throw new Error(`shape '${base}' not found`)
    return s
}

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(PNG, PIXEL)
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    await executeOps(DECK, {
        ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Base" } } },
            { op: "el.add", slide: "title:Base", elements: [
                { type: "textbox", name: "Code", frame: { x: 1, y: 1, w: 8, h: 2 }, text: [
                    { runs: [{ text: "const x = 1", font: "Courier New", size: 14, bold: true }] },
                    { runs: [{ text: "plain tail" }] }
                ] },
                { type: "textbox", name: "Plain", frame: { x: 1, y: 4, w: 8, h: 1 }, text: "just plain text" },
                { type: "image", name: "Pic", frame: { x: 10, y: 1, w: 1, h: 1 }, path: PNG }
            ] }
        ]
    }, opts)
})

describe("state --level full: rich runs", () => {
    it("exposes per-run font/size/bold for a formatted body", async () => {
        const code = byName(await shapesOf(), "Code")
        expect(code.paragraphs).toBeDefined()
        const run = ((code.paragraphs as ParaInfo[])[0] as ParaInfo).runs[0] as RunInfo
        expect(run.font).toBe("Courier New")
        expect(run.size).toBe(14)
        expect(run.bold).toBe(true)
    })

    it("preserves paragraph structure (one entry per a:p)", async () => {
        const paras = byName(await shapesOf(), "Code").paragraphs as ParaInfo[]
        expect(paras.length).toBe(2)
        expect((paras[1] as ParaInfo).runs[0] as RunInfo).toMatchObject({ text: "plain tail" })
    })

    it("omits paragraphs for plain uniform text (flat text suffices)", async () => {
        const plain = byName(await shapesOf(), "Plain")
        expect(plain.text).toBe("just plain text")
        expect(plain.paragraphs).toBeUndefined()
    })
})

describe("state --level full: picture media name", () => {
    it("resolves the blip r:embed to the media file name", async () => {
        const pic = (await shapesOf()).find((s) => s.type === "picture")
        expect(pic?.image).toBeDefined()
        expect(pic?.image).toMatch(/\.png$/i)
    })
})
