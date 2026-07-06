/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the optional `hidden` field on slide.fill / slide.add lets an
**  agent deliberately hide or show a slide ("Hide Slide"), on top of the
**  visibility that is otherwise preserved untouched.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps } from "../../src/commands/apply.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "visibility-op.pptx")

const opts = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

const hiddenOf = async (title: string): Promise<boolean | undefined> => {
    const state = await readDeckState(await DeckArchive.open(DECK))
    return state.slides.find((s) => s.title === title)?.hidden
}

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    await executeOps(DECK, {
        ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "S1" } } }]
    }, opts)
})

describe("slide.fill / slide.add set slide visibility", () => {
    it("slide.fill with hidden:true hides an existing slide", async () => {
        expect(await hiddenOf("S1")).toBe(false)
        await executeOps(DECK, { ops: [{ op: "slide.fill", slide: "title:S1", hidden: true }] }, opts)
        expect(await hiddenOf("S1")).toBe(true)
        await expectIntact(DECK)
    })

    it("slide.fill with hidden:false shows it again", async () => {
        await executeOps(DECK, { ops: [{ op: "slide.fill", slide: "title:S1", hidden: false }] }, opts)
        expect(await hiddenOf("S1")).toBe(false)
    })

    it("slide.add with hidden:true creates a hidden slide", async () => {
        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "S2" } }, hidden: true }]
        }, opts)
        expect(await hiddenOf("S2")).toBe(true)
        expect(await hiddenOf("S1")).toBe(false)
        await expectIntact(DECK)
    })
})
