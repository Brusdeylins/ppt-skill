/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: slide visibility ("Hide Slide", show="0") must survive an
**  apply. Regression guard for the defect where the automizer rebuild dropped
**  the <p:sld show="0"> attribute and silently un-hid every hidden slide.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps } from "../../src/commands/apply.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "visibility.pptx")

const opts = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  set show="0" on the `p:sld` root of one slide part inside the deck zip --
     emulating a deck a user hid in PowerPoint  */
const hideSlidePart = async (file: string, part: string): Promise<void> => {
    const zip = await JSZip.loadAsync(readFileSync(file))
    const xml = await (zip.file(part) as JSZip.JSZipObject).async("string")
    zip.file(part, xml.replace(/<p:sld([ >])/, "<p:sld show=\"0\"$1"))
    writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }))
}

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    await executeOps(DECK, {
        ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Hidden Slide" } } },
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Visible Slide" } } }
        ]
    }, opts)
})

describe("slide visibility is preserved across applies", () => {
    it("reads the hidden flag from the show=\"0\" attribute", async () => {
        const before = await readDeckState(await DeckArchive.open(DECK))
        expect(before.slides.map((s) => s.hidden)).toEqual([false, false])

        await hideSlidePart(DECK, (before.slides[0] as { part: string }).part)

        const marked = await readDeckState(await DeckArchive.open(DECK))
        expect(marked.slides[0]?.hidden).toBe(true)
        expect(marked.slides[1]?.hidden).toBe(false)
    })

    it("keeps an untouched hidden slide hidden after editing another slide", async () => {
        /*  edit the VISIBLE slide; the loss was global, so the untouched
            hidden slide is the real test  */
        await executeOps(DECK, {
            ops: [{ op: "slide.fill", slide: "title:Visible Slide", placeholders: { body: { text: "edited" } } }]
        }, opts)

        const after = await readDeckState(await DeckArchive.open(DECK))
        expect(after.slides.find((s) => s.title === "Hidden Slide")?.hidden).toBe(true)
        expect(after.slides.find((s) => s.title === "Visible Slide")?.hidden).toBe(false)
        await expectIntact(DECK)
    })
})
