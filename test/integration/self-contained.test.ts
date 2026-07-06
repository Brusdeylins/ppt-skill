/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the "self-contained deck" features -- custom document
**  properties (docProps/custom.xml round-trip, merge, no-duplication) and
**  seed-from-deck (slide.add without an external --template, using the
**  deck's OWN embedded layouts).
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { expectIntact } from "../util/integrity.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { cmdState } from "../../src/commands/state.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

const withTpl: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }
const noTpl: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  count non-overlapping occurrences of a literal substring in a zip part  */
const countIn = async (file: string, part: string, needle: string): Promise<number> => {
    const zip = await JSZip.loadAsync(readFileSync(file))
    const text = await zip.file(part)?.async("string") ?? ""
    return text.split(needle).length - 1
}

describe("custom document properties", () => {
    const DECK = path.join(TMP, "custom-props.pptx")

    beforeAll(async () => {
        mkdirSync(TMP, { recursive: true })
        writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Base" } } }]
        }, withTpl)
    })

    it("writes custom props into the deck and reads them back", async () => {
        await executeOps(DECK, {
            ops: [{ op: "meta.props", set: { custom: { pptcImageStyle: "Pencil Sketch", pptcInfoStyle: "Minimal" } } }]
        }, withTpl)
        const state = await readDeckState(await DeckArchive.open(DECK))
        expect(state.customProps.pptcImageStyle).toBe("Pencil Sketch")
        expect(state.customProps.pptcInfoStyle).toBe("Minimal")
        await expectIntact(DECK)
    })

    it("merges on re-apply: patched keys update, others persist", async () => {
        await executeOps(DECK, {
            ops: [{ op: "meta.props", set: { custom: { pptcImageStyle: "Cinematic" } } }]
        }, withTpl)
        const state = await readDeckState(await DeckArchive.open(DECK))
        expect(state.customProps.pptcImageStyle).toBe("Cinematic")
        expect(state.customProps.pptcInfoStyle).toBe("Minimal")
        await expectIntact(DECK)
    })

    it("never duplicates the part wiring across applies", async () => {
        /*  exactly one content-type override, one package relationship and one
            part -- otherwise PowerPoint demands a repair  */
        expect(await countIn(DECK, "[Content_Types].xml", "PartName=\"/docProps/custom.xml\"")).toBe(1)
        expect(await countIn(DECK, "_rels/.rels", "custom-properties")).toBe(1)
        const zip = await JSZip.loadAsync(readFileSync(DECK))
        expect(zip.file("docProps/custom.xml")).not.toBeNull()
    })

    it("the `state` command surfaces customProps in its envelope", async () => {
        const out = await cmdState([DECK, "--level", "summary"])
        const result = out.result as { customProps: Record<string, string> }
        expect(result.customProps.pptcImageStyle).toBe("Cinematic")
        expect(result.customProps.pptcInfoStyle).toBe("Minimal")
    })

    it("sets core and custom properties in one op", async () => {
        await executeOps(DECK, {
            ops: [{ op: "meta.props", set: { title: "My Deck", custom: { pptcTopic: "Skills" } } }]
        }, withTpl)
        const core = await (await JSZip.loadAsync(readFileSync(DECK))).file("docProps/core.xml")?.async("string") ?? ""
        expect(core).toContain("My Deck")
        const state = await readDeckState(await DeckArchive.open(DECK))
        expect(state.customProps.pptcTopic).toBe("Skills")
        expect(state.customProps.pptcImageStyle).toBe("Cinematic")
        await expectIntact(DECK)
    })
})

describe("seed-from-deck (slide.add without a template)", () => {
    const DECK = path.join(TMP, "seed-from-deck.pptx")

    beforeAll(async () => {
        mkdirSync(TMP, { recursive: true })
        writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "First" } } }]
        }, withTpl)
    })

    it("adds a slide using the deck's own layouts, no --template", async () => {
        const before = await readDeckState(await DeckArchive.open(DECK))
        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Second" } } }]
        }, noTpl)
        const after = await readDeckState(await DeckArchive.open(DECK))
        expect(after.slides.length).toBe(before.slides.length + 1)
        expect(after.slides[after.slides.length - 1]?.title).toBe("Second")
        expect(after.slides[after.slides.length - 1]?.layoutName).toBe("CONTENT")
        await expectIntact(DECK)
    })

    it("resolves a layout by name from the deck and stays intact", async () => {
        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "TITLE_SLIDE", placeholders: { title: { text: "Cover" } } }]
        }, noTpl)
        await expectIntact(DECK)
    })
})
