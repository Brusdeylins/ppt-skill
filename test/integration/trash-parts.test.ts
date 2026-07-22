/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: parts under the reserved OPC folder [trash]/ (logically
**  deleted parts left by PowerPoint incremental save and OneDrive/SharePoint
**  co-authoring) must neither fail verification nor survive an apply.
**  Regression guard for the defect where the every-part-needs-a-content-type
**  self-check flagged [trash]/NNNN.dat and every apply died with E_INTEGRITY.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { executeOps } from "../../src/commands/apply.js"
import { verifyFile } from "../../src/engine/verify.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "trash-parts.pptx")

const opts = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  inject [trash]/ parts into the deck zip -- emulating a deck that came
     through OneDrive/SharePoint sync or PowerPoint co-authoring  */
const injectTrashParts = async (file: string): Promise<void> => {
    const zip = await JSZip.loadAsync(readFileSync(file))
    zip.file("[trash]/0000.dat", Buffer.from([0x00, 0x01, 0x02, 0x03]))
    zip.file("[trash]/0001.dat", "<p:sld>stale deleted slide</p:sld>")
    writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }))
}

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    await executeOps(DECK, {
        ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "First" } } }]
    }, opts)
    await injectTrashParts(DECK)
})

describe("[trash]/ parts from OneDrive/co-authoring decks", () => {
    it("verify ignores [trash]/ parts (no content type by design)", async () => {
        expect(await verifyFile(DECK)).toEqual([])
    })

    it("apply succeeds and drops [trash]/ parts from the output", async () => {
        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Second" } } }]
        }, opts)
        const zip = await JSZip.loadAsync(readFileSync(DECK))
        expect(Object.keys(zip.files).filter((f) => f.startsWith("[trash]/"))).toEqual([])
        await expectIntact(DECK)
    })
})
