/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Golden: fixed op sequences exercising all ten ops; after every apply the
**  per-slide fingerprint must equal the one frozen from the 1.0.11 engine
**  (pptx-automizer) before it was replaced by the in-place engine. The
**  neutral fixture's snapshots are committed; those of the demo deck and
**  the private templates stay local (gitignored private-templates/golden/),
**  as their content must never be published.
*/

import { describe, expect, it } from "vitest"
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState, readTemplateInfo } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { expectIntact } from "../util/integrity.js"
import { fingerprint, propsOf } from "../util/fingerprint.js"
import { existingScenario, fixedScenario, writeMedia, type Step } from "../util/scenarios.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const FIXTURE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const DEMO = path.join(here, "..", "..", "demo", "kuehlschrank-funktionsweise.pptx")
const PRIVATE = path.join(here, "..", "..", "private-templates")
const LOCAL = path.join(PRIVATE, "golden")
const TMP = path.join(here, "..", "tmp", "golden")

const opts: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  run the steps on a copy of a deck, fingerprinting after every apply  */
const run = async (start: string, tag: string, steps: Step[]): Promise<unknown[]> => {
    mkdirSync(TMP, { recursive: true })
    const deck = path.join(TMP, `${tag}.pptx`)
    copyFileSync(start, deck)
    const prints: unknown[] = []
    for (const step of steps) {
        const ops = step(await readDeckState(await DeckArchive.open(deck)))
        await executeOps(deck, { ops } as never, opts)
        await expectIntact(deck)
        prints.push({ ops: JSON.parse(JSON.stringify(ops).split(TMP).join("<tmp>")) as unknown,
            slides: await fingerprint(deck), props: await propsOf(deck) })
    }
    return prints
}

/**  compare with a local, uncommitted golden file (written on first run)  */
const matchLocal = (name: string, prints: unknown[]): void => {
    const file = path.join(LOCAL, `${name}.json`)
    const actual = JSON.parse(JSON.stringify(prints)) as unknown
    if (!existsSync(file)) {
        mkdirSync(LOCAL, { recursive: true })
        writeFileSync(file, `${JSON.stringify(actual, null, 2)}\n`)
        return
    }
    expect(actual).toEqual(JSON.parse(readFileSync(file, "utf8")))
}

describe("golden engine results", () => {
    const media = writeMedia(TMP)

    it("matches on the neutral fixture", async () => {
        const info = await readTemplateInfo(await DeckArchive.open(FIXTURE))
        expect(await run(FIXTURE, "fixture", [ ...fixedScenario(info.layouts, media), ...existingScenario(media) ])).toMatchSnapshot()
    }, 120000)

    it("matches on an empty deck from the fixture", async () => {
        mkdirSync(TMP, { recursive: true })
        const empty = path.join(TMP, "fixture-empty-start.pptx")
        writeFileSync(empty, await buildEmptyDeck(FIXTURE))
        const info = await readTemplateInfo(await DeckArchive.open(FIXTURE))
        expect(await run(empty, "fixture-empty", fixedScenario(info.layouts, media))).toMatchSnapshot()
    }, 120000)

    it.skipIf(!existsSync(DEMO))("matches on the demo deck (local golden)", async () => {
        matchLocal("demo", await run(DEMO, "demo", existingScenario(media)))
    }, 300000)

    const privates = existsSync(PRIVATE)
        ? readdirSync(PRIVATE).filter((f) => f.endsWith(".potx")).map((f) => path.join(PRIVATE, f))
        : []
    it.skipIf(privates.length === 0)("matches on the private templates (local golden)", async () => {
        for (const tpl of privates) {
            const name = path.basename(tpl, ".potx")
            const empty = path.join(TMP, `${name}-start.pptx`)
            mkdirSync(TMP, { recursive: true })
            writeFileSync(empty, await buildEmptyDeck(tpl))
            const info = await readTemplateInfo(await DeckArchive.open(tpl))
            matchLocal(name, await run(empty, name, [ ...fixedScenario(info.layouts, media), ...existingScenario(media) ]))
        }
    }, 600000)
})
