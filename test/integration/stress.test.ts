/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: stress the write path the way real agent sessions do --
**  many applies in a row on the same deck, every op kind, every element
**  type, plus seeded-random edit combinations -- holding the
**  file-integrity bar after EVERY single apply (the repair-trigger bugs
**  of 0.2.x all came from repeated applies).
*/

import { beforeAll, describe, expect, it } from "vitest"
import JSZip from "jszip"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import type { Op, OpsDocument } from "../../src/schema/ops.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "stress.pptx")
const FUZZ = path.join(TMP, "fuzz.pptx")
const PNG = path.join(TMP, "stress-pixel.png")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    /*  smallest valid 1x1 PNG  */
    writeFileSync(PNG, Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
        "base64"))
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    writeFileSync(FUZZ, await buildEmptyDeck(TEMPLATE))
})

/**  apply one ops document to a deck and assert the file stays intact  */
const applyIntact = async (file: string, doc: OpsDocument): Promise<void> => {
    await executeOps(file, doc, opts)
    await expectIntact(file)
}

describe("multi-apply stress with per-apply integrity", () => {
    it("survives a long agent-like session of sequential applies", async () => {
        /*  round 1: deck skeleton with notes, footer and special chars  */
        await applyIntact(DECK, { ops: [
            { op: "slide.add", ref: "intro", layout: "TITLE_SLIDE",
                placeholders: { title: { text: "Stressdeck" }, subtitle: { text: "Untertitel" } },
                notes: "Begrüßung mit Umlauten: äöü & <spitzen> Klammern." },
            { op: "slide.add", ref: "inhalt", layout: "CONTENT",
                placeholders: { title: { text: "Inhalt" }, body: { text: "Punkt A\nPunkt B" } },
                footer: "Fußzeile & \"Anführung\" <eins>" },
            { op: "slide.add", ref: "frei", layout: "DEFAULT" },
            { op: "slide.add", ref: "bild", layout: "PICTURE",
                placeholders: { title: { text: "Bildfolie" } } }
        ] })

        /*  rounds address slides like a real agent: read state, use ids
            ($refs are document-local by design)  */
        const idOf = async (index: number): Promise<string> => {
            const s = await readDeckState(await DeckArchive.open(DECK))
            return `id:${s.slides[index]?.id}`
        }

        /*  round 2: every element type on the blank slide (index 2)  */
        const frei = await idOf(2)
        await applyIntact(DECK, { ops: [
            { op: "el.add", slide: frei, elements: [
                { type: "textbox", name: "Box", frame: { x: 0.5, y: 0.5, w: 4, h: 1 },
                    text: "Kasten & <Sonderzeichen>" },
                { type: "table", name: "Tab", frame: { x: 0.5, y: 1.8, w: 6 },
                    data: { headers: ["A", "B"], rows: [["1", "2"], ["3", "4"]] } },
                { type: "chart", name: "Diagramm", frame: { x: 7, y: 0.5, w: 5, h: 3 },
                    data: { type: "column", categories: ["Q1", "Q2"],
                        series: [{ name: "Umsatz", values: [10, 14] }] } },
                { type: "shape", name: "Pfeil", shape: "rightArrow",
                    frame: { x: 7, y: 4, w: 2, h: 1 }, fill: "1F4E79" },
                { type: "image", name: "Logo", frame: { x: 10, y: 4, w: 1, h: 1 }, path: PNG },
                { type: "connector", name: "Linie", from: [0.5, 5.5], to: [5.5, 6] }
            ] }
        ] })

        /*  round 3: picture placeholder fill + image prompts  */
        await applyIntact(DECK, { ops: [
            { op: "slide.fill", slide: "title:Bildfolie",
                placeholders: { "image": { image: PNG } } },
            { op: "slide.add", ref: "bild2", layout: "PICTURE",
                placeholders: { title: { text: "Promptfolie" } } },
            { op: "img.prompts", slide: "$bild2",
                prompts: "deep blue (#1F4E79) accent, 4:3 composition, sharp focus" }
        ] })

        /*  round 4: edit churn -- append, retext, remove, copy, move, rm  */
        await applyIntact(DECK, { ops: [
            { op: "slide.fill", slide: "title:Inhalt",
                placeholders: { body: { text: "Punkt C", append: true } } },
            { op: "el.set", slide: "index:2", name: "Box", text: "umgetextet & <ok>" },
            { op: "el.rm", slide: "index:2", name: "Linie" },
            { op: "slide.copy", slide: "title:Inhalt", ref: "kopie" },
            { op: "slide.move", slide: "$kopie", to: 0 },
            { op: "slide.rm", slide: "title:Bildfolie" },
            { op: "meta.props", set: { title: "Stress & Test", author: "pptc" } }
        ] })

        /*  rounds 5-8: repeated re-applies must not accumulate debris  */
        for (let round = 0; round < 4; round++)
            await applyIntact(DECK, { ops: [
                { op: "slide.fill", slide: "index:0",
                    placeholders: { title: { text: `Runde ${round}` } },
                    notes: `Notiz der Runde ${round}`,
                    footer: `Fußzeile ${round}` },
                { op: "slide.add", ref: `extra${round}`, layout: "CONTENT",
                    placeholders: { title: { text: `Extra ${round}` } } },
                { op: "slide.rm", slide: `$extra${round}` }
            ] })

        const state = await readDeckState(await DeckArchive.open(DECK))
        expect(state.slides.length).toBeGreaterThanOrEqual(4)
        expect(state.slides[0]?.title).toBe("Runde 3")
        expect(state.slides[0]?.notes).toBe("Notiz der Runde 3")

        /*  re-applies must not multiply chart/embedding parts (one chart
            was added once -- exactly one part of each may survive)  */
        const zip = await JSZip.loadAsync(readFileSync(DECK))
        const charts = Object.keys(zip.files).filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
        const sheets = Object.keys(zip.files).filter((f) =>
            f.startsWith("ppt/embeddings/") && !f.endsWith("/"))
        expect(charts).toHaveLength(1)
        expect(sheets.length).toBeLessThanOrEqual(1)
    }, 120000)

    it("re-wiring hyperlinks across applies never duplicates rel ids", async () => {
        const link = (label: string): { runs: { text: string, hyperlink: string }[] }[] =>
            [{ runs: [{ text: label, hyperlink: "https://example.com/docs" }] }]
        /*  two applies that each (re)write a hyperlink into the same
            placeholder: the second re-wires and must not collide with
            the rIdPptc left by the first (the probe.pptx repair bug)  */
        await applyIntact(DECK, { ops: [
            { op: "slide.fill", slide: "title:Inhalt", placeholders: { body: { text: link("erster Link") } } }
        ] })
        await applyIntact(DECK, { ops: [
            { op: "slide.fill", slide: "title:Inhalt", placeholders: { body: { text: link("zweiter Link") } } }
        ] })
    })

    it("escapes XML special characters end to end", async () => {
        const state = await readDeckState(await DeckArchive.open(DECK))
        const intro = state.slides.find((s) => s.title === "Stressdeck")
        expect(intro?.notes).toContain("äöü & <spitzen>")
        const box = state.slides.flatMap((s) => s.shapes).find((sh) => sh.name.startsWith("Box"))
        expect(box?.text).toContain("umgetextet & <ok>")
    })
})

describe("seeded-random edit combinations", () => {
    /*  deterministic LCG so failures are reproducible by seed  */
    const lcg = (seed: number): (() => number) => {
        let s = seed
        return (): number => {
            s = (s * 1664525 + 1013904223) % 4294967296
            return s / 4294967296
        }
    }

    it("random op combinations never leave a broken file (seed 42)", async () => {
        const rnd = lcg(42)
        const pick = <T>(arr: T[]): T => arr[Math.floor(rnd() * arr.length)] as T
        let serial = 0

        for (let round = 0; round < 12; round++) {
            const state = await readDeckState(await DeckArchive.open(FUZZ))
            const slides = state.slides
            const ops: Op[] = []
            const opCount = 1 + Math.floor(rnd() * 4)
            for (let i = 0; i < opCount; i++) {
                const id = (): string => `id:${(pick(slides)).id}`
                const kind = slides.length === 0
                    ? "add"
                    : pick(["add", "fill", "copy", "move", "rm", "el", "note", "footer"])
                switch (kind) {
                    case "add":
                        ops.push({ op: "slide.add", layout: pick(["CONTENT", "TWO_COLUMN", "PICTURE", "DEFAULT"]),
                            placeholders: { title: { text: `Folie ${serial++}` } } })
                        break
                    case "fill":
                        ops.push({ op: "slide.fill", slide: id(),
                            placeholders: { title: { text: `Neu ${serial++} & <x>` } } })
                        break
                    case "copy":
                        ops.push({ op: "slide.copy", slide: id() })
                        break
                    case "move":
                        ops.push({ op: "slide.move", slide: id(),
                            to: Math.floor(rnd() * slides.length) })
                        break
                    case "rm":
                        if (slides.length > 2)
                            ops.push({ op: "slide.rm", slide: id() })
                        break
                    case "el":
                        ops.push({ op: "el.add", slide: id(), elements: [
                            { type: "textbox", name: `Fz${serial++}`,
                                frame: { x: rnd() * 8, y: rnd() * 5, w: 2, h: 0.5 },
                                text: `Fuzz & <${serial}>` }
                        ] })
                        break
                    case "note":
                        ops.push({ op: "slide.fill", slide: id(), notes: `Notiz ${serial++}` })
                        break
                    case "footer":
                        ops.push({ op: "slide.fill", slide: id(), footer: `Fuß ${serial++}` })
                        break
                }
            }
            if (ops.length === 0)
                continue
            /*  ops within one document may invalidate each other's targets
                (rm before fill of the same id): planner rejection is fine,
                a written-but-broken file is not  */
            try {
                await executeOps(FUZZ, { ops }, opts)
            }
            catch {
                /*  planner refused: deck must be untouched and intact  */
            }
            await expectIntact(FUZZ)
        }

        const final = await readDeckState(await DeckArchive.open(FUZZ))
        expect(final.slides.length).toBeGreaterThan(0)
    }, 120000)
})
