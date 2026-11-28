/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Golden: the slide features of release 1.2 on the neutral fixture, a
**  fixed op sequence pinned by its own snapshot (the 1.0.11-frozen engine
**  goldens stay untouched). Beyond the generic fingerprint, each slide
**  records what the features write: picture crops and frames, and the
**  language of every text run. Later phases of 1.2 extend the sequence.
*/

import { describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { PptcError } from "../../src/infra/errors.js"
import { elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { expectIntact } from "../util/integrity.js"
import { fingerprint } from "../util/fingerprint.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const FIXTURE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp", "golden-1-2")

const opts: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  a PNG header of the given size (the engine reads only the header)  */
const png = (w: number, h: number): Buffer => {
    const b = Buffer.alloc(33)
    Buffer.from([ 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a ]).copy(b, 0)
    b.writeUInt32BE(13, 8)
    b.write("IHDR", 12, "latin1")
    b.writeUInt32BE(w, 16)
    b.writeUInt32BE(h, 20)
    return b
}

/**  mask the UUID suffix of generated shape names  */
const maskName = (name: string): string =>
    name.replace(/-?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/, "-*")

/**  what the 1.2 features write on each slide: picture crop and frame,
     and the runs of every shape with their language  */
const featurePrint = async (deck: string): Promise<unknown[]> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const out: unknown[] = []
    for (const slide of state.slides) {
        const doc = parseXml(await (zip.file(slide.part) as JSZip.JSZipObject).async("string"))
        const pictures = elements(doc, "p:pic").map((pic) => {
            const src = firstElement(pic, "a:srcRect")
            const off = firstElement(pic, "a:off")
            const ext = firstElement(pic, "a:ext")
            return [ maskName(firstElement(pic, "p:cNvPr")?.getAttribute("name") ?? ""),
                src === null ? "no-crop" : [ "l", "t", "r", "b" ].map((a) => `${a}=${src.getAttribute(a) ?? ""}`).join(","),
                [ off?.getAttribute("x"), off?.getAttribute("y"), ext?.getAttribute("cx"), ext?.getAttribute("cy") ].join(",") ].join("|")
        })
        const runs = [ ...elements(doc, "p:sp"), ...elements(doc, "p:graphicFrame") ].flatMap((shape) => {
            const name = maskName(firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? "")
            return elements(shape, "a:r").map((r) =>
                `${name}: ${firstElement(r, "a:t")?.textContent ?? ""}=${firstElement(r, "a:rPr")?.getAttribute("lang") ?? "-"}`)
        })
        out.push({ pictures, runs })
    }
    return out
}

/**  one step: ops and whether the apply must fail (with its code)  */
interface Step { ops: unknown[], fails?: string }

describe("golden 1.2 features", () => {
    it("matches on an empty deck from the fixture", async () => {
        mkdirSync(TMP, { recursive: true })
        const landscape = path.join(TMP, "landscape.png")
        const portrait = path.join(TMP, "portrait.png")
        writeFileSync(landscape, png(1600, 900))
        writeFileSync(portrait, png(900, 1600))
        const deck = path.join(TMP, "features.pptx")
        writeFileSync(deck, await buildEmptyDeck(FIXTURE))

        const steps: Step[] = [
            /*  A2: text language from pptcDeckLang, with overrides  */
            { ops: [
                { op: "meta.props", set: { custom: { pptcDeckLang: "de-DE" } } },
                { op: "slide.add", layout: "CONTENT", ref: "text", notes: "Notiz", footer: "Fußzeile",
                    placeholders: { title: { text: "Sprache" },
                        body: { text: [ { text: "deutsch" }, { text: "english", lang: "en-GB" },
                            { runs: [ { text: "uno" }, { text: "deux", lang: "fr-FR" } ], lang: "it-IT" } ] } } },
                { op: "el.add", slide: "$text", elements: [
                    { type: "textbox", name: "Box", frame: { x: 0.5, y: 5.5, w: 3, h: 0.6 }, text: "Kasten" },
                    { type: "shape", name: "Form", shape: "rect", frame: { x: 4, y: 5.5, w: 2, h: 0.6 }, text: "Form", lang: "nl-NL" },
                    { type: "table", name: "Tab", frame: { x: 6.5, y: 5.5, w: 4 }, data: { headers: [ "Kopf" ], rows: [ [ "Zelle" ] ] } }
                ] }
            ] },
            /*  B3: picture placeholder fits  */
            { ops: [
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: landscape } } },
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: portrait, fit: "crop" } } },
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: portrait, fit: "contain" } } },
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: landscape, fit: "stretch" } } }
            ] },
            /*  B1: el.set text on a table fails, on a textbox it works  */
            { ops: [ { op: "el.set", slide: "0", name: "Tab", text: "x" } ], fails: "E_SCHEMA" },
            { ops: [ { op: "el.set", slide: "0", name: "Box", text: "Neu" } ] }
        ]

        const prints: unknown[] = []
        for (const step of steps) {
            let error: string | null = null
            try {
                await executeOps(deck, { ops: step.ops } as never, opts)
            }
            catch (err) {
                error = err instanceof PptcError ? err.code : String(err)
            }
            expect(error).toBe(step.fails ?? null)
            await expectIntact(deck)
            prints.push({ ops: JSON.parse(JSON.stringify(step.ops).split(TMP).join("<tmp>")) as unknown, error,
                slides: await fingerprint(deck), features: await featurePrint(deck) })
        }
        expect(prints).toMatchSnapshot()
    }, 120000)
})
