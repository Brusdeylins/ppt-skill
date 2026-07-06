/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: element-level ops and the write-path features that the
**  basic roundtrip does not cover -- el.set/el.rm (incl. same-run
**  cancellation and UUID-suffix matching), prompt boxes, backgrounds,
**  hyperlinks, placeholder images, slide.copy isolation and --out.
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
import { PptcError } from "../../src/infra/errors.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "elements.pptx")
const PNG = path.join(TMP, "pixel.png")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null }

/**  smallest valid 1x1 PNG  */
const PIXEL = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64")

/**  read a slide part (by presentation order) from the deck zip  */
const slideXml = async (file: string, index: number): Promise<string> => {
    const zip = await JSZip.loadAsync(readFileSync(file))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const rels = await (zip.file("ppt/_rels/presentation.xml.rels") as JSZip.JSZipObject).async("string")
    const rids = [...pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)].map((m) => m[1])
    const relMap = new Map([...rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]))
    return await (zip.file(`ppt/${relMap.get(rids[index] as string) as string}`) as JSZip.JSZipObject).async("string")
}

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(PNG, PIXEL)
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    await executeOps(DECK, {
        ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Basis" } } }
        ]
    }, opts)
})

describe("element ops", () => {
    it("cancels a same-run generated element via el.rm", async () => {
        await executeOps(DECK, {
            ops: [
                { op: "el.add", slide: "title:Basis", elements: [
                    { type: "textbox", frame: { x: 1, y: 5, w: 4, h: 0.5 }, text: "wegwerf", name: "Throwaway" }
                ] },
                { op: "el.rm", slide: "title:Basis", name: "Throwaway" }
            ]
        }, opts)
        expect(await slideXml(DECK, 0)).not.toContain("Throwaway")
    })

    it("removes existing elements by base name despite UUID suffixes", async () => {
        await executeOps(DECK, {
            ops: [{ op: "el.add", slide: "title:Basis", elements: [
                { type: "textbox", frame: { x: 1, y: 5, w: 4, h: 0.5 }, text: "bleibt erstmal", name: "Kasten" }
            ] }]
        }, opts)
        expect(await slideXml(DECK, 0)).toContain("Kasten")
        await executeOps(DECK, { ops: [{ op: "el.rm", slide: "title:Basis", name: "Kasten" }] }, opts)
        expect(await slideXml(DECK, 0)).not.toContain("Kasten")
    })

    it("fails cleanly when el.rm addresses nothing", async () => {
        await expect(executeOps(DECK,
            { ops: [{ op: "el.rm", slide: "title:Basis", name: "GibtsNicht" }] }, opts))
            .rejects.toSatisfy((err: unknown) => (err as PptcError).code === "E_ADDR_NOTFOUND")
    })

    it("patches a same-run generated textbox via el.set", async () => {
        await executeOps(DECK, {
            ops: [
                { op: "el.add", slide: "title:Basis", elements: [
                    { type: "textbox", frame: { x: 1, y: 5.6, w: 4, h: 0.5 }, text: "alt", name: "SetBox" }
                ] },
                { op: "el.set", slide: "title:Basis", name: "SetBox", text: "neu gesetzt" }
            ]
        }, opts)
        const xml = await slideXml(DECK, 0)
        expect(xml).toContain("neu gesetzt")
        expect(xml).not.toContain(">alt<")
    })

    it("adds and removes prompt boxes on picture-less layouts with a clean error", async () => {
        await expect(executeOps(DECK,
            { ops: [{ op: "img.prompts", slide: "title:Basis", prompts: "x" }] }, opts))
            .rejects.toSatisfy((err: unknown) => (err as PptcError).code === "E_ADDR_NOTFOUND")
    })
})

describe("write-path features", () => {
    it("sets a solid background and wires hyperlinks externally", async () => {
        await executeOps(DECK, {
            ops: [
                { op: "slide.fill", slide: "title:Basis", background: { color: "#1A1A2E" } },
                { op: "el.add", slide: "title:Basis", elements: [
                    { type: "textbox", frame: { x: 1, y: 6.2, w: 5, h: 0.4 }, name: "Link",
                        text: [{ runs: [{ text: "mehr", hyperlink: "https://example.org/mehr" }] }] }
                ] }
            ]
        }, opts)
        const xml = await slideXml(DECK, 0)
        expect(xml).toContain("1A1A2E")
        expect(xml).not.toContain("pptc-hlink")
        const zip = await JSZip.loadAsync(readFileSync(DECK))
        const relsParts = Object.keys(zip.files).filter((f) => /slides\/_rels\/.*\.rels$/.test(f))
        let external = false
        for (const part of relsParts) {
            const rels = await (zip.file(part) as JSZip.JSZipObject).async("string")
            if (rels.includes("example.org/mehr") && rels.includes("TargetMode=\"External\""))
                external = true
        }
        expect(external).toBe(true)
    })

    it("fills picture placeholders with image media", async () => {
        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "TITLE_SLIDE",
                placeholders: { title: { text: "Bildfolie" } } }]
        }, opts)
        /*  the fixture has no picture placeholder; expect the clean error  */
        await expect(executeOps(DECK, {
            ops: [{ op: "slide.fill", slide: "title:Bildfolie",
                placeholders: { image: { image: PNG } } }]
        }, opts)).rejects.toSatisfy((err: unknown) => (err as PptcError).code === "E_ADDR_NOTFOUND")
    })

    it("keeps copies isolated: el.set on the copy must not patch the original", async () => {
        await executeOps(DECK, {
            ops: [
                { op: "slide.add", ref: "orig", layout: "CONTENT",
                    placeholders: { title: { text: "Original" } } },
                { op: "el.add", slide: "$orig", elements: [
                    { type: "textbox", frame: { x: 1, y: 5, w: 4, h: 0.5 }, text: "geteilt?", name: "Iso" }
                ] },
                { op: "slide.copy", slide: "$orig", ref: "klon" },
                { op: "el.set", slide: "$klon", name: "Iso", text: "nur im Klon" }
            ]
        }, opts)
        const state = await readDeckState(await DeckArchive.open(DECK))
        const orig = state.slides.find((s) => s.title === "Original" && state.slides.indexOf(s) ===
            state.slides.findIndex((x) => x.title === "Original")) ?? state.slides[0]
        const klon = state.slides.filter((s) => s.title === "Original")[1]
        const textOf = (slide: typeof orig, name: string): string | null =>
            slide?.shapes.find((sh) => sh.name.startsWith(name))?.text ?? null
        expect(textOf(orig, "Iso")).toBe("geteilt?")
        expect(textOf(klon, "Iso")).toBe("nur im Klon")
    })

    it("writes to --out and leaves the source untouched", async () => {
        const before = readFileSync(DECK)
        const variant = path.join(TMP, "variant-out.pptx")
        await executeOps(DECK,
            { ops: [{ op: "meta.props", set: { subject: "Variante" } }] },
            { ...opts, outFile: variant })
        expect(readFileSync(DECK).equals(before)).toBe(true)
        const zip = await JSZip.loadAsync(readFileSync(variant))
        const core = await (zip.file("docProps/core.xml") as JSZip.JSZipObject).async("string")
        expect(core).toContain("Variante")
    })

    it("warns when a new element covers a text placeholder", async () => {
        const result = await executeOps(DECK, {
            ops: [
                { op: "slide.add", layout: "CONTENT",
                    placeholders: { title: { text: "Overlap-Probe" }, body: { text: "Inhalt" } } },
                { op: "el.add", slide: "title:Overlap-Probe", elements: [
                    { type: "textbox", name: "Ueberlappt",
                        frame: { x: 1, y: 2, w: 8, h: 3 }, text: "sitzt auf dem Body" }
                ] }
            ]
        }, opts)
        const overlap = result.warnings.filter((w) => w.code === "W_ELEMENT_OVERLAP")
        expect(overlap.length).toBeGreaterThan(0)
        expect(overlap[0]?.element).toBe("Ueberlappt")
    })

    it("escalates the overlap warning under --strict", async () => {
        await expect(executeOps(DECK, {
            ops: [{ op: "el.add", slide: "title:Overlap-Probe", elements: [
                { type: "textbox", name: "NochEiner", frame: { x: 1, y: 2, w: 8, h: 3 }, text: "x" }
            ] }]
        }, { ...opts, strict: true })).rejects.toSatisfy((err: unknown) =>
            (err as PptcError).code === "E_LINT")
    })

    it("does not warn for prompt boxes or picture-area placement", async () => {
        const result = await executeOps(DECK, {
            ops: [
                { op: "slide.add", layout: "PICTURE",
                    placeholders: { title: { text: "Bildfrei" } } },
                /*  prompt box overlays the picture placeholder by design  */
                { op: "img.prompts", slide: "title:Bildfrei", prompts: "p" },
                /*  a free element over the PICTURE area is fine too  */
                { op: "el.add", slide: "title:Bildfrei", elements: [
                    { type: "textbox", name: "AiNote",
                        frame: { x: 1, y: 2.2, w: 2, h: 0.3 }, text: "AI-generated" }
                ] }
            ]
        }, opts)
        expect(result.warnings.filter((w) => w.code === "W_ELEMENT_OVERLAP")).toHaveLength(0)
    })

    it("replacing an element (el.rm + el.add) does not warn against itself", async () => {
        await executeOps(DECK, {
            ops: [
                { op: "slide.add", layout: "DEFAULT", ref: "leer", at: 0 },
                { op: "el.add", slide: "$leer", elements: [
                    { type: "textbox", name: "Ersatz", frame: { x: 1, y: 1, w: 4, h: 2 }, text: "alt" }
                ] }
            ]
        }, opts)
        const result = await executeOps(DECK, {
            ops: [
                { op: "el.rm", slide: "index:0", name: "Ersatz" },
                { op: "el.add", slide: "index:0", elements: [
                    { type: "textbox", name: "Ersatz", frame: { x: 1, y: 1, w: 4, h: 2 }, text: "neu" }
                ] }
            ]
        }, { ...opts, strict: true })
        expect(result.warnings.filter((w) => w.code === "W_ELEMENT_OVERLAP")).toHaveLength(0)
    })

    it("written decks pass the file-integrity validation", async () => {
        await expectIntact(DECK)
        await expectIntact(path.join(TMP, "variant-out.pptx"))
    })

    it("does not accumulate dangling slide relationships across many applies", async () => {
        const deck = path.join(TMP, "churn.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await executeOps(deck, { ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "A" } } },
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "B" } } }
        ] }, opts)
        /*  automizer re-imports every slide on each apply and mints a fresh
            "-created" rel; without the prune the rels file would grow by one
            per slide per apply until PowerPoint demands a repair  */
        for (let i = 0; i < 6; i++)
            await executeOps(deck, { ops: [
                { op: "meta.props", set: { title: `pass ${i}` } }
            ] }, opts)
        /*  grow the deck on a later apply: docProps/app.xml must follow,
            or PowerPoint demands a repair over the stale slide count  */
        await executeOps(deck, { ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "C" } } }
        ] }, opts)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
        const rels = await (zip.file("ppt/_rels/presentation.xml.rels") as JSZip.JSZipObject).async("string")
        const app = await (zip.file("docProps/app.xml") as JSZip.JSZipObject).async("string")
        const slideCount = [...pres.matchAll(/<p:sldId /g)].length
        const slideRels = [...rels.matchAll(/Type="[^"]*\/slide"/g)].length
        expect(slideRels).toBe(slideCount)
        expect(/<Slides>(\d+)<\/Slides>/.exec(app)?.[1]).toBe(String(slideCount))
        await expectIntact(deck)
    })

    it("repairs two slides that share one notesSlide", async () => {
        const deck = path.join(TMP, "sharednotes.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await executeOps(deck, { ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "N1" } }, notes: "notes one" },
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "N2" } }, notes: "notes two" }
        ] }, opts)
        /*  inject the aliasing automizer's part renumbering can leave behind:
            point the second slide's notes rel at the first slide's notes part  */
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const relsParts = Object.keys(zip.files)
            .filter((f) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(f))
        const targets: { part: string, text: string, notes: string }[] = []
        for (const part of relsParts) {
            const txt = await (zip.file(part) as JSZip.JSZipObject).async("string")
            const m = /Target="(\.\.\/notesSlides\/[^"]+)"/.exec(txt)
            if (m !== null)
                targets.push({ part, text: txt, notes: m[1] as string })
        }
        expect(targets.length).toBeGreaterThanOrEqual(2)
        const a = targets[0] as { part: string, text: string, notes: string }
        const b = targets[1] as { part: string, text: string, notes: string }
        zip.file(b.part, b.text.replace(b.notes, a.notes))   // b now shares a's notes
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
        /*  any apply must re-establish a 1:1 slide <-> notesSlide mapping  */
        await executeOps(deck, { ops: [{ op: "meta.props", set: { title: "x" } }] }, opts)
        const z2 = await JSZip.loadAsync(readFileSync(deck))
        const refs = new Map<string, number>()
        for (const part of Object.keys(z2.files).filter((f) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(f))) {
            const txt = await (z2.file(part) as JSZip.JSZipObject).async("string")
            const m = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(txt)
            if (m !== null)
                refs.set(m[1] as string, (refs.get(m[1] as string) ?? 0) + 1)
        }
        expect(Math.max(...refs.values())).toBe(1)
        await expectIntact(deck)
    })

    it("state exposes table geometry/cells and autoshape styling for round-trip edits", async () => {
        const deck = path.join(TMP, "introspect.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await executeOps(deck, { ops: [
            { op: "slide.add", layout: "CONTENT", ref: "s", placeholders: { title: { text: "Introspect" } } },
            { op: "el.add", slide: "$s", elements: [
                { type: "table", name: "T", frame: { x: 1, y: 1.5, w: 8, h: 1.2 },
                    data: { headers: ["Tag", "What"], rows: [["<flow>", "a flow"], ["<gate/>", "a gate"]] } },
                { type: "shape", name: "Box", shape: "roundRect", text: "Hi",
                    frame: { x: 1, y: 3.5, w: 2.5, h: 1 },
                    fill: "A01441", border: "139EAD", borderPt: 2, fontColor: "FFFFFF", fontSize: 14 }
            ] }
        ] }, opts)

        const deckState = await readDeckState(await DeckArchive.open(deck))
        const shapes = deckState.slides[0]?.shapes ?? []
        const table = shapes.find((s) => s.type === "table")
        expect(table?.frame).not.toBeNull()
        expect(table?.frame?.w).toBeCloseTo(8, 1)
        expect(table?.table).toEqual([["Tag", "What"], ["<flow>", "a flow"], ["<gate/>", "a gate"]])
        expect(table?.colWidths?.length).toBe(2)

        const box = shapes.find((s) => s.name.startsWith("Box"))
        expect(box?.shape).toBe("roundRect")
        expect(box?.fill).toBe("A01441")
        expect(box?.border).toBe("139EAD")
        expect(box?.borderPt).toBeCloseTo(2, 1)
        expect(box?.fontColor).toBe("FFFFFF")
        expect(box?.fontSize).toBeCloseTo(14, 1)
        await expectIntact(deck)
    })
})

describe("font-size lint (W_FONT_TOO_SMALL)", () => {
    const deck = path.join(TMP, "fonts.pptx")
    const fontWarnings = (r: Awaited<ReturnType<typeof executeOps>>): typeof r.warnings =>
        r.warnings.filter((w) => w.code === "W_FONT_TOO_SMALL")

    beforeAll(async () => {
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        /*  slide 0: a CONTENT filler; slide 1: a blank DEFAULT surface (no text
            placeholder) so el.add elements never trip the overlap lint -- keeping
            these font-size assertions isolated from W_ELEMENT_OVERLAP  */
        await executeOps(deck, { ops: [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Fonts" } } },
            { op: "slide.add", layout: "DEFAULT" }
        ] }, opts)
    })

    it("flags a free textbox run below the 11pt floor", async () => {
        const r = await executeOps(deck, { ops: [
            { op: "el.add", slide: "index:1", elements: [
                { type: "textbox", name: "Tiny", frame: { x: 0.3, y: 3, w: 3, h: 0.3 },
                    text: [{ runs: [{ text: "winzig", size: 8 }] }] }
            ] }
        ] }, opts)
        const f = fontWarnings(r)
        expect(f).toHaveLength(1)
        expect(f[0]?.element).toBe("Tiny")
        expect(f[0]?.fontPt).toBe(8)
        expect(f[0]?.minPt).toBe(11)
    })

    it("does not flag 12pt or exactly 11pt", async () => {
        const r = await executeOps(deck, { ops: [
            { op: "el.add", slide: "index:1", elements: [
                { type: "textbox", name: "Ok12", frame: { x: 0.3, y: 3.5, w: 3, h: 0.3 },
                    text: [{ runs: [{ text: "ok", size: 12 }] }] },
                { type: "textbox", name: "Ok11", frame: { x: 4, y: 3.5, w: 3, h: 0.3 },
                    text: [{ runs: [{ text: "grenz", size: 11 }] }] }
            ] }
        ] }, opts)
        expect(fontWarnings(r)).toHaveLength(0)
    })

    it("flags small shape, table and chart fonts", async () => {
        const r = await executeOps(deck, { ops: [
            { op: "el.add", slide: "index:1", elements: [
                { type: "shape", name: "S", shape: "rect", text: "x", fontSize: 9,
                    frame: { x: 0.3, y: 1, w: 1.5, h: 0.6 } },
                { type: "table", name: "T", frame: { x: 2, y: 1, w: 3, h: 1 },
                    data: { rows: [["a", "b"]], style: { fontSize: 8 } } },
                { type: "chart", name: "C", frame: { x: 5.3, y: 1, w: 3, h: 2 },
                    data: { type: "bar", categories: ["a", "b"],
                        series: [{ name: "s", values: [1, 2] }], fontSize: 8 } }
            ] }
        ] }, opts)
        expect(fontWarnings(r).map((w) => w.element).sort()).toEqual(["C", "S", "T"])
    })

    it("respects --min-font-pt: 0 disables, a higher floor tightens", async () => {
        const off = await executeOps(deck, { ops: [
            { op: "el.add", slide: "index:1", elements: [
                { type: "shape", name: "S0", shape: "rect", text: "x", fontSize: 8,
                    frame: { x: 0.3, y: 4, w: 1.5, h: 0.6 } }
            ] }
        ] }, { ...opts, minFontPt: 0 })
        expect(fontWarnings(off)).toHaveLength(0)

        const tight = await executeOps(deck, { ops: [
            { op: "el.add", slide: "index:1", elements: [
                { type: "shape", name: "S14", shape: "rect", text: "x", fontSize: 12,
                    frame: { x: 0.3, y: 4.8, w: 1.5, h: 0.6 } }
            ] }
        ] }, { ...opts, minFontPt: 14 })
        expect(fontWarnings(tight)).toHaveLength(1)
        expect(fontWarnings(tight)[0]?.minPt).toBe(14)
    })

    it("escalates a small font under --strict (exit 7 / E_LINT)", async () => {
        await expect(executeOps(deck, { ops: [
            { op: "el.add", slide: "index:1", elements: [
                { type: "shape", name: "SS", shape: "rect", text: "x", fontSize: 8,
                    frame: { x: 0.3, y: 5.5, w: 1.5, h: 0.6 } }
            ] }
        ] }, { ...opts, strict: true })).rejects.toSatisfy((err: unknown) =>
            (err as PptcError).code === "E_LINT")
    })

    it("flags a placeholder run that overrides the template size too small", async () => {
        const r = await executeOps(deck, { ops: [
            { op: "slide.add", layout: "CONTENT",
                placeholders: { title: { text: "PhFont" },
                    body: { text: [{ runs: [{ text: "mini", size: 7 }] }] } } }
        ] }, opts)
        const f = fontWarnings(r)
        expect(f.some((w) => typeof w.placeholder === "number" && w.fontPt === 7)).toBe(true)
    })

    it("does not flag prompt-box captions on a picture layout", async () => {
        const r = await executeOps(deck, { ops: [
            { op: "slide.add", layout: "PICTURE", placeholders: { title: { text: "Bild" } } },
            { op: "img.prompts", slide: "title:Bild", prompts: "a prompt" }
        ] }, opts)
        expect(fontWarnings(r)).toHaveLength(0)
    })

    it("written deck passes the file-integrity validation", async () => {
        await expectIntact(deck)
    })
})
