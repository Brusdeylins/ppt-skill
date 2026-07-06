/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: build a deck from the neutral fixture template via the real
**  write path (executeOps), re-read it, and verify content, structure, rev
**  semantics and the all-or-nothing failure guarantee.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck, buildSeed } from "../../src/engine/seed.js"
import { expectIntact } from "../util/integrity.js"
import { DeckArchive, readDeckState, readTemplateInfo } from "../../src/engine/reader.js"
import { executeOps } from "../../src/commands/apply.js"
import { PptcError } from "../../src/infra/errors.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "integration.pptx")

const opts = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
})

describe("roundtrip against the fixture template", () => {
    it("describes the template layouts generically", async () => {
        const info = await readTemplateInfo(await DeckArchive.open(TEMPLATE))
        expect(info.layouts.length).toBeGreaterThanOrEqual(3)
        const twoCol = info.layouts.find((l) => l.name === "TWO_COLUMN")
        expect(twoCol).toBeDefined()
        expect(twoCol?.placeholders.filter((p) => p.kind === "body")).toHaveLength(2)
    })

    it("builds slides with fills, elements, notes and refs", async () => {
        const result = await executeOps(DECK, {
            ops: [
                { op: "slide.add", ref: "first", layout: "CONTENT",
                    placeholders: { title: { text: "Erste Folie" }, body: { text: "Inhalt A\nInhalt B" } },
                    notes: "Notiztext" },
                { op: "slide.add", layout: "TWO_COLUMN",
                    placeholders: { title: { text: "Spalten" } } },
                { op: "el.add", slide: "title:Spalten", elements: [
                    { type: "table", frame: { x: 1, y: 2, w: 11 },
                        data: { headers: ["A", "B"], rows: [["1", "2"]] } }
                ] },
                { op: "slide.move", slide: "$first", to: 1 }
            ]
        }, opts)
        expect(result.result.applied).toBe(4)
        expect(result.result.slideCount).toBe(2)
        expect(result.result.slides["first"]).toMatchObject({ index: 1 })

        const state = await readDeckState(await DeckArchive.open(DECK))
        expect(state.slides).toHaveLength(2)
        expect(state.slides[0]?.title).toBe("Spalten")
        expect(state.slides[1]?.title).toBe("Erste Folie")
        expect(state.slides[1]?.notes).toBe("Notiztext")
        expect(state.slides[0]?.shapes.some((s) => s.type === "table")).toBe(true)
        expect(state.rev).toBe(result.rev.after)
    })

    it("enforces the optimistic lock", async () => {
        await expect(executeOps(DECK,
            { expectRev: "wrong-rev-0000", ops: [{ op: "slide.rm", slide: "index:0" }] }, opts))
            .rejects.toSatisfy((err: unknown) => (err as PptcError).code === "E_REV_CONFLICT")
    })

    it("leaves the deck untouched when an op fails", async () => {
        const before = readFileSync(DECK)
        await expect(executeOps(DECK, {
            ops: [
                { op: "slide.fill", slide: "index:0", placeholders: { title: { text: "geändert" } } },
                { op: "slide.rm", slide: "id:424242" }
            ]
        }, opts)).rejects.toSatisfy((err: unknown) =>
            (err as PptcError).code === "E_ADDR_NOTFOUND"
            && ((err as PptcError).details as { failedAt: number }).failedAt === 1)
        expect(readFileSync(DECK).equals(before)).toBe(true)
    })

    it("does not write on --dry-run", async () => {
        const before = readFileSync(DECK)
        const result = await executeOps(DECK,
            { ops: [{ op: "slide.rm", slide: "index:0" }] },
            { ...opts, dryRun: true })
        expect(result.result.dryRun).toBe(true)
        expect(readFileSync(DECK).equals(before)).toBe(true)
    })

    it("rejects schema violations with issue paths", async () => {
        await expect(executeOps(DECK, { ops: [{ op: "slide.fill" }] }, opts))
            .rejects.toSatisfy((err: unknown) => (err as PptcError).code === "E_SCHEMA")
    })

    it("keeps footer and slide-number placeholders on seed slides", async () => {
        const zip = await JSZip.loadAsync(readFileSync(TEMPLATE))
        const layoutPart = "ppt/slideLayouts/slideLayout1.xml"
        const layout = await (zip.file(layoutPart) as JSZip.JSZipObject).async("string")
        const sldNumSp = "<p:sp><p:nvSpPr><p:cNvPr id=\"90\" name=\"Foliennummer\"/><p:cNvSpPr/>"
            + "<p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"90\"/></p:nvPr></p:nvSpPr><p:spPr/>"
            + "<p:txBody><a:bodyPr/><a:lstStyle/><a:p>"
            + "<a:fld id=\"{11111111-2222-3333-4444-555555555555}\" type=\"slidenum\"><a:t>1</a:t></a:fld>"
            + "</a:p></p:txBody></p:sp>"
        zip.file(layoutPart, layout.replace("</p:spTree>", `${sldNumSp}</p:spTree>`))
        const seed = await buildSeed(await zip.generateAsync({ type: "nodebuffer" }))
        const out = await JSZip.loadAsync(seed.bytes)
        const slide = await (out.file("ppt/slides/slide1.xml") as JSZip.JSZipObject).async("string")
        expect(slide).toContain("type=\"sldNum\"")
        expect(slide).toContain("type=\"slidenum\"")
    })

    it("written deck passes the file-integrity validation", async () => {
        await expectIntact(DECK)
    })

    it("strips template sections (they only group the example slides)", async () => {
        const zip = await JSZip.loadAsync(readFileSync(TEMPLATE))
        const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
        const section = "<p:extLst><p:ext uri=\"{521415D9-36F7-43E2-AB2F-B90AF26B5E84}\">"
            + "<p14:sectionLst xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\">"
            + "<p14:section name=\"Beispiele\" id=\"{11111111-2222-3333-4444-555555555555}\">"
            + "<p14:sldIdLst><p14:sldId id=\"256\"/></p14:sldIdLst></p14:section>"
            + "</p14:sectionLst></p:ext></p:extLst>"
        zip.file("ppt/presentation.xml", pres.replace("</p:presentation>", `${section}</p:presentation>`))
        const seed = await buildSeed(await zip.generateAsync({ type: "nodebuffer" }))
        const out = await JSZip.loadAsync(seed.bytes)
        const outPres = await (out.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
        expect(outPres).not.toContain("sectionLst")
    })

    it("strips a master-view lastView so decks open in normal view", async () => {
        const zip = await JSZip.loadAsync(readFileSync(TEMPLATE))
        const viewPr = await (zip.file("ppt/viewProps.xml") as JSZip.JSZipObject).async("string")
        zip.file("ppt/viewProps.xml", viewPr
            .replace(/ lastView="[^"]*"/g, "")
            .replace("<p:viewPr ", "<p:viewPr lastView=\"sldMasterView\" "))
        const seed = await buildSeed(await zip.generateAsync({ type: "nodebuffer" }))
        const out = await JSZip.loadAsync(seed.bytes)
        const outViewPr = await (out.file("ppt/viewProps.xml") as JSZip.JSZipObject).async("string")
        expect(outViewPr).not.toContain("lastView=")
    })
})
