/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the language of the text pptc writes. Every run of a
**  placeholder fill, el.set, footer, notes and el.add textbox, shape and
**  table carries `lang`: the deck language (pptcDeckLang, else the deck's
**  own default text language) unless the element, paragraph or run states
**  one. `state --level full` reads it back per shape.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import type { Element } from "@xmldom/xmldom"
import { elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  run texts with their language, per shape name prefix, of a slide  */
const runLangs = async (deck: string, index: number): Promise<Map<string, string[]>> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const part = (state.slides[index] as { part: string }).part
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const doc = parseXml(await (zip.file(part) as JSZip.JSZipObject).async("string"))
    const out = new Map<string, string[]>()
    for (const shape of [ ...elements(doc, "p:sp"), ...elements(doc, "p:graphicFrame") ]) {
        const name = (firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? "").replace(/-[0-9a-f-]{36}$/, "")
        out.set(name, elements(shape, "a:r").map((r: Element) =>
            `${firstElement(r, "a:t")?.textContent ?? ""}=${firstElement(r, "a:rPr")?.getAttribute("lang") ?? "-"}`))
    }
    return out
}

/**  the language of every run of a slide's notes  */
const notesLangs = async (deck: string, index: number): Promise<string[]> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const part = (state.slides[index] as { part: string }).part
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const rels = await (zip.file(`ppt/slides/_rels/${path.posix.basename(part)}.rels`) as JSZip.JSZipObject).async("string")
    const target = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(rels)?.[1] as string
    const notes = parseXml(await (zip.file(`ppt/notesSlides/${target}`) as JSZip.JSZipObject).async("string"))
    return elements(notes, "a:rPr").map((r) => r.getAttribute("lang") ?? "-")
}

describe("text language", () => {
    const DE = path.join(TMP, "lang-de.pptx")

    beforeAll(async () => {
        mkdirSync(TMP, { recursive: true })
        writeFileSync(DE, await buildEmptyDeck(TEMPLATE))
        await executeOps(DE, { ops: [
            { op: "meta.props", set: { custom: { pptcDeckLang: "de-DE" } } },
            { op: "slide.add", layout: "CONTENT", ref: "s", notes: "Notiz", footer: "Fußzeile",
                placeholders: { title: { text: "Übersicht" },
                    body: { text: [ { text: "Punkt" }, { text: "point", lang: "en-GB" },
                        { runs: [ { text: "eins" }, { text: "deux", lang: "fr-FR" } ], lang: "it-IT" } ] } } },
            { op: "el.add", slide: "$s", elements: [
                { type: "textbox", name: "Box", frame: { x: 0.5, y: 5.5, w: 3, h: 0.6 }, text: "Kasten" },
                { type: "textbox", name: "Mixed", frame: { x: 4, y: 5.5, w: 3, h: 0.6 }, lang: "nl-NL",
                    text: [ { runs: [ { text: "doos" }, { text: "box", lang: "en-US" } ] } ] },
                { type: "shape", name: "Form", shape: "rect", frame: { x: 8, y: 5.5, w: 2, h: 0.6 }, text: "Form" },
                { type: "table", name: "Tab", frame: { x: 0.5, y: 6.3, w: 4 }, data: { headers: [ "Kopf" ], rows: [ [ "Zelle" ] ] } },
                { type: "table", name: "TabFr", frame: { x: 5, y: 6.3, w: 4 }, lang: "fr-FR", data: { rows: [ [ "cellule" ] ] } }
            ] }
        ] } as never, opts)
        await expectIntact(DE)
    })

    it("writes the pptcDeckLang into placeholder runs, footer and notes", async () => {
        const langs = await runLangs(DE, 0)
        expect(langs.get("PptcPh-0")).toEqual([ "Übersicht=de-DE" ])
        expect(langs.get("PptcPh-1")).toEqual([ "Punkt=de-DE", "point=en-GB", "eins=it-IT", "deux=fr-FR" ])
        expect([ ...langs.values() ].flat()).toContain("Fußzeile=de-DE")
        expect(await notesLangs(DE, 0)).toEqual([ "de-DE" ])
    })

    it("writes it into el.add textboxes, shapes and tables, with overrides", async () => {
        const langs = await runLangs(DE, 0)
        expect(langs.get("Box")).toEqual([ "Kasten=de-DE" ])
        expect(langs.get("Mixed")).toEqual([ "doos=nl-NL", "box=en-US" ])
        expect(langs.get("Form")).toEqual([ "Form=de-DE" ])
        expect(langs.get("Tab")).toEqual([ "Kopf=de-DE", "Zelle=de-DE" ])
        expect(langs.get("TabFr")).toEqual([ "cellule=fr-FR" ])
    })

    it("keeps the deck language for later applies and el.set", async () => {
        await executeOps(DE, { ops: [ { op: "el.set", slide: "0", name: "Box", text: "Neu" } ] } as never, opts)
        await expectIntact(DE)
        expect((await runLangs(DE, 0)).get("Box")).toEqual([ "Neu=de-DE" ])
    })

    it("reads the language back per shape", async () => {
        const shapes = (await readDeckState(await DeckArchive.open(DE))).slides[0]?.shapes ?? []
        expect(shapes.find((sh) => sh.name === "PptcPh-0")?.lang).toBe("de-DE")
        expect(shapes.find((sh) => sh.name.startsWith("Mixed-"))?.lang).toBe("nl-NL")
        expect(shapes.find((sh) => sh.name.startsWith("TabFr-"))?.lang).toBe("fr-FR")
    })

    it("takes the deck's default text language without pptcDeckLang", async () => {
        const deck = path.join(TMP, "lang-default.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        await executeOps(deck, { ops: [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Title" } } } ] } as never, opts)
        expect((await runLangs(deck, 0)).get("PptcPh-0")).toEqual([ "Title=en-US" ])

        /*  a German deck: the presentation's default text style says de-DE  */
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
        zip.file("ppt/presentation.xml", pres.replace("<a:defPPr><a:defRPr lang=\"en-US\"/>", "<a:defPPr><a:defRPr lang=\"de-DE\"/>"))
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
        await executeOps(deck, { ops: [ { op: "slide.fill", slide: "0", placeholders: { title: { text: "Titel" } } } ] } as never, opts)
        await expectIntact(deck)
        expect((await runLangs(deck, 0)).get("PptcPh-0")).toEqual([ "Titel=de-DE" ])
    })
})
