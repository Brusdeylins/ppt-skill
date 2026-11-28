/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the deck language -- normalization of a stated language, the
**  precedence of its sources (pptcDeckLang, presentation default text
**  style, first run on a slide, en-US) and reading them from a package.
*/

import { describe, expect, it } from "vitest"
import JSZip from "jszip"
import { deckLang, normalizeLang, pickDeckLang } from "../../src/engine/lang.js"

const NS = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
    + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
    + "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""

/**  a minimal package: presentation (with an optional default text
     language), one slide (with an optional run language), custom props  */
const pkg = async (opts: { presLang?: string, runLang?: string, custom?: string }): Promise<JSZip> => {
    const zip = new JSZip()
    const style = opts.presLang === undefined ? ""
        : `<p:defaultTextStyle><a:defPPr><a:defRPr lang="xx-XX"/></a:defPPr><a:lvl1pPr><a:defRPr lang="${opts.presLang}"/></a:lvl1pPr></p:defaultTextStyle>`
    zip.file("ppt/presentation.xml", `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>${style}</p:presentation>`)
    zip.file("ppt/_rels/presentation.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>"
        + "</Relationships>")
    const run = opts.runLang === undefined ? "<a:r><a:t>x</a:t></a:r>" : `<a:r><a:rPr lang="${opts.runLang}"/><a:t>x</a:t></a:r>`
    zip.file("ppt/slides/slide1.xml", `<p:sld ${NS}><p:cSld><p:spTree><p:sp><p:txBody><a:p>${run}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
    if (opts.custom !== undefined)
        zip.file("docProps/custom.xml", "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" "
            + "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
            + `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="pptcDeckLang"><vt:lpwstr>${opts.custom}</vt:lpwstr></property>`
            + "</Properties>")
    return zip
}

describe("language normalization", () => {
    it("keeps tags, normalizing their case", () => {
        expect(normalizeLang("de-DE")).toBe("de-DE")
        expect(normalizeLang("DE-at")).toBe("de-AT")
        expect(normalizeLang(" en-GB ")).toBe("en-GB")
        expect(normalizeLang("zh-Hant-TW")).toBe("zh-Hant-TW")
    })

    it("gives bare codes and language names their usual region", () => {
        expect(normalizeLang("de")).toBe("de-DE")
        expect(normalizeLang("en")).toBe("en-US")
        expect(normalizeLang("German")).toBe("de-DE")
        expect(normalizeLang("Deutsch")).toBe("de-DE")
        expect(normalizeLang("français")).toBe("fr-FR")
    })

    it("names no language for anything else", () => {
        expect(normalizeLang("")).toBeNull()
        expect(normalizeLang(undefined)).toBeNull()
        expect(normalizeLang("Klingon please")).toBeNull()
    })
})

describe("deck language precedence", () => {
    it("prefers pptcDeckLang, then the presentation, then the first run", () => {
        expect(pickDeckLang("de", "en-US", "fr-FR")).toBe("de-DE")
        expect(pickDeckLang(null, "en-GB", "fr-FR")).toBe("en-GB")
        expect(pickDeckLang("no language", null, "fr-FR")).toBe("fr-FR")
        expect(pickDeckLang(null, null, null)).toBe("en-US")
    })

    it("reads the sources from a package", async () => {
        expect(await deckLang(await pkg({ presLang: "en-US", runLang: "fr-FR", custom: "de-DE" }), null)).toBe("de-DE")
        expect(await deckLang(await pkg({ presLang: "de-DE", runLang: "fr-FR" }), null)).toBe("de-DE")
        expect(await deckLang(await pkg({ presLang: "", runLang: "fr-FR" }), null)).toBe("xx-XX")
        expect(await deckLang(await pkg({ runLang: "fr-FR" }), null)).toBe("fr-FR")
        expect(await deckLang(await pkg({}), null)).toBe("en-US")
    })

    it("lets the apply's own custom properties win", async () => {
        const zip = await pkg({ presLang: "en-US", custom: "fr-FR" })
        expect(await deckLang(zip, { pptcDeckLang: "Deutsch" })).toBe("de-DE")
        expect(await deckLang(zip, { other: "x" })).toBe("fr-FR")
        expect(await deckLang(zip, { pptcDeckLang: "" })).toBe("en-US")
    })
})
