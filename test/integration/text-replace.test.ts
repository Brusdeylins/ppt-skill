/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: text.replace finds and replaces plain text in shapes,
**  placeholders and table cells (notes on request, fields never), across
**  runs split by formatting (the first touched run keeps its format), with
**  sequence semantics -- it sees the deck after every earlier op of the
**  document and nothing of the later ones -- and reports its replacements
**  per output slide.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import JSZip from "jszip"
import { elements, firstElement } from "../../src/engine/xml.js"
import { apply, applyResult, craftSlide, failure, freshDeck, shapeNamed, slideDoc, stateOf } from "../util/elements.js"

let deck: string

/**  a textbox with a text split over formatted runs and a field  */
const SPLIT = "<p:sp><p:nvSpPr><p:cNvPr id=\"950\" name=\"Split\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>"
    + "<p:spPr><a:xfrm><a:off x=\"0\" y=\"5500000\"/><a:ext cx=\"3000000\" cy=\"400000\"/></a:xfrm>"
    + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>"
    + "<a:p><a:r><a:rPr lang=\"de-DE\" b=\"1\"/><a:t>Hal</a:t></a:r><a:r><a:rPr lang=\"de-DE\" i=\"1\"/><a:t>lo We</a:t></a:r>"
    + "<a:r><a:rPr lang=\"de-DE\"/><a:t>lt!</a:t></a:r></a:p>"
    + "<a:p><a:r><a:rPr lang=\"de-DE\"/><a:t>Seite </a:t></a:r><a:fld id=\"{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}\" type=\"slidenum\">"
    + "<a:rPr lang=\"de-DE\"/><a:t>Hallo</a:t></a:fld><a:r><a:rPr lang=\"de-DE\"/><a:t> Hallo</a:t></a:r></a:p>"
    + "</p:txBody></p:sp>"

/**  the texts of a slide by shape name prefix  */
const texts = async (index: number): Promise<Record<string, string | null>> =>
    Object.fromEntries(((await stateOf(deck)).slides[index]?.shapes ?? [])
        .map((sh) => [ sh.name.replace(/-[0-9a-f-]{36}$/, ""), sh.table !== undefined ? JSON.stringify(sh.table) : sh.text ]))

beforeAll(async () => {
    deck = await freshDeck("text-replace.pptx")
    await apply(deck, [
        { op: "slide.add", layout: "CONTENT", ref: "a", notes: "Hallo Welt in den Notizen",
            placeholders: { title: { text: "Hallo Welt" }, body: { text: "hallo welt\nHalloWelt" } } },
        { op: "el.add", slide: "$a", elements: [ { type: "table", name: "Tab", frame: { x: 0.5, y: 6, w: 4 },
            data: { headers: [ "Hallo", "Welt" ], rows: [ [ "Hallo Welt", 1 ] ] } } ] },
        { op: "slide.add", layout: "CONTENT", ref: "b", placeholders: { title: { text: "Zweite Welt" } } }
    ])
    await craftSlide(deck, 0, (xml) => xml.replace("</p:spTree>", `${SPLIT}</p:spTree>`))
})

describe("text.replace on existing content", () => {
    it("replaces across split runs, keeps the first run's format and leaves fields alone", async () => {
        const result = await applyResult(deck, [ { op: "text.replace", find: "Hallo Welt", replace: "Servus Erde", slide: "0", matchCase: true } ])
        const split = shapeNamed(await slideDoc(deck, 0), "Split")
        const [ p1, p2 ] = elements(split, "a:p")
        expect(elements(p1 as never, "a:t").map((t) => t.textContent)).toEqual([ "Servus Erde", "!" ])
        expect(firstElement(p1 as never, "a:rPr")?.getAttribute("b")).toBe("1")
        expect(elements(p2 as never, "a:t").map((t) => t.textContent)).toEqual([ "Seite ", "Hallo", " Hallo" ])
        const t = await texts(0)
        expect(t["PptcPh-0"]).toBe("Servus Erde")
        expect(t["PptcPh-1"]).toBe("hallo welt\nHalloWelt")
        expect(t["Tab"]).toBe(JSON.stringify([ [ "Hallo", "Welt" ], [ "Servus Erde", "1" ] ]))
        expect((await stateOf(deck)).slides[0]?.notes).toBe("Hallo Welt in den Notizen")
        expect(result.result.replacements).toEqual([ { op: 0, count: 3, slides: [ { id: 256, index: 0, count: 3 } ] } ])
    })

    it("matches case-insensitively by default, whole words on request, notes on request", async () => {
        const result = await applyResult(deck, [
            { op: "text.replace", find: "welt", replace: "Erde", wholeWord: true, notes: true }
        ])
        const t = await texts(0)
        expect(t["PptcPh-1"]).toBe("hallo Erde\nHalloWelt")
        expect(t["Tab"]).toBe(JSON.stringify([ [ "Hallo", "Erde" ], [ "Servus Erde", "1" ] ]))
        expect((await stateOf(deck)).slides[0]?.notes).toBe("Hallo Erde in den Notizen")
        expect((await texts(1))["PptcPh-0"]).toBe("Zweite Erde")
        expect(result.result.replacements).toEqual([ { op: 0, count: 4, slides: [
            { id: 256, index: 0, count: 3 }, { id: 257, index: 1, count: 1 } ] } ])
    })

    it("reports zero matches without failing and leaves slides without a match untouched", async () => {
        const slideBytes = async (): Promise<string[]> => {
            const zip = await JSZip.loadAsync(readFileSync(deck))
            return await Promise.all((await stateOf(deck)).slides.map(async (sl) =>
                await (zip.file(sl.part) as JSZip.JSZipObject).async("string")))
        }
        const before = await slideBytes()
        const result = await applyResult(deck, [ { op: "text.replace", find: "nirgends", replace: "x", notes: true } ])
        expect(result.result.replacements).toEqual([ { op: 0, count: 0, slides: [] } ])
        expect(await slideBytes()).toEqual(before)
    })

    it("rejects an empty find and an unknown slide", async () => {
        expect(await failure(deck, [ { op: "text.replace", find: "", replace: "x" } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "text.replace", find: "x", replace: "y", slide: "id:9999" } ])).toBe("E_ADDR_NOTFOUND")
    })

    it("rejects a line break in find or replace (it works within a paragraph)", async () => {
        for (const [ find, replace ] of [ [ "a\nb", "x" ], [ "a\rb", "x" ], [ "Hallo", "Zeile 1\nZeile 2" ], [ "Hallo", "x\r" ] ])
            expect(await failure(deck, [ { op: "text.replace", find, replace } ])).toBe("E_SCHEMA")
    })
})

describe("text.replace in sequence", () => {
    it("sees what earlier ops planned and nothing of the later ones", async () => {
        const result = await applyResult(deck, [
            { op: "el.add", slide: "1", elements: [ { type: "textbox", name: "Vorher", frame: { x: 1, y: 2, w: 3, h: 1 }, text: "foo" } ] },
            { op: "slide.fill", slide: "1", notes: "foo notes", footer: "foo Fuß",
                placeholders: { body: { text: [ { runs: [ { text: "f", bold: true }, { text: "oo" } ] } ] } } },
            { op: "text.replace", find: "foo", replace: "bar", slide: "1", notes: true },
            { op: "el.add", slide: "1", elements: [ { type: "textbox", name: "Nachher", frame: { x: 5, y: 2, w: 3, h: 1 }, text: "foo" } ] },
            { op: "text.replace", find: "bar", replace: "baz", slide: "1" }
        ])
        const t = await texts(1)
        expect(t["Vorher"]).toBe("baz")
        expect(t["Nachher"]).toBe("foo")
        expect(t["PptcPh-1"]).toBe("baz")
        expect((await stateOf(deck)).slides[1]?.notes).toBe("bar notes")
        expect(Object.entries(t).find(([ k ]) => k.startsWith("Footer"))?.[1]).toBe("baz Fuß")
        expect(result.result.replacements).toEqual([
            { op: 2, count: 4, slides: [ { id: 257, index: 1, count: 4 } ] },
            { op: 4, count: 3, slides: [ { id: 257, index: 1, count: 3 } ] } ])
    })

    it("does not see text an earlier el.set replaced, but text a later one replaces", async () => {
        await apply(deck, [ { op: "el.add", slide: "1", elements: [
            { type: "textbox", name: "Eins", frame: { x: 1, y: 3.5, w: 2, h: 0.5 }, text: "alt" },
            { type: "textbox", name: "Zwei", frame: { x: 4, y: 3.5, w: 2, h: 0.5 }, text: "alt" } ] } ])
        const result = await applyResult(deck, [
            { op: "el.set", slide: "1", name: "Eins", text: "neu alt" },
            { op: "text.replace", find: "alt", replace: "ALT", slide: "1", matchCase: true },
            { op: "el.set", slide: "1", name: "Zwei", text: "alt bleibt" }
        ])
        const t = await texts(1)
        expect(t["Eins"]).toBe("neu ALT")
        expect(t["Zwei"]).toBe("alt bleibt")
        expect(result.result.replacements?.[0]?.count).toBe(2)
    })

    it("follows table edits, copies and title selectors", async () => {
        const result = await applyResult(deck, [
            { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 0, col: 0, text: "Kopf" } ], addRows: { rows: [ [ "Kopf", "x" ] ] } },
            { op: "text.replace", find: "Kopf", replace: "Haupt", slide: "0" },
            { op: "text.replace", find: "Servus", replace: "Grüezi" },
            { op: "slide.copy", slide: "title:Grüezi Erde", ref: "c" },
            { op: "slide.fill", slide: "$c", placeholders: { title: { text: "Kopie" } } }
        ])
        expect((await texts(0))["Tab"]).toBe(JSON.stringify([ [ "Haupt", "Erde" ], [ "Grüezi Erde", "1" ], [ "Haupt", "x" ] ]))
        expect((await stateOf(deck)).slides.map((s) => s.title)).toEqual([ "Grüezi Erde", "Kopie", "Zweite Erde" ])
        expect((await texts(1))["Split"]).toBe("Grüezi Erde!\nSeite Hallo Hallo")
        expect(result.result.replacements?.map((r) => [ r.op, r.count, r.slides.map((s) => s.index) ])).toEqual([
            [ 1, 4, [ 0, 1 ] ], [ 2, 6, [ 0, 1 ] ] ])
    })
})
