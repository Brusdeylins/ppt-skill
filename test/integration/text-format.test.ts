/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the format of written text. el.add text without a color
**  takes the template's (no run fill; table headers keep their documented
**  white); el.set merges the given run and paragraph fields onto the kept
**  format; `bullet: true` writes a bullet character where no list style
**  brings one; replaced text drops the old shrink-on-overflow scale;
**  cloned slide-number and date fields take the deck language.
*/

import { describe, expect, it } from "vitest"
import type { Element } from "@xmldom/xmldom"
import { elements, firstElement } from "../../src/engine/xml.js"
import { apply, craftSlide, freshDeck, shapeNamed, slideDoc } from "../util/elements.js"

/**  the fill colors of the runs of a shape: the color value, "-" without  */
const runFills = (shape: Element): string[] =>
    elements(shape, "a:r").map((r) => {
        const fill = firstElement(firstElement(r, "a:rPr") as Element, "a:solidFill")
        return fill === null ? "-" : (fill.firstChild as Element).getAttribute("val") ?? "?"
    })

/**  the bullet facts of each paragraph: bullet kind, marL, indent  */
const bullets = (shape: Element): string[] =>
    elements(shape, "a:p").map((p) => {
        const pPr = firstElement(p, "a:pPr")
        const kind = pPr === null ? [] : [ "a:buChar", "a:buAutoNum", "a:buNone" ].filter((k) => firstElement(pPr, k) !== null)
        return `${kind.join("") || "-"}:${pPr?.getAttribute("marL") ?? ""}:${pPr?.getAttribute("indent") ?? ""}`
    })

describe("el.add text color", () => {
    it("writes no color for text without one -- it inherits the template's", async () => {
        const deck = await freshDeck("format-color.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "textbox", name: "Box", frame: { x: 1, y: 1, w: 4, h: 1 },
                text: [ { runs: [ { text: "plain " }, { text: "red", color: "FF0000" }, { text: " link", hyperlink: "https://example.com" } ] } ] },
            { type: "textbox", name: "Plain", frame: { x: 1, y: 2.2, w: 4, h: 1 }, text: "eins\nzwei" },
            { type: "shape", name: "Shape", shape: "rect", frame: { x: 6, y: 1, w: 3, h: 1 }, text: "shape" },
            { type: "shape", name: "Colored", shape: "rect", frame: { x: 6, y: 2.2, w: 3, h: 1 }, text: "c", fontColor: "accent2" },
            { type: "table", name: "Tab", frame: { x: 1, y: 4, w: 6 }, data: { headers: [ "A", { text: "B", color: "00FF00" } ],
                rows: [ [ "1", { text: "2", color: "0000FF" } ] ] } },
            { type: "table", name: "Themed", frame: { x: 1, y: 5.5, w: 6 }, data: { headers: [ "H" ], rows: [ [ "x" ] ],
                style: { headerFg: "bg1", headerBg: "accent1" } } }
        ] } ])
        const doc = await slideDoc(deck, 0)
        expect(runFills(shapeNamed(doc, "Box"))).toEqual([ "-", "FF0000", "-" ])
        expect(runFills(shapeNamed(doc, "Plain"))).toEqual([ "-", "-" ])
        expect(runFills(shapeNamed(doc, "Shape"))).toEqual([ "-" ])
        expect(runFills(shapeNamed(doc, "Colored"))).toEqual([ "accent2" ])

        /*  the header keeps its documented white on the dark default fill  */
        expect(runFills(shapeNamed(doc, "Tab"))).toEqual([ "FFFFFF", "00FF00", "-", "0000FF" ])
        expect(runFills(shapeNamed(doc, "Themed"))).toEqual([ "bg1", "-" ])

        /*  nothing else in a run's properties carries the token  */
        expect(elements(doc, "a:srgbClr").map((c) => c.getAttribute("val")).filter((v) => /^0C0F/.test(v ?? ""))).toEqual([])
    })
})

describe("el.set text merges onto the kept format", () => {
    it("keeps the run and paragraph format a run or paragraph field does not set", async () => {
        const deck = await freshDeck("format-merge.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "textbox", name: "Box", frame: { x: 1, y: 1, w: 5, h: 2 },
                text: [ { runs: [ { text: "orig", size: 24, color: "C0504D", font: "Georgia", italic: true } ], alignment: "center" } ] } ] } ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "Box", text: [
            { runs: [ { text: "fett", bold: true }, { text: "grün", color: "00B050", italic: false } ] },
            { text: "rechts", alignment: "right", size: 12 },
            { text: "ebene", level: 1 } ] } ])
        const box = shapeNamed(await slideDoc(deck, 0), "Box")
        const rPrs = elements(box, "a:rPr")
        expect(rPrs.map((r) => [ r.getAttribute("b"), r.getAttribute("i"), r.getAttribute("sz"),
            firstElement(r, "a:latin")?.getAttribute("typeface") ])).toEqual([
            [ "1", "1", "2400", "Georgia" ], [ null, "0", "2400", "Georgia" ], [ null, "1", "1200", "Georgia" ], [ null, "1", "2400", "Georgia" ] ])
        expect(runFills(box)).toEqual([ "C0504D", "00B050", "C0504D", "C0504D" ])
        expect(rPrs.every((r) => elements(r, "a:solidFill").length === 1)).toBe(true)
        expect(elements(box, "a:p").map((p) => [ firstElement(p, "a:pPr")?.getAttribute("algn"), firstElement(p, "a:pPr")?.getAttribute("lvl") ?? null ]))
            .toEqual([ [ "ctr", null ], [ "r", null ], [ "ctr", "1" ] ])
    })
})

describe("bullets outside list styles", () => {
    it("writes a bullet character with its hanging indent into textboxes and shapes", async () => {
        const deck = await freshDeck("format-bullet.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { body: { text: "x" } } }, { op: "el.add", slide: "0", elements: [
            { type: "textbox", name: "Listed", frame: { x: 1, y: 1, w: 5, h: 2 }, text: [ { text: "orig", bullet: true } ] },
            { type: "textbox", name: "Plain", frame: { x: 1, y: 3.5, w: 5, h: 2 }, text: "plain" },
            { type: "shape", name: "Form", shape: "rect", frame: { x: 7, y: 1, w: 3, h: 2 }, text: "form" },
            { type: "table", name: "Tab", frame: { x: 7, y: 4, w: 3 }, data: { rows: [ [ "z" ] ] } } ] } ])
        const listed = bullets(shapeNamed(await slideDoc(deck, 0), "Listed"))
        expect(listed).toEqual([ "a:buChar:342900:-342900" ])

        /*  el.set keeps the kept bullet; a box without one gets one  */
        await apply(deck, [
            { op: "el.set", slide: "0", name: "Listed", text: [ { text: "eins", bullet: true }, { text: "zwei", bullet: true } ] },
            { op: "el.set", slide: "0", name: "Plain", text: [ { text: "a", bullet: true }, { text: "b", bullet: true, level: 1 },
                { text: "c", bullet: "number" }, { text: "d", bullet: false } ] },
            { op: "el.set", slide: "0", name: "Form", text: [ { text: "f", bullet: true } ] },
            { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 0, col: 0, text: [ { text: "zelle", bullet: true } ] } ] },
            { op: "el.set", slide: "0", name: "PptcPh-1", text: [ { text: "Liste", bullet: true } ] }
        ])
        const doc = await slideDoc(deck, 0)
        expect(bullets(shapeNamed(doc, "Listed"))).toEqual([ listed[0], listed[0] ])
        expect(bullets(shapeNamed(doc, "Plain"))).toEqual([ "a:buChar:342900:-342900", "a:buChar:685800:-342900",
            "a:buAutoNum:0:0", "a:buNone:0:0" ])
        expect(bullets(shapeNamed(doc, "Form"))[0]).toBe("a:buChar:342900:-342900")
        expect(bullets(shapeNamed(doc, "Tab"))).toEqual([ "a:buChar:342900:-342900" ])
        expect(firstElement(shapeNamed(doc, "Tab"), "a:buChar")?.getAttribute("char")).toBe("•")

        /*  a placeholder keeps the bullet of its list style  */
        expect(bullets(shapeNamed(doc, "PptcPh-1"))).toEqual([ "-::" ])
    })
})

describe("bullets in placeholders whose list style shows none", () => {
    it("writes a bullet character where the inherited level has buNone (a subtitle)", async () => {
        const deck = await freshDeck("format-bullet-ph.pptx")
        await apply(deck, [ { op: "slide.add", layout: "TITLE_SLIDE", placeholders: {
            title: { text: "Titel" }, subtitle: { text: [ { text: "Punkt", bullet: true }, { text: "ohne" } ] } } },
        { op: "slide.add", layout: "CONTENT", placeholders: { body: { text: [ { text: "Liste", bullet: true } ] } } } ])
        const sub = shapeNamed(await slideDoc(deck, 0), "PptcPh-1")
        expect(bullets(sub)).toEqual([ "a:buChar:342900:-342900", "-::" ])

        /*  el.set on the subtitle too; a body with list-style bullets stays without  */
        await apply(deck, [ { op: "el.set", slide: "0", name: "PptcPh-1", text: [ { text: "neu", bullet: true, level: 1 } ] } ])
        expect(bullets(shapeNamed(await slideDoc(deck, 0), "PptcPh-1"))).toEqual([ "a:buChar:685800:-342900" ])
        expect(bullets(shapeNamed(await slideDoc(deck, 1), "PptcPh-1"))).toEqual([ "-::" ])
    })
})

describe("el.set with another level than the kept paragraph", () => {
    it("drops the kept indent, so the level takes effect", async () => {
        const deck = await freshDeck("format-level.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "textbox", name: "Box", frame: { x: 1, y: 1, w: 5, h: 2 }, text: [ { text: "orig", bullet: true } ] } ] } ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "Box", text: [ { text: "gleich", level: 0 }, { text: "tiefer", level: 1 },
            { text: "liste", level: 2, bullet: true } ] } ])
        expect(bullets(shapeNamed(await slideDoc(deck, 0), "Box"))).toEqual([ "a:buChar:342900:-342900",
            "a:buChar:685800:-342900", "a:buChar:1028700:-342900" ])
    })
})

describe("replaced text and the shrink-on-overflow scale", () => {
    it("drops fontScale and lnSpcReduction, keeps normAutofit", async () => {
        const deck = await freshDeck("format-autofit.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { body: { text: "lang" } } },
            { op: "el.add", slide: "0", elements: [ { type: "textbox", name: "Box", frame: { x: 1, y: 1, w: 3, h: 1 }, text: "box" } ] } ])
        await craftSlide(deck, 0, (xml) => xml.replace(/<a:bodyPr([^>]*)\/>/g, "<a:bodyPr$1><a:normAutofit fontScale=\"62500\" lnSpcReduction=\"20000\"/></a:bodyPr>")
            .replace(/<a:bodyPr([^>]*)>(?!<a:normAutofit)/g, "<a:bodyPr$1><a:normAutofit fontScale=\"62500\" lnSpcReduction=\"20000\"/>"))
        const fits = (shape: Element): (string | null)[][] =>
            elements(shape, "a:normAutofit").map((f) => [ f.getAttribute("fontScale"), f.getAttribute("lnSpcReduction") ])
        let doc = await slideDoc(deck, 0)
        expect(fits(shapeNamed(doc, "PptcPh-1"))).toEqual([ [ "62500", "20000" ] ])
        expect(fits(shapeNamed(doc, "Box"))).toEqual([ [ "62500", "20000" ] ])

        await apply(deck, [ { op: "slide.fill", slide: "0", placeholders: { body: { text: "kurz" } } },
            { op: "el.set", slide: "0", name: "Box", text: "neu" } ])
        doc = await slideDoc(deck, 0)
        expect(fits(shapeNamed(doc, "PptcPh-1"))).toEqual([ [ null, null ] ])
        expect(fits(shapeNamed(doc, "Box"))).toEqual([ [ null, null ] ])
    })
})

describe("cloned footer-band fields", () => {
    it("take the deck language", async () => {
        const deck = await freshDeck("format-fields.pptx")
        await apply(deck, [ { op: "meta.props", set: { custom: { pptcDeckLang: "de-DE" } } },
            { op: "slide.add", layout: "CONTENT", slideNumber: true, date: true, footer: "Fuß" } ])
        const doc = await slideDoc(deck, 0)
        const band = elements(doc, "p:sp").filter((sp) => [ "sldNum", "dt", "ftr" ].includes(firstElement(sp, "p:ph")?.getAttribute("type") ?? ""))
        expect(band.length).toBe(3)
        const langs = band.flatMap((sp) => [ ...elements(sp, "a:rPr"), ...elements(sp, "a:endParaRPr") ].map((r) => r.getAttribute("lang")))
        expect(langs.length).toBeGreaterThanOrEqual(3)
        expect(new Set(langs)).toEqual(new Set([ "de-DE" ]))
        expect(elements(doc, "a:fld").length).toBe(2)
    })
})
