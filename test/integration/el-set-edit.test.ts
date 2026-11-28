/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: el.set changes existing elements in place -- frame (move,
**  resize, rotate; an inherited placeholder frame is written into the
**  slide; a group keeps its children's coordinates), fill, border, font
**  color, alt text and title -- each field checked against the element's
**  type; elements planned earlier in the document get their spec patched,
**  the overlap lint sees planned frames, el.set text keeps the first run's
**  format, and el.add never describes a picture by its file path.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { elements, firstElement } from "../../src/engine/xml.js"
import { TMP, apply, craftSlide, failure, freshDeck, png, shapeNamed, slideDoc, stateOf } from "../util/elements.js"

const IMG = path.join(TMP, "el-set-edit.png")
let deck: string

/**  a group of two textboxes in child coordinates (el.add makes none)  */
const GROUP = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"900\" name=\"Gruppe\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
    + "<p:grpSpPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"914400\" cy=\"914400\"/>"
    + "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"914400\" cy=\"914400\"/></a:xfrm></p:grpSpPr>"
    + [ 901, 902 ].map((id) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Kind ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
        + `<p:spPr><a:xfrm><a:off x="0" y="${(id - 901) * 457200}"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>`
        + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>child ${id}</a:t></a:r></a:p></p:txBody></p:sp>`).join("")
    + "</p:grpSp>"

beforeAll(async () => {
    writeFileSync(IMG, png(400, 300))
    deck = await freshDeck("el-set-edit.pptx")
    await apply(deck, [
        { op: "slide.add", layout: "CONTENT", ref: "s", placeholders: { title: { text: "Edit" } } },
        { op: "el.add", slide: "$s", elements: [
            { type: "textbox", name: "Box", frame: { x: 0.5, y: 5.5, w: 3, h: 0.6 },
                text: [ { runs: [ { text: "fett rot", bold: true, color: "C00000", size: 20 } ], alignment: "center" } ] },
            { type: "shape", name: "Form", shape: "roundRect", frame: { x: 4, y: 5.5, w: 2, h: 0.6 }, text: "form" },
            { type: "table", name: "Tab", frame: { x: 6.5, y: 5.5, w: 3 }, data: { rows: [ [ "a", "b" ], [ "c", "d" ] ] } },
            { type: "chart", name: "Diagramm", frame: { x: 10, y: 4, w: 3, h: 2 }, altText: "Umsatz",
                data: { type: "column", categories: [ "x" ], series: [ { name: "s", values: [ 1 ] } ] } },
            { type: "image", name: "Bild", frame: { x: 0.5, y: 6.5, w: 0.5, h: 0.5 }, path: IMG, title: "Foto" },
            { type: "connector", name: "Linie", from: [ 2, 6.6 ], to: [ 4, 6.9 ] }
        ] }
    ])
    await craftSlide(deck, 0, (xml) => xml.replace("</p:spTree>", `${GROUP}</p:spTree>`))
})

describe("el.set frame", () => {
    it("moves, resizes and rotates a textbox; state reads it back", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Box", frame: { x: 1, w: 2.5, rotation: -90 } } ])
        const box = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Box-"))
        expect(box?.frame).toEqual({ x: 1, y: 5.5, w: 2.5, h: 0.6 })
        expect(box?.rotation).toBe(270)
        await apply(deck, [ { op: "el.set", slide: "0", name: "Box", frame: { rotation: 0 } } ])
        expect((await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Box-"))?.rotation).toBeUndefined()
    })

    it("writes an inherited placeholder frame into the slide", async () => {
        const before = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.placeholderKind === "title")
        expect(before?.frame).toBeNull()
        await apply(deck, [ { op: "el.set", slide: "0", name: before?.name ?? "", frame: { y: 0.3 } } ])
        const after = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.placeholderKind === "title")
        expect(after?.frame?.y).toBe(0.3)
        expect(after?.frame?.w).toBeGreaterThan(0)
        expect(after?.text).toBe("Edit")
    })

    it("moves a group without touching its children's coordinates", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Gruppe", frame: { x: 2, y: 2, w: 2 } } ])
        const group = shapeNamed(await slideDoc(deck, 0), "Gruppe")
        const xfrm = firstElement(group, "a:xfrm")
        expect(firstElement(xfrm as never, "a:off")?.getAttribute("x")).toBe("1828800")
        expect(firstElement(xfrm as never, "a:ext")?.getAttribute("cx")).toBe("1828800")
        expect(firstElement(xfrm as never, "a:chExt")?.getAttribute("cx")).toBe("914400")
        const child = elements(group, "p:sp")[1]
        expect(firstElement(child as never, "a:off")?.getAttribute("y")).toBe("457200")
    })

    it("scales a table's columns to its new width", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Tab", frame: { w: 4 } } ])
        const tab = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Tab-"))
        expect(tab?.frame?.w).toBe(4)
        expect(tab?.colWidths).toEqual([ 2, 2 ])
    })
})

describe("el.set style and alt text", () => {
    it("sets fill, border and font color of a shape", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Form", fill: "accent2", border: "00B050", borderPt: 2.5, fontColor: "FFFFFF" } ])
        const form = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Form-"))
        expect(form?.border).toBe("00B050")
        expect(form?.borderPt).toBe(2.5)
        expect(form?.fontColor).toBe("FFFFFF")
        expect(firstElement(firstElement(shapeNamed(await slideDoc(deck, 0), "Form"), "p:spPr") as never, "a:schemeClr")?.getAttribute("val")).toBe("accent2")
    })

    it("sets the outline of a picture and a connector", async () => {
        await apply(deck, [
            { op: "el.set", slide: "0", name: "Bild", border: "000000", borderPt: 1 },
            { op: "el.set", slide: "0", name: "Linie", border: "FF0000", borderPt: 3 }
        ])
        const shapes = (await stateOf(deck)).slides[0]?.shapes ?? []
        expect(shapes.find((sh) => sh.name.startsWith("Linie-"))?.border).toBe("FF0000")
        expect(shapes.find((sh) => sh.name.startsWith("Linie-"))?.borderPt).toBe(3)
    })

    it("writes alt text and title, an empty value removes them", async () => {
        let shapes = (await stateOf(deck)).slides[0]?.shapes ?? []
        expect(shapes.find((sh) => sh.name.startsWith("Diagramm-"))?.altText).toBe("Umsatz")
        expect(shapes.find((sh) => sh.name.startsWith("Bild-"))?.title).toBe("Foto")
        await apply(deck, [
            { op: "el.set", slide: "0", name: "Tab", altText: "Zahlen", title: "Tabelle" },
            { op: "el.set", slide: "0", name: "Gruppe", altText: "Gruppe A" },
            { op: "el.set", slide: "0", name: "Diagramm", altText: "" }
        ])
        shapes = (await stateOf(deck)).slides[0]?.shapes ?? []
        expect(shapes.find((sh) => sh.name.startsWith("Tab-"))).toMatchObject({ altText: "Zahlen", title: "Tabelle" })
        expect(shapes.find((sh) => sh.name === "Gruppe")?.altText).toBe("Gruppe A")
        expect(shapes.find((sh) => sh.name.startsWith("Diagramm-"))?.altText).toBeUndefined()
    })

    it("never describes an el.add picture by its local file path", async () => {
        const zipText = readFileSync(deck).toString("latin1")
        expect(zipText.includes(TMP)).toBe(false)
        const descr = elements(await slideDoc(deck, 0), "p:cNvPr").map((el) => el.getAttribute("descr") ?? "")
        expect(descr.some((d) => d.includes("/") || d.endsWith(".png"))).toBe(false)
    })
})

describe("el.set text keeps the format", () => {
    it("keeps the first paragraph's and run's format when the new text sets none", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Box", text: "neu\nzweite" } ])
        const box = shapeNamed(await slideDoc(deck, 0), "Box")
        const paras = elements(box, "a:p")
        expect(paras.map((p) => firstElement(p, "a:pPr")?.getAttribute("algn"))).toEqual([ "ctr", "ctr" ])
        expect(elements(box, "a:rPr").map((r) => [ r.getAttribute("b"), r.getAttribute("sz") ])).toEqual([ [ "1", "2000" ], [ "1", "2000" ] ])
        expect(elements(box, "a:srgbClr").map((c) => c.getAttribute("val"))).toEqual([ "C00000", "C00000" ])
    })

    it("takes the format the new text gives, merged onto the kept one", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Box", text: [ { runs: [ { text: "kursiv", italic: true } ], alignment: "right" } ] } ])
        const box = shapeNamed(await slideDoc(deck, 0), "Box")
        expect(firstElement(box, "a:pPr")?.getAttribute("algn")).toBe("r")
        const rPr = firstElement(box, "a:rPr")
        expect([ rPr?.getAttribute("i"), rPr?.getAttribute("b"), rPr?.getAttribute("sz") ]).toEqual([ "1", "1", "2000" ])
        expect(elements(box, "a:srgbClr").map((c) => c.getAttribute("val"))).toEqual([ "C00000" ])
    })
})

describe("el.set type checks", () => {
    const cases: [ string, Record<string, unknown> ][] = [
        [ "Bild", { fill: "FF0000" } ],
        [ "Box", { image: IMG } ],
        [ "Tab", { frame: { rotation: 45 } } ],
        [ "Diagramm", { fontColor: "FF0000" } ],
        [ "Form", { cells: [ { row: 0, col: 0, text: "x" } ] } ],
        [ "Linie", { fill: "FF0000" } ],
        [ "Gruppe", { border: "FF0000" } ]
    ]
    for (const [ name, set ] of cases)
        it(`rejects ${Object.keys(set).join(",")} on ${name}`, async () => {
            const before = readFileSync(deck)
            expect(await failure(deck, [ { op: "el.set", slide: "0", name, ...set } ])).toBe("E_SCHEMA")
            expect(readFileSync(deck).equals(before)).toBe(true)
        })

    it("needs at least one field", async () => {
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "Box" } ])).toBe("E_SCHEMA")
    })

    it("rejects a missing element and one removed earlier in the document", async () => {
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "Nichts", fill: "FF0000" } ])).toBe("E_ADDR_NOTFOUND")
        expect(await failure(deck, [ { op: "el.rm", slide: "0", name: "Form" },
            { op: "el.set", slide: "0", name: "Form", fill: "FF0000" } ])).toBe("E_ADDR_NOTFOUND")
    })
})

describe("el.set on elements planned in the same document", () => {
    it("patches the spec of an el.add element", async () => {
        await apply(deck, [
            { op: "slide.add", layout: "Title Only", ref: "p", placeholders: { title: { text: "Geplant" } } },
            { op: "el.add", slide: "$p", elements: [
                { type: "textbox", name: "Neu", frame: { x: 1, y: 2, w: 2, h: 1 }, text: "neu" },
                { type: "connector", name: "Strich", from: [ 1, 4 ], to: [ 3, 4 ] },
                { type: "image", name: "Pic", frame: { x: 5, y: 2, w: 1, h: 1 }, path: IMG } ] },
            { op: "el.set", slide: "$p", name: "Neu", frame: { x: 2, rotation: 90 }, fill: "accent1", fontColor: "FFFFFF", altText: "Kasten" },
            { op: "el.set", slide: "$p", name: "Strich", border: "00FF00", borderPt: 4 },
            { op: "el.set", slide: "$p", name: "Pic", altText: "Bild", frame: { w: 2 } }
        ])
        const shapes = (await stateOf(deck)).slides[1]?.shapes ?? []
        const neu = shapes.find((sh) => sh.name.startsWith("Neu-"))
        expect(neu).toMatchObject({ frame: { x: 2, y: 2, w: 2, h: 1 }, rotation: 90, altText: "Kasten" })
        expect(shapes.find((sh) => sh.name.startsWith("Strich-"))).toMatchObject({ border: "00FF00", borderPt: 4 })
        expect(shapes.find((sh) => sh.name.startsWith("Pic-"))).toMatchObject({ altText: "Bild", frame: { w: 2 } })
    })

    it("rejects fields a planned element does not take", async () => {
        expect(await failure(deck, [
            { op: "el.add", slide: "1", elements: [ { type: "connector", name: "C2", from: [ 1, 5 ], to: [ 2, 5 ] } ] },
            { op: "el.set", slide: "1", name: "C2", frame: { x: 1 } } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [
            { op: "el.add", slide: "1", elements: [ { type: "image", name: "I2", frame: { x: 1, y: 5, w: 1, h: 1 }, path: IMG } ] },
            { op: "el.set", slide: "1", name: "I2", fill: "FF0000" } ])).toBe("E_SCHEMA")
    })
})

describe("el.set frames and the overlap lint", () => {
    it("warns when a move covers a text shape", async () => {
        const title = (await stateOf(deck)).slides[1]?.shapes.find((sh) => sh.placeholderKind === "title")
        const warnings = await apply(deck, [ { op: "el.set", slide: "1", name: "Neu", frame: { x: title?.frame?.x ?? 0.5, y: 0.6 } } ])
        expect(warnings).toEqual(expect.arrayContaining([ expect.objectContaining({ code: "W_ELEMENT_OVERLAP" }) ]))
    })

    it("sees a placeholder moved earlier in the document", async () => {
        const title = (await stateOf(deck)).slides[1]?.shapes.find((sh) => sh.placeholderKind === "title")
        const frame = { x: 4, y: 0.6, w: 3, h: 0.5 }
        const covered = await apply(deck, [ { op: "el.add", slide: "1", elements: [
            { type: "textbox", name: "Probe1", frame, text: "x" } ] } ])
        expect(covered.some((w) => (w as { code: string }).code === "W_ELEMENT_OVERLAP")).toBe(true)
        const moved = await apply(deck, [
            { op: "el.set", slide: "1", name: title?.name ?? "", frame: { x: 0.5, y: 6, w: 4, h: 0.6 } },
            { op: "el.rm", slide: "1", name: "Probe1" },
            { op: "el.add", slide: "1", elements: [ { type: "textbox", name: "Probe2", frame, text: "x" } ] } ])
        expect(moved.filter((w) => (w as { element?: string }).element === "Probe2")).toEqual([])
    })
})
