/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: fills and shapes. A solid fill takes a `transparency`
**  (the color's `a:alpha`) on el.add shapes, textboxes and table cells and
**  on el.set (a new fill keeps it); text without a color on a fill takes
**  the theme's dark or light slot that contrasts with it; the curated
**  preset geometries come back from `state`; the theme fonts `+mj-lt` and
**  `+mn-lt` are written as the run's latin typeface.
*/

import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import JSZip from "jszip"
import type { Element } from "@xmldom/xmldom"
import { ShapeTypeSchema } from "../../src/schema/payloads.js"
import { elements, firstElement } from "../../src/engine/xml.js"
import { apply, attrs, failure, freshDeck, shapeNamed, slideDoc, stateOf } from "../util/elements.js"

/**  the alpha values of the colors of a shape's own solid fill  */
const alphaOf = (shape: Element): (string | null)[] =>
    attrs(firstElement(firstElement(shape, "p:spPr") as Element, "a:solidFill") as Element, "a:alpha", "val")

/**  the state of an element on slide 0 by name prefix  */
const stateShape = async (deck: string, name: string): Promise<Record<string, unknown>> =>
    (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name === name || sh.name.startsWith(`${name}-`)) as never

/**  the scheme or RGB color values of the run fills of a shape or cell  */
const runColors = (scope: Element): string[] =>
    elements(scope, "a:rPr").map((rPr) => {
        const fill = firstElement(rPr, "a:solidFill")
        return fill === null ? "-" : (fill.firstChild as Element).getAttribute("val") ?? "?"
    })

describe("fill transparency", () => {
    it("writes el.add transparency into the fill color and reads it back", async () => {
        const deck = await freshDeck("fill-alpha.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "shape", name: "Card", shape: "roundRect", frame: { x: 1, y: 1, w: 3, h: 1 }, fill: "accent1", transparency: 40, text: "a" },
            { type: "textbox", name: "Box", frame: { x: 5, y: 1, w: 3, h: 1 }, fill: "#FF0000", transparency: 25, text: "b" },
            { type: "table", name: "Tab", frame: { x: 1, y: 3, w: 4 }, data: { rows: [ [ { text: "c", fill: "00FF00", transparency: 50 }, "d" ] ] } }
        ] } ])
        const doc = await slideDoc(deck, 0)
        expect(alphaOf(shapeNamed(doc, "Card"))).toEqual([ "60000" ])
        expect(firstElement(firstElement(shapeNamed(doc, "Card"), "p:spPr") as Element, "a:schemeClr")?.getAttribute("val")).toBe("accent1")
        expect(alphaOf(shapeNamed(doc, "Box"))).toEqual([ "75000" ])
        expect(attrs(firstElement(shapeNamed(doc, "Tab"), "a:tcPr") as Element, "a:alpha", "val")).toEqual([ "50000" ])
        expect((await stateShape(deck, "Card")).transparency).toBe(40)
        expect((await stateShape(deck, "Box")).transparency).toBe(25)
    })

    it("sets el.set transparency with or without a new fill; a new fill keeps it", async () => {
        const deck = await freshDeck("fill-alpha-set.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "slide.add", layout: "DEFAULT" },
            { op: "el.add", slide: "0", elements: [
                { type: "shape", name: "Card", shape: "rect", frame: { x: 1, y: 1, w: 3, h: 1 }, fill: "1F4E79" },
                { type: "textbox", name: "Plain", frame: { x: 5, y: 1, w: 3, h: 1 }, text: "t" } ] },
            { op: "el.add", slide: "1", elements: [ { type: "textbox", name: "Other", frame: { x: 1, y: 1, w: 3, h: 1 }, text: "o" } ] } ])
        const parts = (await stateOf(deck)).slides.map((s) => s.part)
        const otherBytes = async (): Promise<Buffer> =>
            await ((await JSZip.loadAsync(readFileSync(deck))).file(parts[1] as string) as JSZip.JSZipObject).async("nodebuffer")
        const other = await otherBytes()

        await apply(deck, [ { op: "el.set", slide: "0", name: "Card", transparency: 30 } ])
        expect(alphaOf(shapeNamed(await slideDoc(deck, 0), "Card"))).toEqual([ "70000" ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "Card", fill: "accent2" } ])
        expect(await stateShape(deck, "Card")).toMatchObject({ transparency: 30 })
        expect(firstElement(firstElement(shapeNamed(await slideDoc(deck, 0), "Card"), "p:spPr") as Element, "a:schemeClr")
            ?.getAttribute("val")).toBe("accent2")
        await apply(deck, [ { op: "el.set", slide: "0", name: "Card", transparency: 0 } ])
        expect(alphaOf(shapeNamed(await slideDoc(deck, 0), "Card"))).toEqual([])
        expect((await stateShape(deck, "Card")).transparency).toBeUndefined()

        /*  without a fill of its own only together with one  */
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "Plain", transparency: 50 } ])).toBe("E_SCHEMA")
        await apply(deck, [ { op: "el.set", slide: "0", name: "Plain", fill: "lt2", transparency: 50 } ])
        expect((await stateShape(deck, "Plain")).transparency).toBe(50)
        await apply(deck, [ { op: "el.set", slide: "0", name: "Plain", fill: "FFFFFF" },
            { op: "el.set", slide: "0", name: "Plain", transparency: 20 } ])
        expect((await stateShape(deck, "Plain")).transparency).toBe(20)
        expect((await otherBytes()).equals(other)).toBe(true)

        /*  a planned element: its spec takes it  */
        await apply(deck, [ { op: "el.add", slide: "0", elements: [ { type: "shape", name: "New", shape: "ellipse",
            frame: { x: 9, y: 1, w: 1, h: 1 }, fill: "accent3" } ] }, { op: "el.set", slide: "0", name: "New", transparency: 60 } ])
        expect(alphaOf(shapeNamed(await slideDoc(deck, 0), "New"))).toEqual([ "40000" ])
    })

    it("rejects a transparency without a fill and wrong values at plan time", async () => {
        const deck = await freshDeck("fill-alpha-schema.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "textbox", name: "Plain", frame: { x: 1, y: 1, w: 3, h: 1 }, text: "t" } ] } ])
        const before = readFileSync(deck)
        const add = (el: Record<string, unknown>): unknown[] => [ { op: "el.add", slide: "0", elements: [ el ] } ]
        const shape = { type: "shape", shape: "rect", frame: { x: 1, y: 3, w: 1, h: 1 } }
        expect(await failure(deck, add({ ...shape, transparency: 50 }))).toBe("E_SCHEMA")
        expect(await failure(deck, add({ ...shape, fill: "accent1", transparency: 101 }))).toBe("E_SCHEMA")
        expect(await failure(deck, add({ ...shape, fill: "accent1", transparency: "50" }))).toBe("E_SCHEMA")
        expect(await failure(deck, add({ type: "table", frame: { x: 1, y: 3, w: 2 }, data: { rows: [ [ { text: "c", transparency: 10 } ] ] } })))
            .toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "Plain", transparency: -1 } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ ...add({ type: "textbox", name: "N", frame: { x: 5, y: 3, w: 1, h: 1 }, text: "n" }),
            { op: "el.set", slide: "0", name: "N", transparency: 10 } ])).toBe("E_SCHEMA")
        expect(readFileSync(deck).equals(before)).toBe(true)
    })
})

describe("text on a fill", () => {
    it("takes the theme's dark or light slot that contrasts with the fill", async () => {
        const deck = await freshDeck("fill-text.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "textbox", name: "Dark", frame: { x: 1, y: 1, w: 3, h: 1 }, fill: "1F1F1F", text: [ { runs: [ { text: "a" }, { text: "b", color: "FF0000" } ] } ] },
            { type: "textbox", name: "Light", frame: { x: 5, y: 1, w: 3, h: 1 }, fill: "FFFF66", text: "c" },
            { type: "textbox", name: "None", frame: { x: 9, y: 1, w: 3, h: 1 }, text: "d" },
            { type: "shape", name: "Shape", shape: "rect", frame: { x: 1, y: 2.5, w: 3, h: 1 }, fill: "dk1", text: "e" },
            { type: "shape", name: "Own", shape: "rect", frame: { x: 5, y: 2.5, w: 3, h: 1 }, fill: "dk1", fontColor: "accent2", text: "f" },
            { type: "table", name: "Tab", frame: { x: 1, y: 4, w: 6 }, data: { headers: [ "H", { text: "I", fill: "FFFFFF" } ],
                rows: [ [ "1", { text: "2", fill: "000000" } ], [ "3", "4" ] ], style: { altRowBg: "222222" } } },
            { type: "table", name: "LightTab", frame: { x: 1, y: 6, w: 6 }, data: { headers: [ "J" ], rows: [ [ "5" ] ], style: { headerBg: "EEEEEE" } } }
        ] } ])
        const doc = await slideDoc(deck, 0)
        expect(runColors(shapeNamed(doc, "Dark"))).toEqual([ "lt1", "FF0000" ])
        expect(runColors(shapeNamed(doc, "Light"))).toEqual([ "dk1" ])
        expect(runColors(shapeNamed(doc, "None"))).toEqual([ "-" ])
        expect(runColors(shapeNamed(doc, "Shape"))).toEqual([ "lt1" ])
        expect(runColors(shapeNamed(doc, "Own"))).toEqual([ "accent2" ])

        /*  header: the default white, a light cell fill dark; body: alt rows and cell fills  */
        const cells = elements(shapeNamed(doc, "Tab"), "a:tc").map(runColors).map((c) => c[0])
        expect(cells).toEqual([ "FFFFFF", "dk1", "-", "lt1", "lt1", "lt1" ])
        expect(elements(shapeNamed(doc, "LightTab"), "a:tc").map(runColors).map((c) => c[0])).toEqual([ "dk1", "-" ])
    })
})

describe("shape presets", () => {
    it("writes every preset as its prstGeom and state reads it back", async () => {
        const deck = await freshDeck("shape-presets.pptx")
        const presets = ShapeTypeSchema.options.filter((p) => p !== "line")
        expect(presets.length).toBeGreaterThanOrEqual(45)
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: presets.map((shape, i) => ({
            type: "shape", name: `P${i}`, shape, frame: { x: 0.2 + (i % 10) * 1.3, y: 0.3 + Math.floor(i / 10) * 1.4, w: 1, h: 1 }, fill: "accent1"
        })) } ])
        const doc = await slideDoc(deck, 0)
        const shapes = (await stateOf(deck)).slides[0]?.shapes ?? []
        presets.forEach((shape, i) => {
            expect(firstElement(shapeNamed(doc, `P${i}`), "a:prstGeom")?.getAttribute("prst"), shape).toBe(shape)
            expect(shapes.find((sh) => sh.name.startsWith(`P${i}-`))?.shape, shape).toBe(shape)
        })
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [ { type: "shape", shape: "smileyFace",
            frame: { x: 1, y: 1, w: 1, h: 1 } } ] } ])).toBe("E_SCHEMA")
    })
})

describe("theme fonts", () => {
    it("writes +mj-lt and +mn-lt as the run's latin typeface", async () => {
        const deck = await freshDeck("theme-fonts.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: [ { text: "Titel", font: "+mj-lt" } ] } } },
            { op: "el.add", slide: "0", elements: [ { type: "textbox", name: "Box", frame: { x: 1, y: 5, w: 4, h: 1 },
                text: [ { runs: [ { text: "Kopf", font: "+mj-lt" }, { text: " Text", font: "+mn-lt" } ] } ] } ] } ])
        const doc = await slideDoc(deck, 0)
        const faces = (scope: Element): (string | null)[] => elements(scope, "a:latin").map((l) => l.getAttribute("typeface"))
        expect(faces(shapeNamed(doc, "Box"))).toEqual([ "+mj-lt", "+mn-lt" ])
        expect(faces(shapeNamed(doc, "PptcPh-0"))).toEqual([ "+mj-lt" ])
        expect((await stateShape(deck, "Box")).fontFace).toBe("+mj-lt")
    })
})
