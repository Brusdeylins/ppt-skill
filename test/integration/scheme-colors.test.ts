/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: theme colors. Every color field takes a bare theme slot
**  (accent1-accent6, tx1, tx2, bg1, bg2, dk1, lt1, dk2, lt2, hlink,
**  folHlink) besides hex
**  and writes it as `a:schemeClr` -- in el.add elements (also the slots
**  PptxGenJS does not know, and chart series colors), in placeholder runs,
**  in the background; a hex color that equals a color token stays hex;
**  `state` resolves theme colors to RRGGBB.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import JSZip from "jszip"
import { elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { apply, attrs, chartOf, failure, freshDeck, shapeOf, stateOf } from "../util/elements.js"

let deck: string

beforeAll(async () => {
    deck = await freshDeck("scheme-colors.pptx")
})

describe("theme colors", () => {
    it("writes theme slots of el.add elements as schemeClr", async () => {
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [
            { type: "shape", name: "Card", shape: "rect", frame: { x: 1, y: 1, w: 3, h: 1 },
                fill: "accent2", border: "dk2", fontColor: "lt1", text: "card" },
            { type: "textbox", name: "Note", frame: { x: 1, y: 2.5, w: 3, h: 1 },
                text: [ { runs: [ { text: "a", color: "tx2", highlight: "bg2" } ] } ] },
            { type: "connector", name: "Line", from: [ 5, 1 ], to: [ 8, 1 ], color: "accent6" },
            { type: "table", name: "Tab", frame: { x: 1, y: 4, w: 4 },
                data: { headers: [ "H" ], rows: [ [ { text: "c", fill: "lt2" } ] ], style: { headerBg: "accent1", headerFg: "bg1" } } },
            { type: "chart", name: "Bars", frame: { x: 5, y: 2, w: 4, h: 3 }, data: { type: "column",
                categories: [ "a", "b" ], series: [ { name: "s", values: [ 1, 2 ] }, { name: "t", values: [ 2, 1 ] } ],
                colors: [ "accent3", "0C0FE0" ] } }
        ] } ])
        const scheme = async (name: string): Promise<(string | null)[]> => attrs((await shapeOf(deck, name)).shape, "a:schemeClr", "val")
        expect(await scheme("Card")).toEqual(expect.arrayContaining([ "accent2", "dk2", "lt1" ]))
        expect(await scheme("Note")).toEqual([ "tx2", "bg2" ])
        expect(await scheme("Line")).toEqual([ "accent6" ])
        expect(await scheme("Tab")).toEqual(expect.arrayContaining([ "accent1", "bg1", "lt2" ]))
        const chart = await chartOf(deck, "Bars")
        expect(attrs(chart, "a:schemeClr", "val")).toContain("accent3")
        /*  a hex value that happens to look like a token stays hex  */
        expect(attrs(chart, "a:srgbClr", "val")).toContain("0C0FE0")
        const card = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Card-"))
        expect(card?.fill).toMatch(/^[0-9A-F]{6}$/)
    })

    it("writes theme slots of placeholder runs and the background", async () => {
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", background: { color: "accent5" },
            placeholders: { title: { text: [ { runs: [ { text: "Title", color: "accent1", highlight: "lt2" } ] } ] },
                body: { text: [ { text: "body", color: "#1F4E79" } ] } } } ])
        const state = await stateOf(deck)
        const part = (state.slides.at(-1) as { part: string }).part
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const doc = parseXml(await (zip.file(part) as JSZip.JSZipObject).async("string"))
        expect(firstElement(firstElement(doc, "p:bg") ?? doc.documentElement as never, "a:schemeClr")?.getAttribute("val")).toBe("accent5")
        const runs = elements(doc, "a:rPr").map((rPr) => [ ...attrs(rPr, "a:schemeClr", "val"), ...attrs(rPr, "a:srgbClr", "val") ])
        expect(runs.filter((r) => r.length > 0)).toEqual([ [ "accent1", "lt2" ], [ "1F4E79" ] ])
    })

    it("writes the hyperlink slots hlink and folHlink, distinct from inheriting text", async () => {
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "$last", elements: [
            { type: "shape", name: "Links", shape: "rect", frame: { x: 1, y: 1, w: 3, h: 1 }, fill: "hlink", border: "folHlink",
                text: [ { runs: [ { text: "a", color: "accent1" }, { text: "b" } ] } ] },
            /*  every slot in one apply: the token range holds them all  */
            { type: "textbox", name: "All", frame: { x: 1, y: 3, w: 6, h: 1 }, text: [ { runs: [
                "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "tx1", "tx2", "bg1", "bg2",
                "dk1", "lt1", "dk2", "lt2", "hlink", "folHlink" ].map((color) => ({ text: color, color })) } ] }
        ] } ].map((op, i) => i === 0 ? { ...op, ref: "last" } : op))
        const links = (await shapeOf(deck, "Links")).shape
        expect(attrs(firstElement(links, "p:spPr") as never, "a:schemeClr", "val")).toEqual([ "hlink", "folHlink" ])
        expect(attrs(links, "a:srgbClr", "val")).toEqual([])
        const all = (await shapeOf(deck, "All")).shape
        expect(attrs(all, "a:schemeClr", "val")).toEqual([ "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
            "tx1", "tx2", "bg1", "bg2", "dk1", "lt1", "dk2", "lt2", "hlink", "folHlink" ])
        expect(elements(all, "a:solidFill").length).toBe(16)
        expect(attrs(all, "a:srgbClr", "val")).toEqual([])
    })

    it("rejects an unknown color name", async () => {
        expect(await failure(deck, [ { op: "slide.fill", slide: "0", background: { color: "accent7" } } ])).toBe("E_SCHEMA")
    })
})
