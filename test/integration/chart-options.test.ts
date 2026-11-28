/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: chart labels, axes, legend, percent stacking, smoothing
**  and doughnut holes, written into the chart part; scatter series share
**  their x values; combinations a chart type does not support fail with
**  E_SCHEMA before anything is written.
*/

import { beforeAll, describe, expect, it } from "vitest"
import type { Element } from "@xmldom/xmldom"
import { elements, firstElement } from "../../src/engine/xml.js"
import { apply, attrs, chartOf, failure, freshDeck } from "../util/elements.js"

let deck: string

beforeAll(async () => {
    deck = await freshDeck("chart-options.pptx")
    await apply(deck, [ { op: "slide.add", layout: "DEFAULT" } ])
})

/**  add a chart and return its chart part  */
const chart = async (name: string, data: Record<string, unknown>): Promise<Element> => {
    await apply(deck, [ { op: "el.add", slide: "0", elements: [ { type: "chart", name, frame: { x: 1, y: 1, w: 5, h: 4 }, data } ] } ])
    return await chartOf(deck, name)
}

const bars = { type: "column", categories: [ "a", "b" ], series: [ { name: "s", values: [ 1, 2 ] }, { name: "t", values: [ 3, 1 ] } ] }

/**  the axis title texts of a chart part  */
const axisTitles = (doc: Element): string[] =>
    [ ...elements(doc, "c:catAx"), ...elements(doc, "c:valAx") ]
        .map((ax) => elements(firstElement(ax, "c:title") ?? ax, "a:t").map((t) => t.textContent).join(""))

describe("chart options", () => {
    it("writes value labels, their format and position, legend position and axes", async () => {
        const doc = await chart("Labelled", { ...bars, labels: { format: "0.0", position: "insideEnd" }, legend: "top",
            axes: { x: { title: "Quarter" }, y: { title: "Revenue", min: 0, max: 5, format: "#,##0 €" } } })
        expect(attrs(doc, "c:showVal", "val")).toContain("1")
        expect(attrs(doc, "c:dLblPos", "val")).toContain("inEnd")
        expect(attrs(doc, "c:numFmt", "formatCode")).toEqual(expect.arrayContaining([ "0.0", "#,##0 €" ]))
        expect(attrs(doc, "c:legendPos", "val")).toEqual([ "t" ])
        expect(axisTitles(doc)).toEqual([ "Quarter", "Revenue" ])
        const valAx = elements(doc, "c:valAx")[0] as Element
        expect([ attrs(valAx, "c:min", "val"), attrs(valAx, "c:max", "val") ]).toEqual([ [ "0" ], [ "5" ] ])
    })

    it("stacks to 100% with a percent axis", async () => {
        const doc = await chart("Percent", { ...bars, stacked: "percent", labels: { position: "center" } })
        expect(attrs(doc, "c:grouping", "val")).toEqual([ "percentStacked" ])
        expect(attrs(doc, "c:overlap", "val")).toEqual([ "100" ])
        expect(attrs(elements(doc, "c:valAx")[0] as Element, "c:numFmt", "formatCode")).toEqual([ "0%" ])
    })

    it("gives area charts their grouping: stacked, percent-stacked (area and areaStacked)", async () => {
        const area = { type: "area", categories: [ "a", "b" ], series: [ { name: "s", values: [ 1, 2 ] }, { name: "t", values: [ 3, 4 ] } ] }
        const grouping = (doc: Element): (string | null)[] =>
            elements(doc, "c:areaChart").map((a) => firstElement(a, "c:grouping")?.getAttribute("val") ?? null)
        const first = (doc: Element): string | undefined =>
            (elements(doc, "c:areaChart")[0]?.firstChild as Element | null)?.nodeName
        const percent = await chart("AreaPct", { ...area, stacked: "percent" })
        expect(grouping(percent)).toEqual([ "percentStacked" ])
        expect(first(percent)).toBe("c:grouping")
        expect(attrs(elements(percent, "c:valAx")[0] as Element, "c:numFmt", "formatCode")).toEqual([ "0%" ])
        expect(grouping(await chart("AreaStackedPct", { ...area, type: "areaStacked", stacked: "percent" }))).toEqual([ "percentStacked" ])
        expect(grouping(await chart("AreaStacked", { ...area, type: "areaStacked" }))).toEqual([ "stacked" ])
        expect(grouping(await chart("AreaPlain", area))).toEqual([ null ])
    })

    it("shows pie percentages, doughnut holes and smooth lines", async () => {
        const pie = await chart("Pie", { type: "doughnut", categories: [ "a", "b" ], series: [ { name: "s", values: [ 1, 3 ] } ],
            labels: { percent: true }, holeSize: 65, legend: "right" })
        expect(attrs(pie, "c:holeSize", "val")).toEqual([ "65" ])
        expect(new Set(attrs(firstElement(pie, "c:ser") as Element, "c:showPercent", "val"))).toEqual(new Set([ "1" ]))
        expect(attrs(pie, "c:legendPos", "val")).toEqual([ "r" ])
        const line = await chart("Smooth", { type: "line", categories: [ "a", "b" ], series: [ { name: "s", values: [ 1, 3 ] } ], smooth: true })
        expect(attrs(line, "c:smooth", "val")).toEqual([ "1" ])
    })

    it("writes scatter series with their shared x values", async () => {
        const doc = await chart("Points", { type: "scatter", series: [ { name: "s", x: [ 1, 2, 3 ], y: [ 2, 1, 3 ] },
            { name: "t", x: [ 1, 2, 3 ], y: [ 5, 6, 7 ] } ] })
        const values = (tag: string): string[][] => elements(doc, "c:ser")
            .map((ser) => elements(firstElement(ser, tag) ?? ser, "c:v").map((v) => v.textContent ?? ""))
        expect(values("c:xVal")).toEqual([ [ "1", "2", "3" ], [ "1", "2", "3" ] ])
        expect(values("c:yVal")).toEqual([ [ "2", "1", "3" ], [ "5", "6", "7" ] ])
    })

    it("gives the x axis of a scatter chart its own range and format", async () => {
        const doc = await chart("Scatter", { type: "scatter", series: [ { name: "s", x: [ 1, 2, 3 ], y: [ 2, 1, 3 ] } ],
            axes: { x: { min: 0, max: 4, format: "0.0" }, y: { format: "0%" } } })
        const [ x, y ] = elements(doc, "c:valAx")
        expect(attrs(x as Element, "c:numFmt", "formatCode")).toEqual([ "0.0" ])
        expect(attrs(x as Element, "c:max", "val")).toEqual([ "4" ])
        expect(attrs(y as Element, "c:numFmt", "formatCode")).toEqual([ "0%" ])
    })

    it("rejects options the chart type does not support", async () => {
        const reject = async (data: Record<string, unknown>): Promise<string | null> =>
            await failure(deck, [ { op: "el.add", slide: "0", elements: [ { type: "chart", frame: { x: 1, y: 1, w: 5, h: 4 }, data } ] } ])
        const pie = { type: "pie", categories: [ "a" ], series: [ { name: "s", values: [ 1 ] } ] }
        expect(await reject({ ...pie, axes: { y: { title: "x" } } })).toBe("E_SCHEMA")
        expect(await reject({ ...bars, labels: { percent: true } })).toBe("E_SCHEMA")
        expect(await reject({ ...bars, stacked: "percent", labels: { position: "outsideEnd" } })).toBe("E_SCHEMA")
        expect(await reject({ ...pie, type: "doughnut", labels: { position: "center" } })).toBe("E_SCHEMA")
        expect(await reject({ ...bars, axes: { x: { min: 1 } } })).toBe("E_SCHEMA")
        expect(await reject({ ...bars, axes: { y: { min: 5, max: 1 } } })).toBe("E_SCHEMA")
        expect(await reject({ ...pie, type: "line", stacked: "percent" })).toBe("E_SCHEMA")
        expect(await reject({ ...bars, smooth: true })).toBe("E_SCHEMA")
        expect(await reject({ ...pie, holeSize: 50 })).toBe("E_SCHEMA")
        expect(await reject({ type: "scatter", series: [ { name: "s", x: [ 1, 2 ], y: [ 1 ] } ] })).toBe("E_SCHEMA")
        expect(await reject({ type: "scatter", series: [ { name: "s", x: [ 1, 2 ], y: [ 1, 2 ] },
            { name: "t", x: [ 3, 4 ], y: [ 1, 2 ] } ] })).toBe("E_SCHEMA")
        expect(await reject({ type: "bubble", series: [ { name: "s", x: [ 1, 2 ], y: [ 1, 2 ] } ] })).toBe("E_SCHEMA")
    })
})
