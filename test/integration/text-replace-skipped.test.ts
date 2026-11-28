/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: text.replace leaves chart and SmartArt text as it is and
**  says so -- W_REPLACE_SKIPPED for each targeted slide whose chart or
**  SmartArt part holds a match.
*/

import { describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import { applyResult, freshDeck, stateOf } from "../util/elements.js"

type Warning = { code: string, slide: { index: number }, message: string }

/**  the W_REPLACE_SKIPPED warnings of an apply  */
const skipped = async (deck: string, ops: unknown[]): Promise<Warning[]> =>
    ((await applyResult(deck, ops)).warnings as Warning[]).filter((w) => w.code === "W_REPLACE_SKIPPED")

/**  give a slide a SmartArt data part (content and relationship only)  */
const addSmartArt = async (deck: string, index: number, text: string): Promise<void> => {
    const part = ((await stateOf(deck)).slides[index] as { part: string }).part
    const zip = await JSZip.loadAsync(readFileSync(deck))
    zip.file("ppt/diagrams/data91.xml", "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\""
        + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><dgm:ptLst><dgm:pt modelId=\"1\"><dgm:t><a:bodyPr/>"
        + `<a:p><a:r><a:t>${text}</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>`)
    const ct = await (zip.file("[Content_Types].xml") as JSZip.JSZipObject).async("string")
    zip.file("[Content_Types].xml", ct.replace("</Types>", "<Override PartName=\"/ppt/diagrams/data91.xml\""
        + " ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml\"/></Types>"))
    const relsPart = `ppt/slides/_rels/${path.posix.basename(part)}.rels`
    const rels = await (zip.file(relsPart) as JSZip.JSZipObject).async("string")
    zip.file(relsPart, rels.replace("</Relationships>", "<Relationship Id=\"rIdDgm91\""
        + " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData\""
        + " Target=\"../diagrams/data91.xml\"/></Relationships>"))
    writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
}

describe("text.replace and chart or SmartArt text", () => {
    it("warns when a targeted slide's chart or SmartArt holds the search text", async () => {
        const deck = await freshDeck("replace-skipped.pptx")
        await applyResult(deck, [
            { op: "slide.add", layout: "Title Only", ref: "a", placeholders: { title: { text: "Umsatz alt" } } },
            { op: "slide.add", layout: "Title Only", ref: "b", placeholders: { title: { text: "Zweite" } } },
            { op: "el.add", slide: "$a", elements: [ { type: "chart", name: "Chart", frame: { x: 1, y: 2, w: 5, h: 3 },
                data: { type: "column", title: "Umsatz alt", categories: [ "Q1" ], series: [ { name: "alt", values: [ 1 ] } ] } } ] }
        ])
        await addSmartArt(deck, 1, "Schritt alt")

        const all = await skipped(deck, [ { op: "text.replace", find: "alt", replace: "neu" } ])
        expect(all.map((w) => [ w.slide.index, /leaves (\w+) text/.exec(w.message)?.[1] ])).toEqual([ [ 0, "chart" ], [ 1, "SmartArt" ] ])
        expect(all[0]?.message).toMatch(/text\.replace \(op 0\).*'alt'.*ppt\/charts\/chart\d+\.xml/)
        expect((await stateOf(deck)).slides[0]?.title).toBe("Umsatz neu")

        /*  only the targeted slide; no match, no warning  */
        expect((await skipped(deck, [ { op: "text.replace", find: "alt", replace: "neu", slide: "1" } ])).map((w) => w.slide.index)).toEqual([ 1 ])
        expect(await skipped(deck, [ { op: "text.replace", find: "fehlt", replace: "x" } ])).toEqual([])
        expect(await skipped(deck, [ { op: "text.replace", find: "ALT", replace: "x", matchCase: true } ])).toEqual([])

        /*  a number of the chart's value cache is no chart text  */
        await applyResult(deck, [ { op: "slide.fill", slide: "0", placeholders: { title: { text: "Wert 1234" } } },
            { op: "el.add", slide: "0", elements: [ { type: "chart", name: "Numbers", frame: { x: 6, y: 2, w: 3, h: 3 },
                data: { type: "column", categories: [ "Q1", "1999" ], series: [ { name: "s", values: [ 1234, 7 ] } ] } } ] } ])
        expect(await skipped(deck, [ { op: "text.replace", find: "1234", replace: "5678", slide: "0" } ])).toEqual([])
        expect((await stateOf(deck)).slides[0]?.title).toBe("Wert 5678")
        /*  a category (string cache) still counts  */
        expect((await skipped(deck, [ { op: "text.replace", find: "1999", replace: "2000", slide: "0" } ])).length).toBe(1)

        /*  a chart an earlier op removes is no longer shown  */
        expect(await skipped(deck, [ { op: "el.rm", slide: "0", name: "Chart" },
            { op: "text.replace", find: "alt", replace: "neu", slide: "0" } ])).toEqual([])
    })
})
