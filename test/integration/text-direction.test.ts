/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: text direction and rotated geometry in the lints. A
**  textbox or shape writes `bodyPr@vert` (vert, vert270, wordArtVert) and
**  `state` reports the same names; the capacity of a placeholder whose
**  text runs vertically (vert/vert270) swaps width and height; the overlap
**  lint takes the axis-aligned box a rotated element covers.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import { DeckArchive, readTemplateInfo } from "../../src/engine/reader.js"
import { estimateCapacity } from "../../src/core/describe/capacity.js"
import { elements, firstElement, parseXml, serializeXml } from "../../src/engine/xml.js"
import { TEMPLATE, TMP, apply, freshDeck, shapeOf, stateOf } from "../util/elements.js"

let deck: string

beforeAll(async () => {
    deck = await freshDeck("text-direction.pptx")
    await apply(deck, [ { op: "slide.add", layout: "Title Only" }, { op: "slide.add", layout: "Title Only" } ])
})

describe("text direction", () => {
    it("writes and reads the direction of textboxes and shapes", async () => {
        await apply(deck, [ { op: "el.add", slide: "1", elements: [
            { type: "textbox", name: "Down", frame: { x: 0.5, y: 2, w: 1, h: 4 }, text: "down", textDirection: "rotate90" },
            { type: "shape", name: "Up", shape: "rect", frame: { x: 2, y: 2, w: 1, h: 4 }, text: "up", textDirection: "rotate270" },
            { type: "textbox", name: "Stack", frame: { x: 3.5, y: 2, w: 1, h: 4 }, text: "abc", textDirection: "stacked" },
            { type: "textbox", name: "Flat", frame: { x: 5, y: 2, w: 3, h: 1 }, text: "flat", textDirection: "horizontal" }
        ] } ])
        const vert = async (name: string): Promise<string | null | undefined> =>
            firstElement((await shapeOf(deck, name)).shape, "a:bodyPr")?.getAttribute("vert")
        expect([ await vert("Down"), await vert("Up"), await vert("Stack"), await vert("Flat") ])
            .toEqual([ "vert", "vert270", "wordArtVert", "horz" ])
        const read = Object.fromEntries((await stateOf(deck)).slides[1]?.shapes
            .map((sh) => [ sh.name.replace(/-.*$/, ""), sh.textDirection ]) ?? [])
        expect([ read["Down"], read["Up"], read["Stack"], read["Flat"] ]).toEqual([ "rotate90", "rotate270", "stacked", undefined ])
    })

    it("lints the overlap of a rotated element by the box it covers", async () => {
        /*  0.2 in high below the title -- turned by 90 degrees it reaches up into it  */
        const box = { type: "shape", shape: "rect", frame: { x: 4, y: 2.5, w: 3, h: 0.2 } }
        const flat = await apply(deck, [ { op: "el.add", slide: "0", elements: [ { ...box, name: "Lying" } ] } ])
        expect(flat).toEqual([])
        const turned = await apply(deck, [ { op: "el.add", slide: "0", elements: [
            { ...box, name: "Standing", frame: { ...box.frame, rotation: 90 } } ] } ])
        expect(turned).toEqual([ expect.objectContaining({ code: "W_ELEMENT_OVERLAP", element: "Standing" }) ])
    })
})

describe("vertical placeholder capacity", () => {
    it("swaps width and height for text rotated by 90 or 270 degrees", async () => {
        const zip = await JSZip.loadAsync(readFileSync(TEMPLATE))
        const info = await readTemplateInfo(await DeckArchive.open(TEMPLATE))
        const content = info.layouts.find((l) => l.name === "CONTENT")
        const body = content?.placeholders.find((p) => p.kind === "body")
        const part = (await Promise.all(Object.keys(zip.files).filter((f) => /slideLayouts\/slideLayout\d+\.xml$/.test(f))
            .map(async (f) => [ f, await (zip.file(f) as JSZip.JSZipObject).async("string") ] as const)))
            .find(([ , xml ]) => xml.includes("name=\"CONTENT\"")) as readonly [ string, string ]
        const doc = parseXml(part[1])
        const sp = elements(doc, "p:sp").find((el) => firstElement(el, "p:ph")?.getAttribute("idx") === "1")
        const bodyPr = sp === undefined ? null : firstElement(sp, "a:bodyPr")
        bodyPr?.setAttribute("vert", "vert270")
        zip.file(part[0], serializeXml(doc))
        const vertical = path.join(TMP, "vertical-template.pptx")
        writeFileSync(vertical, await zip.generateAsync({ type: "nodebuffer" }))
        const swapped = (await readTemplateInfo(await DeckArchive.open(vertical))).layouts
            .find((l) => l.name === "CONTENT")?.placeholders.find((p) => p.kind === "body")
        const frame = body?.frame as { x: number, y: number, w: number, h: number }
        expect(bodyPr).not.toBeNull()
        expect(swapped?.frame).toEqual(frame)
        expect(swapped?.capacity).toEqual(estimateCapacity({ ...frame, w: frame.h, h: frame.w }, body?.capacity?.fontSizePt ?? 0))
        expect(swapped?.capacity).not.toEqual(body?.capacity)
    })
})
