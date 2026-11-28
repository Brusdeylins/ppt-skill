/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the stacking order. el.add places a new element in front
**  (default) or where its `zOrder` says, el.set moves an existing or
**  planned element to the front or back or one step; the moves follow the
**  op order, `state` lists the shapes back to front, an element sent to
**  the back covers nothing (no overlap warning), the content guard
**  accepts the reorder and untouched slides keep their bytes.
*/

import { describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import { TMP, apply, failure, freshDeck, png, stateOf } from "../util/elements.js"

/**  the names of a slide's elements and title back to front, generated
     suffixes cut (the footer band placeholders left out)  */
const order = async (deck: string, index = 0): Promise<string[]> =>
    ((await stateOf(deck)).slides[index]?.shapes ?? []).filter((sh) => sh.placeholderIdx === null || sh.placeholderKind === "title")
        .map((sh) => sh.name.replace(/-[0-9a-f-]{36}$/, ""))

const box = (name: string, x: number, extra: Record<string, unknown> = {}): Record<string, unknown> =>
    ({ type: "shape", name, shape: "rect", frame: { x, y: 4, w: 1, h: 1 }, fill: "accent1", ...extra })

describe("z-order", () => {
    it("places el.add elements in front or where zOrder says", async () => {
        const deck = await freshDeck("z-add.pptx")
        await apply(deck, [ { op: "slide.add", layout: "Title Only", placeholders: { title: { text: "Z" } } },
            { op: "el.add", slide: "0", elements: [ box("A", 1), box("B", 2), box("C", 3, { zOrder: "back" }),
                box("D", 4, { zOrder: "backward" }), box("E", 5, { zOrder: "forward" }) ] } ])
        expect(await order(deck)).toEqual([ "C", "PptcPh-0", "A", "D", "B", "E" ])
    })

    it("moves existing elements front, back, forward and backward in op order", async () => {
        const deck = await freshDeck("z-set.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" },
            { op: "el.add", slide: "0", elements: [ box("A", 1), box("B", 2), box("C", 3), box("D", 4) ] } ])
        expect(await order(deck)).toEqual([ "A", "B", "C", "D" ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "D", zOrder: "back" } ])
        expect(await order(deck)).toEqual([ "D", "A", "B", "C" ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "D", zOrder: "front" },
            { op: "el.set", slide: "0", name: "A", zOrder: "forward" } ])
        expect(await order(deck)).toEqual([ "B", "A", "C", "D" ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "C", zOrder: "backward" },
            { op: "el.set", slide: "0", name: "C", zOrder: "backward" } ])
        expect(await order(deck)).toEqual([ "C", "B", "A", "D" ])

        /*  at the edge a step changes nothing  */
        await apply(deck, [ { op: "el.set", slide: "0", name: "C", zOrder: "backward" },
            { op: "el.set", slide: "0", name: "D", zOrder: "forward" } ])
        expect(await order(deck)).toEqual([ "C", "B", "A", "D" ])
    })

    it("follows the op order between new and existing elements and edits planned ones", async () => {
        const deck = await freshDeck("z-ops.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" },
            { op: "el.add", slide: "0", elements: [ box("A", 1), box("B", 2) ] } ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "A", zOrder: "front" },
            { op: "el.add", slide: "0", elements: [ box("X", 3) ] } ])
        expect(await order(deck)).toEqual([ "B", "A", "X" ])
        await apply(deck, [ { op: "el.add", slide: "0", elements: [ box("Y", 4) ] },
            { op: "el.set", slide: "0", name: "B", zOrder: "front" },
            { op: "el.add", slide: "0", elements: [ box("Z", 5) ] },
            { op: "el.set", slide: "0", name: "Z", zOrder: "back" } ])
        expect(await order(deck)).toEqual([ "Z", "A", "X", "Y", "B" ])

        /*  a copy keeps the order its planned elements got  */
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT", ref: "n" },
            { op: "el.add", slide: "$n", elements: [ box("P", 1), box("Q", 2) ] },
            { op: "el.set", slide: "$n", name: "Q", zOrder: "back" },
            { op: "slide.copy", slide: "$n" } ])
        expect([ await order(deck, 1), await order(deck, 2) ]).toEqual([ [ "Q", "P" ], [ "Q", "P" ] ])
    })

    it("lints an element sent to the back as covering nothing", async () => {
        const deck = await freshDeck("z-lint.pptx")
        await apply(deck, [ { op: "slide.add", layout: "Title Only", placeholders: { title: { text: "Titel" } } } ])
        const panel = { type: "shape", name: "Panel", shape: "rect", frame: { x: 0, y: 0, w: 13, h: 2 }, fill: "lt2" }
        const codes = (warnings: unknown[]): string[] => (warnings as { code: string }[]).map((w) => w.code)
        expect(codes(await apply(deck, [ { op: "el.add", slide: "0", elements: [ { ...panel, zOrder: "back" } ] } ]))).not.toContain("W_ELEMENT_OVERLAP")
        expect(codes(await apply(deck, [ { op: "el.add", slide: "0", elements: [ { ...panel, name: "Over" } ] } ]))).toContain("W_ELEMENT_OVERLAP")
        expect((await order(deck))[0]).toBe("Panel")
    })

    it("rejects wrong values and types at plan time", async () => {
        const deck = await freshDeck("z-schema.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [ box("A", 1) ] } ])
        const before = readFileSync(deck)
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "A", zOrder: "top" } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "A", zOrder: 1 } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [ box("B", 2, { zOrder: "bottom" }) ] } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "Missing", zOrder: "back" } ])).toBe("E_ADDR_NOTFOUND")
        expect(readFileSync(deck).equals(before)).toBe(true)
    })

    it("passes the content guard with pictures and charts and keeps other slides' bytes", async () => {
        const deck = await freshDeck("z-guard.pptx")
        const img = path.join(TMP, "z-guard.png")
        writeFileSync(img, png(40, 20))
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "slide.add", layout: "DEFAULT" },
            { op: "el.add", slide: "0", elements: [
                { type: "image", name: "Pic", path: img, frame: { x: 1, y: 1, w: 2, h: 1 } },
                { type: "chart", name: "Chart", frame: { x: 4, y: 1, w: 4, h: 3 },
                    data: { type: "column", categories: [ "a" ], series: [ { name: "s", values: [ 1 ] } ] } },
                box("Box", 9) ] },
            { op: "el.add", slide: "1", elements: [ box("Other", 1) ] } ])
        const parts = (await stateOf(deck)).slides.map((s) => s.part)
        const bytesOf = async (part: string): Promise<Buffer> =>
            await ((await JSZip.loadAsync(readFileSync(deck))).file(part) as JSZip.JSZipObject).async("nodebuffer")
        const other = await bytesOf(parts[1] as string)
        await apply(deck, [ { op: "el.set", slide: "0", name: "Box", zOrder: "back" },
            { op: "el.set", slide: "0", name: "Pic", zOrder: "front" } ])
        expect(await order(deck)).toEqual([ "Box", "Chart", "Pic" ])
        expect((await bytesOf(parts[1] as string)).equals(other)).toBe(true)
        const shapes = (await stateOf(deck)).slides[0]?.shapes ?? []
        expect(shapes.filter((sh) => sh.placeholderIdx === null).map((sh) => sh.type)).toEqual([ "shape", "chart", "picture" ])
        expect(shapes[0]?.name).toMatch(/^Box-/)
    })
})
