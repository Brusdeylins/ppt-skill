/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the stacking order. el.add places a new element in front
**  (default) or where its `zOrder` says, el.set moves an existing or
**  planned element to the front or back or one step; the moves follow the
**  op order (a step counts the shapes there at that op), `state` lists the
**  shapes back to front, an element sent to the back covers nothing and
**  is no obstacle of later ones (no overlap warning), nor is a panel
**  without text (also in a later apply, under --strict), the content guard
**  accepts the reorder and untouched slides keep their bytes.
*/

import { describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import { executeOps } from "../../src/commands/apply.js"
import { OPTS, TMP, apply, failure, freshDeck, png, stateOf } from "../util/elements.js"

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

    it("counts only the shapes present at an el.add's point for forward and backward", async () => {
        const deck = await freshDeck("z-step.pptx")
        await apply(deck, [ { op: "slide.add", layout: "DEFAULT" }, { op: "el.add", slide: "0", elements: [ box("P", 1) ] } ])
        await apply(deck, [ { op: "el.add", slide: "0", elements: [ box("X", 2, { zOrder: "backward" }) ] },
            { op: "el.add", slide: "0", elements: [ box("Y", 3) ] },
            { op: "el.add", slide: "0", elements: [ box("V", 4, { zOrder: "forward" }) ] },
            { op: "el.add", slide: "0", elements: [ box("W", 5) ] } ])
        expect(await order(deck)).toEqual([ "X", "P", "Y", "V", "W" ])
    })

    it("leaves an element placed or moved to the back out as an obstacle of later elements", async () => {
        const deck = await freshDeck("z-panel.pptx")
        await apply(deck, [ { op: "slide.add", layout: "Title Only", placeholders: { title: { text: "Titel" } } } ])
        const codes = (warnings: unknown[]): string[] => (warnings as { code: string }[]).map((w) => w.code)
        const text = (name: string, x: number, y: number): Record<string, unknown> =>
            ({ type: "textbox", name, frame: { x: x + 0.2, y: y + 0.2, w: 1.5, h: 0.6 }, text: name })
        const panel = (name: string, x: number, y: number, extra: Record<string, unknown> = {}): Record<string, unknown> =>
            ({ type: "shape", name, shape: "rect", frame: { x, y, w: 2.5, h: 1.5 }, fill: "lt2", ...extra })

        /*  planned at the back in separate ops and in one op  */
        expect(codes(await apply(deck, [ { op: "el.add", slide: "0", elements: [ panel("BG", 0.5, 2, { zOrder: "back" }) ] },
            { op: "el.add", slide: "0", elements: [ text("Q", 0.5, 2) ] } ]))).not.toContain("W_ELEMENT_OVERLAP")
        expect(codes(await apply(deck, [ { op: "el.add", slide: "0",
            elements: [ panel("BG2", 3.5, 2, { zOrder: "back" }), text("R", 3.5, 2) ] } ]))).not.toContain("W_ELEMENT_OVERLAP")

        /*  text still is an obstacle  */
        expect(codes(await apply(deck, [ { op: "el.add", slide: "0", elements: [ text("S", 0.6, 2.1) ] } ]))).toContain("W_ELEMENT_OVERLAP")

        /*  an existing element moved to the back, a planned one moved  */
        await apply(deck, [ { op: "el.add", slide: "0", elements: [ panel("Card", 6.5, 2) ] } ])
        expect(codes(await apply(deck, [ { op: "el.set", slide: "0", name: "Card", zOrder: "back" },
            { op: "el.add", slide: "0", elements: [ text("T", 6.5, 2) ] } ]))).not.toContain("W_ELEMENT_OVERLAP")
        expect(codes(await apply(deck, [ { op: "el.add", slide: "0", elements: [ panel("Mid", 0.5, 4) ] },
            { op: "el.set", slide: "0", name: "Mid", zOrder: "back" },
            { op: "el.add", slide: "0", elements: [ text("U", 0.5, 4) ] } ]))).not.toContain("W_ELEMENT_OVERLAP")

        /*  brought to the front again with text, it counts again (a
            panel without text never does)  */
        expect(codes(await apply(deck, [ { op: "el.set", slide: "0", name: "Card", zOrder: "front", text: "Karte" },
            { op: "el.add", slide: "0", elements: [ text("V", 7.5, 2.7) ] } ]))).toContain("W_ELEMENT_OVERLAP")
    })

    it("leaves a panel without text out as an obstacle, in later applies and under --strict", async () => {
        const deck = await freshDeck("z-textless.pptx")
        const codes = (warnings: unknown[]): string[] => (warnings as { code: string }[]).map((w) => w.code)
        const strict = async (ops: unknown[]): Promise<unknown> =>
            await executeOps(deck, { ops } as never, { ...OPTS, strict: true })
        const text = (name: string, x: number, y: number): Record<string, unknown> =>
            ({ type: "textbox", name, frame: { x, y, w: 3, h: 0.8 }, text: name })
        await apply(deck, [ { op: "slide.add", layout: "Title Only", placeholders: { title: { text: "Titel" } } },
            { op: "el.add", slide: "0", elements: [ { type: "shape", name: "Panel", shape: "roundRect",
                frame: { x: 0.5, y: 2, w: 12, h: 4.5 }, fill: "lt2", transparency: 50, zOrder: "back" }, text("T1", 1, 2.5) ] } ])

        /*  the panel from the first apply covers nothing in the next ones  */
        await expect(strict([ { op: "el.add", slide: "0", elements: [ text("T2", 6, 4.5) ] } ])).resolves.toBeDefined()
        await expect(strict([ { op: "el.set", slide: "0", name: "T1", frame: { y: 3.2 } } ])).resolves.toBeDefined()

        /*  a planned shape without text or picture, in front, is no
            obstacle either, as in a later apply  */
        const photo = path.join(TMP, "z-textless.png")
        writeFileSync(photo, png(40, 30))
        expect(codes(await apply(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "shape", name: "Card", shape: "rect", frame: { x: 6, y: 2.2, w: 3, h: 1.5 }, fill: "accent1" },
            { type: "image", name: "Foto", frame: { x: 9.5, y: 2.2, w: 2.5, h: 1.5 }, path: photo } ] },
            { op: "el.add", slide: "0", elements: [ text("T3", 6.2, 2.4), text("T5", 9.7, 2.4) ] } ]))).not.toContain("W_ELEMENT_OVERLAP")

        /*  a panel that gets text is one  */
        expect(await apply(deck, [ { op: "el.set", slide: "0", name: "Panel", text: "Hinweis", zOrder: "front" },
            { op: "el.add", slide: "0", elements: [ text("T4", 1, 5.5) ] } ])).toEqual([
            expect.objectContaining({ code: "W_ELEMENT_OVERLAP", element: "T4", covers: expect.stringMatching(/^Panel-/) }) ])
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
