/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: connector arrows and dashes. `startArrow` decorates the
**  `from` end (headEnd), `endArrow` the `to` end (tailEnd), also for lines
**  drawn right to left (flipped); `none` and an omitted end write no
**  decoration; `longDash` is `lgDash`. `state` reads the arrows back.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { firstElement } from "../../src/engine/xml.js"
import { apply, attrs, failure, freshDeck, shapeOf, stateOf } from "../util/elements.js"

let deck: string

beforeAll(async () => {
    deck = await freshDeck("connector-arrows.pptx")
    await apply(deck, [ { op: "slide.add", layout: "DEFAULT", ref: "s" } ])
})

describe("connector arrows", () => {
    it("writes both ends and the dash of a right-to-left line", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "connector", name: "Back", from: [ 6, 2 ], to: [ 2, 3 ],
                startArrow: "oval", endArrow: "triangle", dash: "longDash" } ] } ])
        const { shape } = await shapeOf(deck, "Back")
        expect(attrs(shape, "a:headEnd", "type")).toEqual([ "oval" ])
        expect(attrs(shape, "a:tailEnd", "type")).toEqual([ "triangle" ])
        expect(attrs(shape, "a:prstDash", "val")).toEqual([ "lgDash" ])
        expect(firstElement(shape, "a:xfrm")?.getAttribute("flipH")).toBe("1")
        const read = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name.startsWith("Back-"))
        expect([ read?.startArrow, read?.endArrow ]).toEqual([ "oval", "triangle" ])
    })

    it("writes no decoration for none or an omitted end", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "connector", name: "Plain", from: [ 1, 1 ], to: [ 3, 1 ], startArrow: "none", dash: "dashDot" } ] } ])
        const { shape } = await shapeOf(deck, "Plain")
        expect(attrs(shape, "a:headEnd", "type")).toEqual([])
        expect(attrs(shape, "a:tailEnd", "type")).toEqual([])
        expect(attrs(shape, "a:prstDash", "val")).toEqual([ "dashDot" ])
    })

    it("rejects an unknown arrow type", async () => {
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "connector", from: [ 1, 1 ], to: [ 3, 1 ], endArrow: "curly" } ] } ])).toBe("E_SCHEMA")
    })
})
