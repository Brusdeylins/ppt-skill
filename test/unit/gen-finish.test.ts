/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: finishing the PptxGenJS output -- theme colors go in as hex
**  tokens no element uses literally and come out as `a:schemeClr` (with
**  their color modifiers); a paragraph keeps its first `a:pPr` only.
*/

import { describe, expect, it } from "vitest"
import { GenColors, finishGenerated } from "../../src/engine/elements.js"
import { SCHEME_COLORS } from "../../src/schema/payloads.js"
import { elements, parseXml, serializeXml } from "../../src/engine/xml.js"

const A = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""

describe("theme color tokens", () => {
    it("passes hex through and gives every theme slot one token", () => {
        const colors = new GenColors([])
        expect(colors.gen("1F4E79")).toBe("1F4E79")
        const accent = colors.gen("accent1")
        expect(accent).toMatch(/^[0-9A-F]{6}$/)
        expect(colors.gen("accent1")).toBe(accent)
        expect(colors.gen("dk2")).not.toBe(accent)
        expect([ ...colors.tokens.values() ]).toEqual([ "accent1", "dk2" ])
    })

    it("gives every theme slot (hlink and folHlink too) a token distinct from the inherit token", () => {
        const colors = new GenColors([])
        const tokens = SCHEME_COLORS.map((slot) => colors.gen(slot))
        expect(new Set([ ...tokens, colors.inherit ]).size).toBe(SCHEME_COLORS.length + 1)
        expect(SCHEME_COLORS).toEqual(expect.arrayContaining([ "hlink", "folHlink" ]))
        /*  a literal collision moves the whole range, still distinct  */
        const moved = new GenColors([ { type: "textbox", frame: { x: 0, y: 0 }, text: "x", fill: colors.inherit } ])
        const all = [ ...SCHEME_COLORS.map((slot) => moved.gen(slot)), moved.inherit ]
        expect(new Set(all).size).toBe(SCHEME_COLORS.length + 1)
        expect(all).not.toContain(colors.inherit)
    })

    it("avoids tokens an element uses literally", () => {
        const plain = new GenColors([]).gen("accent1")
        const colors = new GenColors([ { type: "textbox", frame: { x: 0, y: 0 }, text: "x", fill: plain } ])
        expect(colors.gen("accent1")).not.toBe(plain)
    })

    it("turns tokens into schemeClr and keeps one pPr per paragraph", () => {
        const colors = new GenColors([])
        const token = colors.gen("lt2")
        const doc = parseXml(`<a:x ${A}><a:solidFill><a:srgbClr val="${token}"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`
            + "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>"
            + "<a:p><a:pPr algn=\"ctr\"/><a:r><a:t>a</a:t></a:r><a:pPr/><a:r><a:t>b</a:t></a:r></a:p></a:x>")
        expect(finishGenerated(doc, colors.tokens)).toBe(true)
        const xml = serializeXml(doc)
        expect(xml).toContain("<a:schemeClr val=\"lt2\"><a:alpha val=\"50000\"/></a:schemeClr>")
        expect(xml).toContain("<a:srgbClr val=\"FF0000\"/>")
        expect(elements(doc, "a:pPr").map((p) => p.getAttribute("algn"))).toEqual([ "ctr" ])
        expect(finishGenerated(doc, colors.tokens)).toBe(false)
    })
})
