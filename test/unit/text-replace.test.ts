/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the plain-text replace of text.replace over runs -- matches
**  across run borders go into the first touched run, emptied runs go,
**  case and whole-word matching, rich text with its formatting kept.
*/

import { describe, expect, it } from "vitest"
import { replacePlain, replaceRichText, replaceRuns, type ReplaceSpec } from "../../src/core/text-replace.js"

const spec = (find: string, replace: string, extra: Partial<ReplaceSpec> = {}): ReplaceSpec =>
    ({ find, replace, matchCase: false, wholeWord: false, ...extra })

describe("replaceRuns", () => {
    it("leaves the runs as they are without a match", () => {
        const texts = [ "ab", "cd" ]
        expect(replaceRuns(texts, spec("x", "y"))).toEqual({ texts, count: 0 })
    })

    it("puts the replacement into the first touched run and empties the others", () => {
        expect(replaceRuns([ "Hal", "lo We", "lt!" ], spec("Hallo Welt", "Servus"))).toEqual({ texts: [ "Servus", null, "!" ], count: 1 })
        expect(replaceRuns([ "xHa", "llo" ], spec("Hallo", "Hi"))).toEqual({ texts: [ "xHi", null ], count: 1 })
        expect(replaceRuns([ "a", "b", "c" ], spec("b", ""))).toEqual({ texts: [ "a", null, "c" ], count: 1 })
    })

    it("replaces every match, also several in one run", () => {
        expect(replaceRuns([ "aXa", "Xa" ], spec("a", "b"))).toEqual({ texts: [ "bXb", "Xb" ], count: 3 })
        expect(replaceRuns([ "aa" ], spec("a", "aa"))).toEqual({ texts: [ "aaaa" ], count: 2 })
    })

    it("matches case and whole words on request", () => {
        expect(replaceRuns([ "Welt welt" ], spec("welt", "Erde", { matchCase: true })).texts).toEqual([ "Welt Erde" ])
        expect(replaceRuns([ "Welt Weltall welt_x" ], spec("welt", "Erde", { wholeWord: true })).texts).toEqual([ "Erde Weltall welt_x" ])
        expect(replaceRuns([ "Straße STRASSE" ], spec("straße", "Weg")).texts).toEqual([ "Weg STRASSE" ])
        expect(replaceRuns([ "a.b a+b" ], spec("a.b", "x")).texts).toEqual([ "x a+b" ])
    })
})

describe("replacePlain and replaceRichText", () => {
    it("never matches across a line", () => {
        expect(replacePlain("ab\ncd", spec("b\nc", "x"))).toEqual({ text: "ab\ncd", count: 0 })
        expect(replacePlain("ab\nab", spec("ab", "x"))).toEqual({ text: "x\nx", count: 2 })
    })

    it("keeps the runs' formatting and drops emptied runs", () => {
        const done = replaceRichText([ { runs: [ { text: "Ha", bold: true }, { text: "llo", italic: true } ], level: 1 },
            { text: "Hallo", color: "FF0000" } ], spec("hallo", "Hi"))
        expect(done).toEqual({ count: 2, text: [ { runs: [ { text: "Hi", bold: true } ], level: 1 }, { text: "Hi", color: "FF0000" } ] })
    })

    it("returns the very input when nothing matched", () => {
        const text = [ { text: "x" } ]
        expect(replaceRichText(text, spec("y", "z")).text).toBe(text)
    })
})
