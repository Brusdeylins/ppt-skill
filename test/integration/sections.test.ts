/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: PowerPoint sections. The fill field `section` starts a
**  section at a slide (a slide inside a section splits it, the slides
**  before the first section of a deck without sections form "Default
**  Section"), `section.rename` renames one, `section.rm` merges its slides
**  into the previous section (the next one when it was the first). The
**  ops compose with the structural ops of the same document; kept
**  sections keep their id; `state` reports each slide's section.
*/

import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import JSZip from "jszip"
import { cmdState } from "../../src/commands/state.js"
import { apply, failure, freshDeck, stateOf } from "../util/elements.js"

/**  a titled slide.add  */
const add = (title: string, extra: Record<string, unknown> = {}): Record<string, unknown> =>
    ({ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: title } }, ...extra })

/**  section name → slide titles, in section order  */
const sections = async (deck: string): Promise<[ string, string[] ][]> => {
    const state = await stateOf(deck)
    const title = new Map(state.slides.map((s) => [ s.id, s.title ?? "" ]))
    return state.sections.map((s) => [ s.name, s.slideIds.map((id) => title.get(id) ?? `?${id}`) ])
}

/**  the section ids of the deck  */
const ids = async (deck: string): Promise<string[]> => (await stateOf(deck)).sections.map((s) => s.id)

describe("sections", () => {
    it("starts sections, splits them and files the first slides into the Default Section", async () => {
        const deck = await freshDeck("sections-start.pptx")
        await apply(deck, [ add("A"), add("B"), add("C"), add("D") ])
        await apply(deck, [ { op: "slide.fill", slide: "title:C", section: "Teil 2" } ])
        expect(await sections(deck)).toEqual([ [ "Default Section", [ "A", "B" ] ], [ "Teil 2", [ "C", "D" ] ] ])
        const [ first, second ] = await ids(deck)
        expect(first).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-8[0-9A-F]{3}-[0-9A-F]{12}\}$/)

        /*  a new slide inside a section splits it; one before a section
            start forms a section of its own  */
        await apply(deck, [ add("B2", { at: 2, section: "Mitte" }), add("Z", { at: 0, section: "Vorn" }) ])
        expect(await sections(deck)).toEqual([ [ "Vorn", [ "Z" ] ], [ "Default Section", [ "A", "B" ] ],
            [ "Mitte", [ "B2" ] ], [ "Teil 2", [ "C", "D" ] ] ])

        /*  a slide inside a section splits it; kept sections keep their id  */
        await apply(deck, [ { op: "slide.fill", slide: "title:D", section: "Ende" } ])
        expect(await sections(deck)).toEqual([ [ "Vorn", [ "Z" ] ], [ "Default Section", [ "A", "B" ] ],
            [ "Mitte", [ "B2" ] ], [ "Teil 2", [ "C" ] ], [ "Ende", [ "D" ] ] ])
        const now = await ids(deck)
        expect(now[1]).toBe(first)
        expect(now[3]).toBe(second)

        /*  the same name again is a no-op on its members, an error elsewhere  */
        await apply(deck, [ { op: "slide.fill", slide: "title:D", section: "Ende" } ])
        expect(await failure(deck, [ { op: "slide.fill", slide: "title:A", section: "Ende" } ])).toBe("E_SCHEMA")
    })

    it("starts one section over all slides at the first slide; the first slide of a section takes it over", async () => {
        const deck = await freshDeck("sections-first.pptx")
        await apply(deck, [ add("A", { section: "Eins" }), add("B"), add("C") ])
        expect(await sections(deck)).toEqual([ [ "Eins", [ "A", "B", "C" ] ] ])
        await apply(deck, [ { op: "slide.fill", slide: "title:A", section: "Start" } ])
        expect(await sections(deck)).toEqual([ [ "Start", [ "A", "B", "C" ] ] ])
    })

    it("renames and removes sections, composing with moves and removals of the same document", async () => {
        const deck = await freshDeck("sections-ops.pptx")
        await apply(deck, [ add("A", { section: "S1" }), add("B"), add("C", { section: "S2" }), add("D"),
            add("E", { section: "S3" }), add("F") ])
        const before = await ids(deck)
        await apply(deck, [
            { op: "section.rename", from: "S2", to: "Mitte" },
            { op: "slide.move", slide: "title:A", to: 3 },
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "G" } } },
            { op: "slide.fill", slide: "title:G", section: "Neu" }
        ])
        expect(await sections(deck)).toEqual([ [ "S1", [ "B" ] ], [ "Mitte", [ "C", "D", "A" ] ], [ "S3", [ "E", "F" ] ],
            [ "Neu", [ "G" ] ] ])
        expect((await ids(deck)).slice(0, 3)).toEqual(before)

        /*  removed: its slides join the previous section, the next one for the first  */
        await apply(deck, [ { op: "section.rm", name: "Mitte" }, { op: "section.rm", name: "S1" } ])
        expect(await sections(deck)).toEqual([ [ "S3", [ "B", "C", "D", "A", "E", "F" ] ], [ "Neu", [ "G" ] ] ])
        await apply(deck, [ { op: "section.rm", name: "Neu" }, { op: "section.rm", name: "S3" } ])
        expect(await sections(deck)).toEqual([])
        const zip = await JSZip.loadAsync(readFileSync(deck))
        expect(await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")).not.toContain("sectionLst")
    })

    it("fails on unknown and duplicate names", async () => {
        const deck = await freshDeck("sections-errors.pptx")
        await apply(deck, [ add("A", { section: "S1" }), add("B", { section: "S2" }) ])
        expect(await failure(deck, [ { op: "section.rename", from: "X", to: "Y" } ])).toBe("E_ADDR_NOTFOUND")
        expect(await failure(deck, [ { op: "section.rm", name: "X" } ])).toBe("E_ADDR_NOTFOUND")
        expect(await failure(deck, [ { op: "section.rename", from: "S1", to: "S2" } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ add("C", { section: "S1" }) ])).toBe("E_SCHEMA")
        /*  sequence: after a rename, the old name is free  */
        await apply(deck, [ { op: "section.rename", from: "S1", to: "Intro" }, add("C", { section: "S1" }) ])
        expect(await sections(deck)).toEqual([ [ "Intro", [ "A" ] ], [ "S2", [ "B" ] ], [ "S1", [ "C" ] ] ])
    })

    it("reports each slide's section in state at the text and full levels", async () => {
        const deck = await freshDeck("sections-state.pptx")
        await apply(deck, [ add("A"), add("B", { section: "Haupt" }) ])
        for (const level of [ "text", "full" ]) {
            const out = await cmdState([ deck, "--level", level ]) as { result: { sections: string[], slides: { section?: string }[] } }
            expect(out.result.sections).toEqual([ "Default Section", "Haupt" ])
            expect(out.result.slides.map((s) => s.section)).toEqual([ "Default Section", "Haupt" ])
        }
    })
})
