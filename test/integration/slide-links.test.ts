/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: slide links. A `hyperlink` of a run, shape, textbox or
**  picture takes a URL or `slide:SEL`; a slide link is written as a slide
**  relationship with `ppaction://hlinksldjump`, its target resolved after
**  the whole document (slides added later count, moved ones follow, a
**  missing target fails with E_ADDR_NOTFOUND). Links to slides a later
**  document removes are dropped; copies keep their links; `state --level
**  full` reads them back as `slide:id:N`.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import { readFileSync } from "node:fs"
import type { ShapeInfo, SlideInfo } from "../../src/core/model.js"
import { relTargets, relsPartOf } from "../../src/engine/parts.js"
import { elements, parseXml } from "../../src/engine/xml.js"
import { TMP, apply, failure, freshDeck, png, stateOf } from "../util/elements.js"

let deck: string
let image: string

/**  a slide of the deck by title  */
const slide = async (title: string): Promise<SlideInfo> =>
    (await stateOf(deck)).slides.find((s) => s.title === title) as SlideInfo

/**  a shape of a slide by name prefix  */
const shape = (s: SlideInfo, name: string): ShapeInfo =>
    s.shapes.find((sh) => sh.name.startsWith(name)) as ShapeInfo

/**  the run links of a shape  */
const runLinks = (sh: ShapeInfo): (string | undefined)[] =>
    (sh.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.hyperlink))

/**  every a:hlinkClick of a slide with its relationship  */
const links = async (title: string): Promise<{ action: string | null, type: string, target: string | null, mode: string | null }[]> => {
    const s = await slide(title)
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const doc = parseXml(await (zip.file(s.part) as JSZip.JSZipObject).async("string"))
    const rels = relTargets(s.part, parseXml(await (zip.file(relsPartOf(s.part)) as JSZip.JSZipObject).async("string")))
    return elements(doc, "a:hlinkClick").map((h) => {
        const rel = rels.find((r) => r.id === h.getAttribute("r:id"))
        return { action: h.getAttribute("action"), type: rel?.kind ?? "?", target: rel?.target ?? null,
            mode: rel?.el.getAttribute("TargetMode") ?? null }
    })
}

/**  a chapter slide with a back button to the agenda  */
const chapter = (title: string, ref: string): unknown[] => [
    { op: "slide.add", layout: "Title Only", ref, placeholders: { title: { text: title } } },
    { op: "el.add", slide: `$${ref}`, elements: [
        { type: "shape", name: "Zurück", shape: "roundRect", frame: { x: 8, y: 6, w: 1.5, h: 0.5 }, text: "Agenda",
            hyperlink: "slide:title:Agenda" } ] }
]

beforeAll(async () => {
    deck = await freshDeck("slide-links.pptx")
    image = path.join(TMP, "slide-links.png")
    writeFileSync(image, png(40, 30))

    /*  the agenda links forward to slides added later in the document  */
    await apply(deck, [
        { op: "slide.add", layout: "CONTENT", ref: "agenda", placeholders: { title: { text: "Agenda" },
            body: { text: [ { runs: [ { text: "Kapitel 1", hyperlink: "slide:$k1" } ] },
                { runs: [ { text: "Kapitel 2", hyperlink: "slide:$k2" } ] },
                { runs: [ { text: "Web", hyperlink: "https://example.com/" } ] } ] } } },
        { op: "el.add", slide: "$agenda", elements: [
            { type: "textbox", name: "Box", frame: { x: 0.5, y: 6, w: 3, h: 0.5 },
                text: [ { runs: [ { text: "Ende", hyperlink: "slide:index:3" } ] } ] },
            { type: "image", name: "Bild", frame: { x: 8, y: 6, w: 1, h: 0.75 }, path: image, hyperlink: "slide:$k2" },
            { type: "table", name: "Tab", frame: { x: 4, y: 6, w: 3 },
                data: { rows: [ [ { text: [ { runs: [ { text: "K1", hyperlink: "slide:$k1" } ] } ] }, "x" ] ] } } ] },
        ...chapter("Kapitel 1", "k1"),
        ...chapter("Kapitel 2", "k2"),
        { op: "slide.add", layout: "Title Only", placeholders: { title: { text: "Schluss" } } }
    ])
})

describe("slide links", () => {
    it("writes slide relationships with a slide jump, URLs stay external", async () => {
        const ids = new Map((await stateOf(deck)).slides.map((s) => [ s.title, s.part ]))
        const agenda = await links("Agenda")
        const jumps = agenda.filter((l) => l.type === "slide")
        expect(jumps.every((l) => l.action === "ppaction://hlinksldjump")).toBe(true)
        expect(jumps.map((l) => l.target).sort()).toEqual([ ids.get("Kapitel 1"), ids.get("Kapitel 1"), ids.get("Kapitel 2"),
            ids.get("Kapitel 2"), ids.get("Schluss") ].sort())
        expect(agenda.filter((l) => l.type === "hyperlink")).toEqual([ { action: null, type: "hyperlink", target: null, mode: "External" } ])
        expect(await links("Kapitel 1")).toEqual([ { action: "ppaction://hlinksldjump", type: "slide", target: ids.get("Agenda"), mode: null } ])
    })

    it("reads the links back in state", async () => {
        const agenda = await slide("Agenda")
        const k1 = (await slide("Kapitel 1")).id
        const k2 = (await slide("Kapitel 2")).id
        const end = (await slide("Schluss")).id
        expect(runLinks(shape(agenda, "PptcPh-1"))).toEqual([ `slide:id:${k1}`, `slide:id:${k2}`, "https://example.com/" ])
        expect(runLinks(shape(agenda, "Box"))).toEqual([ `slide:id:${end}` ])
        expect(shape(agenda, "Bild").hyperlink).toBe(`slide:id:${k2}`)
        expect(shape(await slide("Kapitel 2"), "Zurück").hyperlink).toBe(`slide:id:${agenda.id}`)
    })

    it("lets moved targets follow and copies keep their links", async () => {
        await apply(deck, [ { op: "slide.move", slide: "title:Kapitel 2", to: 1 },
            { op: "slide.copy", slide: "title:Kapitel 1", ref: "c" },
            { op: "slide.fill", slide: "$c", placeholders: { title: { text: "Kapitel 1b" } } } ])
        const agenda = await slide("Agenda")
        expect(runLinks(shape(agenda, "PptcPh-1")).slice(0, 2)).toEqual([ `slide:id:${(await slide("Kapitel 1")).id}`,
            `slide:id:${(await slide("Kapitel 2")).id}` ])
        expect(shape(await slide("Kapitel 1b"), "Zurück").hyperlink).toBe(`slide:id:${agenda.id}`)
    })

    it("sets, changes and removes the click target of elements with el.set", async () => {
        await apply(deck, [
            { op: "el.set", slide: "title:Kapitel 1", name: "PptcPh-0", hyperlink: "slide:title:Schluss" },
            { op: "el.set", slide: "title:Kapitel 2", name: "Zurück", hyperlink: "https://example.org/" },
            { op: "el.set", slide: "title:Kapitel 1b", name: "Zurück", hyperlink: "" },
            { op: "slide.add", layout: "Title Only", ref: "n", placeholders: { title: { text: "Neu" } } },
            { op: "el.add", slide: "$n", elements: [ { type: "textbox", name: "Hin", frame: { x: 1, y: 2, w: 2, h: 1 }, text: "hin" } ] },
            { op: "el.set", slide: "$n", name: "Hin", hyperlink: "slide:$n" }
        ])
        expect(shape(await slide("Kapitel 1"), "PptcPh-0").hyperlink).toBe(`slide:id:${(await slide("Schluss")).id}`)
        expect(shape(await slide("Kapitel 2"), "Zurück").hyperlink).toBe("https://example.org/")
        expect(shape(await slide("Kapitel 1b"), "Zurück").hyperlink).toBeUndefined()
        expect((await links("Kapitel 1b")).filter((l) => l.type !== "?")).toEqual([])
        const neu = await slide("Neu")
        expect(shape(neu, "Hin").hyperlink).toBe(`slide:id:${neu.id}`)
    })

    it("fails at plan time when a target does not exist after the document", async () => {
        expect(await failure(deck, [ { op: "el.set", slide: "title:Neu", name: "Hin", hyperlink: "slide:title:Nirgends" } ]))
            .toBe("E_ADDR_NOTFOUND")
        expect(await failure(deck, [
            { op: "slide.add", layout: "Title Only", ref: "x", placeholders: { title: { text: "X" } } },
            { op: "el.set", slide: "title:Neu", name: "Hin", hyperlink: "slide:$x" },
            { op: "slide.rm", slide: "$x" } ])).toBe("E_ADDR_NOTFOUND")
        expect(await failure(deck, [ { op: "el.set", slide: "title:Neu", name: "Hin", hyperlink: "slide:nirgends" } ]))
            .toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.set", slide: "title:Neu", name: "Hin", hyperlink: "slide:index:99" } ]))
            .toBe("E_ADDR_NOTFOUND")
    })

    it("drops links to slides a later document removes", async () => {
        await apply(deck, [ { op: "slide.rm", slide: "title:Schluss" } ])
        expect(shape(await slide("Kapitel 1"), "PptcPh-0").hyperlink).toBeUndefined()
        expect(runLinks(shape(await slide("Agenda"), "Box"))).toEqual([])
        expect((await slide("Agenda")).shapes.find((s) => s.name.startsWith("Box"))?.text).toBe("Ende")
    })
})
