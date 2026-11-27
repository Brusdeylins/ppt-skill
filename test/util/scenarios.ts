/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  test/util/scenarios: fixed op sequences exercising all ten ops, for the
**  golden test. Each step is computed from the current deck state, so the
**  same sequence runs on the neutral fixture, the demo deck and the
**  private templates.
*/

import { mkdirSync, writeFileSync } from "node:fs"
import path from "node:path"
import type { DeckState, Layout } from "../../src/core/model.js"

/**  one apply: the ops to run against the current deck state  */
export type Step = (state: DeckState) => unknown[]

/**  smallest valid 1x1 PNG  */
const PIXEL = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64")

/**  media files the scenarios reference  */
export interface Media { png: string, png2: string, svg: string }

/**
 *  Write the scenario media into a folder.
 *
 *  @param dir - target folder
 *  @returns the file paths
 */
export const writeMedia = (dir: string): Media => {
    mkdirSync(dir, { recursive: true })
    const media = { png: path.join(dir, "scenario-a.png"), png2: path.join(dir, "scenario-b.png"), svg: path.join(dir, "scenario.svg") }
    writeFileSync(media.png, Buffer.concat([ PIXEL, Buffer.from("a") ]))
    writeFileSync(media.png2, Buffer.concat([ PIXEL, Buffer.from("b") ]))
    writeFileSync(media.svg, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10\" height=\"10\"><rect width=\"10\" height=\"10\"/></svg>")
    return media
}

/**  the layouts the scenarios use: one with title and body, one with a
     picture placeholder (null when the deck has none)  */
const pickLayouts = (layouts: Layout[]): { content: Layout, picture: Layout | null } => {
    const has = (l: Layout, kind: string): boolean => l.placeholders.some((p) => p.kind === kind)
    const content = layouts.find((l) => has(l, "title") && l.placeholders.filter((p) => p.kind === "body").length === 1)
        ?? layouts.find((l) => has(l, "title")) ?? layouts[0] as Layout
    return { content, picture: layouts.find((l) => has(l, "picture") && has(l, "title")) ?? null }
}

/**  the element set of the scenarios: one of every el.add type  */
const allElements = (media: Media): unknown[] => [
    { type: "textbox", name: "Box", frame: { x: 0.5, y: 5.5, w: 3, h: 0.8 },
        text: [ { runs: [ { text: "see " }, { text: "docs", hyperlink: "https://example.com/docs" } ] }, { text: "line 2", bullet: false } ] },
    { type: "table", name: "Tab", frame: { x: 4, y: 5.5, w: 4 },
        data: { headers: [ "h1", "h2" ], rows: [ [ "a", 1 ], [ "b", 2 ] ], style: { altRowBg: "EEEEEE" } } },
    { type: "chart", name: "Chart", frame: { x: 8.5, y: 4, w: 4, h: 3 },
        data: { type: "column", categories: [ "x", "y" ], series: [ { name: "S-1", values: [ 1, 2 ] } ], title: "C" } },
    { type: "chart", frame: { x: 8.5, y: 1, w: 3, h: 2.5 },
        data: { type: "pie", categories: [ "p", "q" ], series: [ { name: "S-2", values: [ 3, 4 ] } ] } },
    { type: "shape", name: "Shape", shape: "roundRect", frame: { x: 0.5, y: 1.5, w: 1.5, h: 0.8 }, text: "S", fill: "C00000", fontColor: "FFFFFF" },
    { type: "image", name: "Pic", frame: { x: 2.5, y: 1.5, w: 1, h: 1 }, path: media.png },
    { type: "image", name: "Vec", frame: { x: 4, y: 1.5, w: 1, h: 1 }, path: media.svg },
    { type: "connector", name: "Line", from: [ 0.5, 7 ], to: [ 5, 6.6 ], color: "333333", dash: "dash" }
]

/**
 *  The fixed scenario on an empty or populated deck: every op, kept and
 *  new slides, copies of both, moves, removals and properties.
 *
 *  @param layouts - the deck's layouts
 *  @param media - the scenario media
 *  @returns the steps, one apply each
 */
export const fixedScenario = (layouts: Layout[], media: Media): Step[] => {
    const { content, picture } = pickLayouts(layouts)
    const body = content.placeholders.find((p) => p.kind === "body")
    const steps: Step[] = [
        () => [
            { op: "slide.add", ref: "a", layout: content.name, placeholders: {
                title: { text: "Scenario A" },
                ...(body !== undefined && { [String(body.idx)]: { text: [ { text: "point one", level: 0 },
                    { runs: [ { text: "bold", bold: true }, { text: " and a " }, { text: "link", hyperlink: "https://example.com" } ], level: 1 } ] } })
            }, notes: "notes A\nsecond line" },
            { op: "el.add", slide: "$a", elements: allElements(media) }
        ],
        () => [
            { op: "slide.add", ref: "b", layout: content.name, at: 0, placeholders: { title: { text: "Scenario B" } },
                notes: "notes B", footer: "Footer B", background: { color: "F0F0F0" } },
            { op: "slide.copy", slide: "title:Scenario A", ref: "c" },
            { op: "slide.fill", slide: "$c", placeholders: { title: { text: "Scenario C" } }, notes: "notes C" }
        ],
        () => [
            { op: "el.set", slide: "title:Scenario A", name: "Box", text: "Box updated" },
            { op: "el.rm", slide: "title:Scenario C", name: "Chart" },
            { op: "slide.fill", slide: "title:Scenario A", hidden: true },
            { op: "el.add", slide: "title:Scenario B", elements: [
                { type: "textbox", name: "Late", frame: { x: 1, y: 6, w: 3, h: 0.6 }, text: "late" } ] },
            { op: "el.set", slide: "title:Scenario B", name: "Late", text: "late, retexted" }
        ],
        () => [
            { op: "slide.move", slide: "title:Scenario C", to: 0 },
            { op: "slide.copy", slide: "title:Scenario B", ref: "d" },
            { op: "slide.fill", slide: "$d", placeholders: { title: { text: "Scenario D" } }, hidden: true },
            { op: "meta.props", set: { title: "Scenario deck", subject: "golden", custom: { pptcScenario: "1" } } }
        ]
    ]
    if (picture !== null)
        steps.push(() => [
            { op: "slide.add", ref: "p", layout: picture.name, placeholders: { title: { text: "Scenario P" }, image: { image: media.png2 } } },
            { op: "img.prompts", slide: "$p", prompts: "a calm landscape" }
        ], (state) => [
            { op: "slide.copy", slide: "title:Scenario P", ref: "q" },
            { op: "slide.fill", slide: "$q", placeholders: { title: { text: "Scenario Q" } } },
            { op: "el.rm", slide: "$q", name: "PptcPromptBox" },
            ...(state.slides.length > 3 ? [ { op: "slide.move", slide: "title:Scenario P", to: 1 } ] : [])
        ])
    steps.push(() => [
        { op: "slide.rm", slide: "title:Scenario A" },
        { op: "slide.copy", slide: "title:Scenario C", ref: "e" },
        { op: "slide.fill", slide: "$e", placeholders: { title: { text: "Scenario E" } }, footer: "Footer E" }
    ])
    return steps
}

/**
 *  A scenario on a deck that already has content: edits on kept slides of
 *  every kind (notes, footer, visibility, elements, retext, removal),
 *  copies, moves and removals of existing slides.
 *
 *  @param media - the scenario media
 *  @returns the steps, one apply each
 */
export const existingScenario = (media: Media): Step[] => {
    const at = (state: DeckState, i: number): string => `id:${(state.slides[i % state.slides.length] as { id: number }).id}`
    /*  a uniquely addressable shape of a type; generated names are used
        without their UUID suffix, which differs from deck to deck  */
    const named = (state: DeckState, i: number, type: string): { slide: string, name: string } | null => {
        const address = (name: string): string =>
            name.replace(/-?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/, "")
        for (let k = 0; k < state.slides.length; k++) {
            const s = state.slides[(i + k) % state.slides.length] as DeckState["slides"][number]
            const shape = s.shapes.find((sh) => sh.type === type && address(sh.name) !== ""
                && s.shapes.filter((o) => o.name === address(sh.name) || o.name.startsWith(`${address(sh.name)}-`)).length === 1)
            if (shape !== undefined)
                return { slide: `id:${s.id}`, name: address(shape.name) }
        }
        return null
    }
    return [
        (state) => [
            { op: "slide.fill", slide: at(state, 1), notes: "kept notes replaced" },
            { op: "slide.fill", slide: at(state, 2), footer: "kept footer" },
            { op: "el.add", slide: at(state, 3), elements: [
                { type: "textbox", name: "Added", frame: { x: 1, y: 6.5, w: 3, h: 0.5 }, text: "added to a kept slide" },
                { type: "image", name: "AddedPic", frame: { x: 5, y: 6.5, w: 0.5, h: 0.5 }, path: media.png } ] }
        ],
        (state) => {
            const box = named(state, 4, "textbox") ?? named(state, 4, "shape")
            const pic = named(state, 5, "picture")
            return [
                ...(box !== null ? [ { op: "el.set", slide: box.slide, name: box.name, text: "retexted" } ] : []),
                ...(pic !== null ? [ { op: "el.rm", slide: pic.slide, name: pic.name } ] : []),
                { op: "slide.fill", slide: at(state, 6), hidden: true }
            ]
        },
        (state) => state.slides.map((s) => ({ op: "slide.copy", slide: `id:${s.id}` })),
        (state) => [
            { op: "slide.move", slide: at(state, 0), to: state.slides.length - 1 },
            { op: "slide.rm", slide: at(state, 2) },
            { op: "slide.rm", slide: at(state, 5) }
        ],
        (state) => [
            { op: "slide.copy", slide: at(state, 1), ref: "x" },
            { op: "el.add", slide: "$x", elements: [
                { type: "chart", name: "KeptChart", frame: { x: 1, y: 4, w: 4, h: 2.5 },
                    data: { type: "bar", categories: [ "u", "v" ], series: [ { name: "S-K", values: [ 5, 6 ] } ] } } ] },
            { op: "meta.props", set: { keywords: "kept", custom: { pptcScenario: "kept" } } }
        ]
    ]
}
