/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  commands/state: the read plane for decks. One command, three zoom levels:
**  summary (deck overview), text (all text content), full (complete shape
**  model with geometry and ids). Always includes the rev token.
*/

import { PptcError } from "../infra/errors.js"
import { resolveSlide } from "../core/selector.js"
import type { DeckState, SlideInfo } from "../core/model.js"
import { DeckArchive, readDeckState } from "../engine/reader.js"
import { parse } from "../infra/args.js"

/**  zoom level of the state output  */
type Level = "summary" | "text" | "full"

/**  render one slide at the requested zoom level  */
const renderSlide = (slide: SlideInfo, level: Level): Record<string, unknown> => {
    const base = {
        id: slide.id,
        index: slide.index,
        title: slide.title,
        layout: slide.layoutName,
        layoutIndex: slide.layoutIndex,
        ...(slide.hidden && { hidden: true })
    }
    if (level === "summary")
        return { ...base, shapes: slide.shapes.length, hasNotes: slide.notes !== null }
    const section = slide.section !== null ? { section: slide.section } : {}
    if (level === "text")
        return {
            ...base,
            ...section,
            texts: slide.shapes
                .filter((s) => s.text !== null || s.table !== undefined)
                .map((s) => ({
                    name: s.name,
                    ...(s.placeholderIdx !== null && { placeholder: s.placeholderIdx }),
                    ...(s.text !== null && { text: s.text }),
                    ...(s.table !== undefined && { table: s.table })
                })),
            notes: slide.notes
        }
    return { ...base, ...section, shapes: slide.shapes, notes: slide.notes }
}

/**
 *  CLI command `pptc state <deck> [--slide SEL] [--level summary|text|full]`.
 *
 *  @param argv - raw arguments after the command name
 *  @returns the deck state payload for the envelope
 */
export const cmdState = async (argv: string[]): Promise<Record<string, unknown>> => {
    const args = parse(argv, {
        "slide": { type: "string" },
        "level": { type: "string" },
        "plain": { type: "boolean" }
    }, ["deck"])
    const level = (args.str("level") ?? "text") as Level
    if (!["summary", "text", "full"].includes(level))
        throw new PptcError("E_USAGE", `unknown level '${level}' (summary|text|full)`)

    const deck: DeckState = await readDeckState(await DeckArchive.open(args.positionals[0] as string))
    const selector = args.str("slide")
    const slides = selector === null
        ? deck.slides
        : [deck.slides[resolveSlide(selector, deck.slides, new Map<string, number>()).index] as SlideInfo]
    if (args.flag("plain"))
        return { plain: [
            `${deck.file}  rev:${deck.rev}  ${deck.slides.length} slide(s)`,
            ...slides.map((s) =>
                `  #${s.index} id:${s.id} '${s.title ?? ""}' (layout ${s.layoutIndex})`
                + (s.notes !== null && s.notes !== "" ? "  [notes]" : "")
                + (s.hidden ? "  [hidden]" : ""))
        ].join("\n") }
    return {
        file: deck.file,
        rev: deck.rev,
        result: {
            slideSize: deck.slideSize,
            slideCount: deck.slides.length,
            ...(deck.sections.length > 0 && { sections: deck.sections.map((s) => s.name) }),
            customProps: deck.customProps,
            slides: slides.map((s) => renderSlide(s, level))
        }
    }
}
