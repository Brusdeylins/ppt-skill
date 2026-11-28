/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  test/util/elements: a fresh deck from the neutral fixture, applies
**  checked for integrity, and the XML of a generated element (its shape on
**  the slide, its chart part) by element name.
*/

import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import type { Element } from "@xmldom/xmldom"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { PptcError } from "../../src/infra/errors.js"
import type { DeckState } from "../../src/core/model.js"
import { relTargets, relsPartOf } from "../../src/engine/parts.js"
import { NS_R, elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { expectIntact } from "./integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))

/**  the neutral fixture template  */
export const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")

/**  scratch directory of the tests  */
export const TMP = path.join(here, "..", "tmp")

/**  apply options of the tests  */
export const OPTS: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**
 *  Create an empty deck from the neutral fixture.
 *
 *  @param name - file name under test/tmp
 *  @returns the deck path
 */
export const freshDeck = async (name: string): Promise<string> => {
    mkdirSync(TMP, { recursive: true })
    const deck = path.join(TMP, name)
    writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
    return deck
}

/**
 *  Apply ops and check the written deck for repair triggers.
 *
 *  @param deck - the deck path
 *  @param ops - the ops of the document
 *  @returns the warnings of the apply
 */
export const apply = async (deck: string, ops: unknown[]): Promise<unknown[]> => {
    const result = await executeOps(deck, { ops } as never, OPTS) as unknown as { warnings?: unknown[] }
    await expectIntact(deck)
    return result.warnings ?? []
}

/**
 *  The error code an apply fails with, null when it succeeds.
 *
 *  @param deck - the deck path
 *  @param ops - the ops of the document
 *  @returns the PptcError code
 */
export const failure = async (deck: string, ops: unknown[]): Promise<string | null> => {
    try {
        await executeOps(deck, { ops } as never, OPTS)
        return null
    }
    catch (err) {
        return err instanceof PptcError ? err.code : String(err)
    }
}

/**
 *  The state of a deck.
 *
 *  @param deck - the deck path
 *  @returns its read model
 */
export const stateOf = async (deck: string): Promise<DeckState> =>
    await readDeckState(await DeckArchive.open(deck))

/**
 *  The shape of an element on the last slide, by its name prefix.
 *
 *  @param deck - the deck path
 *  @param name - the element name (without the generated suffix)
 *  @returns the shape element and the slide part
 */
export const shapeOf = async (deck: string, name: string): Promise<{ shape: Element, part: string, zip: JSZip }> => {
    const state = await stateOf(deck)
    const part = (state.slides.at(-1) as { part: string }).part
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const doc = parseXml(await (zip.file(part) as JSZip.JSZipObject).async("string"))
    const cNvPr = elements(doc, "p:cNvPr").find((el) => (el.getAttribute("name") ?? "").startsWith(`${name}-`))
    if (cNvPr === undefined)
        throw new Error(`no element '${name}' on the last slide`)
    return { shape: cNvPr.parentNode?.parentNode as Element, part, zip }
}

/**
 *  The chart part XML of a chart element on the last slide.
 *
 *  @param deck - the deck path
 *  @param name - the element name
 *  @returns the chart part DOM root
 */
export const chartOf = async (deck: string, name: string): Promise<Element> => {
    const { shape, part, zip } = await shapeOf(deck, name)
    const rid = firstElement(shape, "c:chart")?.getAttributeNS(NS_R, "id")
    const rels = parseXml(await (zip.file(relsPartOf(part)) as JSZip.JSZipObject).async("string"))
    const target = relTargets(part, rels).find((r) => r.id === rid)?.target as string
    return parseXml(await (zip.file(target) as JSZip.JSZipObject).async("string")).documentElement as Element
}

/**
 *  The attribute values of every element with a tag below a scope.
 *
 *  @param scope - the element to search
 *  @param tag - the qualified tag name
 *  @param attr - the attribute
 *  @returns the values in document order (null when absent)
 */
export const attrs = (scope: Element, tag: string, attr: string): (string | null)[] =>
    elements(scope, tag).map((el) => el.getAttribute(attr))

/**
 *  A PNG header of the given pixel size (the engine reads only the header).
 *
 *  @param w - width in pixels
 *  @param h - height in pixels
 *  @returns the bytes
 */
export const png = (w: number, h: number): Buffer => {
    const b = Buffer.alloc(33)
    Buffer.from([ 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a ]).copy(b, 0)
    b.writeUInt32BE(13, 8)
    b.write("IHDR", 12, "latin1")
    b.writeUInt32BE(w, 16)
    b.writeUInt32BE(h, 20)
    return b
}
