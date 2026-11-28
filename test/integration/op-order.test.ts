/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: ops of one document that meet on the same element lose
**  nothing silently although the engine applies a slide's work by kind
**  (fills, el.set edits, footer band, elements). el.rm takes the edits
**  and stacking moves planned for the element along; a later slide.fill
**  text replaces an earlier el.set text (an append to it is rejected) and
**  the tracked title follows; an el.set on a footer, slide-number or date
**  placeholder holds although a slide.fill sets that placeholder anew in
**  the same document, in either order -- a stacking move too, also when
**  the layout's copy has another name; one removed is gone. On a layout
**  without such a placeholder the slide keeps its own (a footer text goes
**  into it), and the el.set edits apply to it.
*/

import { beforeEach, describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import JSZip from "jszip"
import { relTargets, relsPartOf } from "../../src/engine/parts.js"
import { elements, firstElement, parseXml, serializeXml } from "../../src/engine/xml.js"
import { apply, craftSlide, failure, freshDeck, stateOf } from "../util/elements.js"

let deck: string

/**  a shape of slide 0 by name prefix  */
const shape = async (name: string): Promise<{ name: string, text: string | null, altText?: string, fontColor?: string } | undefined> =>
    (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.name === name || sh.name.startsWith(`${name}-`)) as never

/**  take the footer-band placeholders of these types off the layout of
     slide 0 (a layout without them; the slide keeps its own)  */
const dropFromLayout = async (types: string[]): Promise<void> => {
    const part = (await stateOf(deck)).slides[0]?.part ?? ""
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const rels = parseXml(await (zip.file(relsPartOf(part)) as JSZip.JSZipObject).async("string"))
    const layout = relTargets(part, rels).find((r) => r.kind === "slideLayout")?.target ?? ""
    const doc = parseXml(await (zip.file(layout) as JSZip.JSZipObject).async("string"))
    const band = elements(doc, "p:sp").filter((sp) => types.includes(firstElement(sp, "p:ph")?.getAttribute("type") ?? ""))
    expect(band).toHaveLength(types.length)
    for (const sp of band)
        sp.parentNode?.removeChild(sp)
    zip.file(layout, serializeXml(doc))
    writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
}

/**  the name of a placeholder of slide 0 by kind  */
const placeholder = async (kind: string): Promise<string> =>
    (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.placeholderKind === kind)?.name ?? ""

beforeEach(async () => {
    deck = await freshDeck("op-order.pptx")
    await apply(deck, [
        { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Ordnung" }, body: { text: "Inhalt" } }, footer: "Fuss" },
        { op: "el.add", slide: "0", elements: [ { type: "textbox", name: "T", frame: { x: 1, y: 6, w: 3, h: 0.5 }, text: "weg" } ] }
    ])
})

describe("el.rm after el.set", () => {
    it("drops the edits and stacking moves of the removed element", async () => {
        await apply(deck, [
            { op: "el.set", slide: "0", name: "T", text: "geändert", fill: "FF0000", zOrder: "back" },
            { op: "el.rm", slide: "0", name: "T" },
            { op: "el.add", slide: "0", elements: [ { type: "textbox", name: "N", frame: { x: 5, y: 6, w: 2, h: 0.5 }, text: "neu", zOrder: "back" } ] },
            { op: "el.rm", slide: "0", name: "N" }
        ])
        expect(await shape("T")).toBeUndefined()
        expect(await shape("N")).toBeUndefined()
    })
})

describe("slide.fill after el.set text", () => {
    it("replaces the el.set text, the title selector follows", async () => {
        const body = await placeholder("body")
        const title = await placeholder("title")
        await apply(deck, [
            { op: "el.set", slide: "0", name: body, text: "von el.set", fontColor: "C00000" },
            { op: "slide.fill", slide: "0", placeholders: { body: { text: "von slide.fill" } } },
            { op: "el.set", slide: "0", name: title, text: "Titel X" },
            { op: "slide.fill", slide: "title:Titel X", placeholders: { title: { text: "Titel Y" } } },
            { op: "slide.fill", slide: "title:Titel Y", notes: "gefunden" }
        ])
        const state = await stateOf(deck)
        expect(state.slides[0]?.title).toBe("Titel Y")
        expect(state.slides[0]?.notes).toBe("gefunden")
        expect((await shape(body))?.text).toBe("von slide.fill")
        expect((await shape(body))?.fontColor).toBe("C00000")
    })

    it("rejects an append to an el.set text of the same document", async () => {
        const body = await placeholder("body")
        expect(await failure(deck, [
            { op: "el.set", slide: "0", name: body, text: "von el.set" },
            { op: "slide.fill", slide: "0", placeholders: { body: { text: "dazu", append: true } } }
        ])).toBe("E_SCHEMA")
    })
})

describe("el.set on footer-band placeholders set anew", () => {
    it("applies the edit after a slide.fill footer in either order", async () => {
        const footer = await placeholder("footer")
        await apply(deck, [
            { op: "slide.fill", slide: "0", footer: "Neuer Fuss" },
            { op: "el.set", slide: "0", name: footer, fontColor: "FF0000", altText: "Fusszeile" }
        ])
        expect(await shape(footer)).toMatchObject({ text: "Neuer Fuss", altText: "Fusszeile", fontColor: "FF0000" })
        await apply(deck, [
            { op: "el.set", slide: "0", name: footer, altText: "zuerst", text: "überholt" },
            { op: "slide.fill", slide: "0", footer: "Dritter Fuss" }
        ])
        expect(await shape(footer)).toMatchObject({ text: "Dritter Fuss", altText: "zuerst" })
    })

    it("drops the edits of a slide number turned off and rejects later ones", async () => {
        const number = await placeholder("slideNumber")
        await apply(deck, [
            { op: "el.set", slide: "0", name: number, altText: "Nummer" },
            { op: "slide.fill", slide: "0", slideNumber: false }
        ])
        expect(await placeholder("slideNumber")).toBe("")
        await apply(deck, [ { op: "slide.fill", slide: "0", date: true } ])
        const date = await placeholder("date")
        expect(await failure(deck, [
            { op: "slide.fill", slide: "0", date: false },
            { op: "el.set", slide: "0", name: date, altText: "Datum" }
        ])).toBe("E_ADDR_NOTFOUND")
    })
})

describe("stacking moves on footer-band placeholders set anew", () => {
    it("moves the layout's copy although it is named otherwise, in either order", async () => {
        const rename = async (): Promise<void> => {
            const footer = await placeholder("footer")
            await craftSlide(deck, 0, (xml) => xml.replace(`name="${footer}"`, "name=\"Mein Fuss\""))
        }
        const back = async (): Promise<{ kind: string | null, text: string | null }> => {
            const first = (await stateOf(deck)).slides[0]?.shapes[0]
            return { kind: first?.placeholderKind ?? null, text: first?.text ?? null }
        }
        await rename()
        await apply(deck, [
            { op: "el.set", slide: "0", name: "Mein Fuss", zOrder: "back" },
            { op: "slide.fill", slide: "0", footer: "Neu" }
        ])
        expect(await back()).toEqual({ kind: "footer", text: "Neu" })
        await apply(deck, [ { op: "el.set", slide: "0", name: await placeholder("footer"), zOrder: "front" } ])
        await rename()
        await apply(deck, [
            { op: "slide.fill", slide: "0", footer: "Neu 2" },
            { op: "el.set", slide: "0", name: "Mein Fuss", zOrder: "back" }
        ])
        expect(await back()).toEqual({ kind: "footer", text: "Neu 2" })
    })
})

describe("footer-band placeholders on a layout without them", () => {
    it("keeps the slide's own on true and a footer text", async () => {
        await dropFromLayout([ "sldNum", "ftr" ])
        const number = await placeholder("slideNumber")
        const footer = await placeholder("footer")
        expect(number).not.toBe("")
        await apply(deck, [ { op: "slide.fill", slide: "0", slideNumber: true, footer: "Anderer Fuss" } ])
        expect(await placeholder("slideNumber")).toBe(number)
        expect(await shape(footer)).toMatchObject({ text: "Anderer Fuss" })
    })

    it("applies an el.set on the kept placeholder, in either order", async () => {
        await dropFromLayout([ "sldNum", "ftr" ])
        const number = await placeholder("slideNumber")
        const footer = await placeholder("footer")
        await apply(deck, [
            { op: "el.set", slide: "0", name: number, fontColor: "FF0000" },
            { op: "slide.fill", slide: "0", slideNumber: true },
            { op: "slide.fill", slide: "0", footer: "Fuss 2" },
            { op: "el.set", slide: "0", name: footer, altText: "Fusszeile", zOrder: "back" }
        ])
        expect(await shape(number)).toMatchObject({ fontColor: "FF0000" })
        expect(await shape(footer)).toMatchObject({ text: "Fuss 2", altText: "Fusszeile" })
        expect((await stateOf(deck)).slides[0]?.shapes[0]?.name).toBe(footer)
        await apply(deck, [
            { op: "slide.fill", slide: "0", slideNumber: true },
            { op: "el.set", slide: "0", name: number, fontColor: "00B050" }
        ])
        expect(await shape(number)).toMatchObject({ fontColor: "00B050" })
    })
})
