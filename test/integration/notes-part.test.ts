/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the notes part speaker notes are written into. A new one
**  is built from the notes master's placeholders as PowerPoint does it
**  (slide image, notes body with the master's idx, slide number), related
**  to the slide and the notes master; an existing one keeps everything but
**  the text of its notes body.
*/

import { describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import type { Document, Element } from "@xmldom/xmldom"
import { elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { apply, freshDeck, stateOf } from "../util/elements.js"

/**  the notes part of a slide with its DOM and relationships  */
const notesOf = async (deck: string, index: number): Promise<{ part: string, doc: Document, rels: string }> => {
    const part = ((await stateOf(deck)).slides[index] as { part: string }).part
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const slideRels = await (zip.file(`ppt/slides/_rels/${path.posix.basename(part)}.rels`) as JSZip.JSZipObject).async("string")
    const target = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(slideRels)?.[1] as string
    const notes = `ppt/notesSlides/${target}`
    return {
        part: notes,
        doc: parseXml(await (zip.file(notes) as JSZip.JSZipObject).async("string")),
        rels: await (zip.file(`ppt/notesSlides/_rels/${target}.rels`) as JSZip.JSZipObject).async("string")
    }
}

/**  the placeholders of a notes part: type and idx  */
const placeholders = (doc: Document): string[] =>
    elements(doc, "p:ph").map((ph) => `${ph.getAttribute("type") ?? ""}:${ph.getAttribute("idx") ?? "-"}`)

/**  the text of the notes body  */
const bodyText = (doc: Document): string[] => {
    const body = elements(doc, "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === "body") as Element
    return elements(body, "a:p").map((p) => elements(p, "a:t").map((t) => t.textContent ?? "").join(""))
}

describe("speaker notes parts", () => {
    it("builds a new notes part from the notes master's placeholders", async () => {
        const deck = await freshDeck("notes-new.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", notes: "Erste\nZweite" } ])
        const { doc, rels } = await notesOf(deck, 0)

        /*  the fixture's notes master: sldImg idx 2, body idx 3, sldNum idx 5
            (its date placeholder has idx 1, which a body idx 1 inherits from)  */
        expect(placeholders(doc)).toEqual([ "sldImg:2", "body:3", "sldNum:5" ])
        expect(bodyText(doc)).toEqual([ "Erste", "Zweite" ])
        expect(elements(doc, "a:fld").map((f) => f.getAttribute("type"))).toEqual([ "slidenum" ])
        expect(elements(doc, "p:cNvPr").map((c) => c.getAttribute("id"))).toEqual([ "1", "2", "3", "4" ])
        expect(rels).toMatch(/relationships\/notesMaster" Target="\.\.\/notesMasters\/notesMaster1\.xml"/)
        expect(rels).toMatch(/relationships\/slide" Target="\.\.\/slides\/slide\d+\.xml"/)
        expect((await stateOf(deck)).slides[0]?.notes).toBe("Erste\nZweite")
    })

    it("keeps an existing notes part and replaces only the text of its body", async () => {
        const deck = await freshDeck("notes-keep.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", notes: "alt" } ])
        const before = await notesOf(deck, 0)

        /*  an existing notes part with content of its own: a shape, a
            formatted body (bodyPr, list style)  */
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const xml = (await (zip.file(before.part) as JSZip.JSZipObject).async("string"))
            .replace("</p:spTree>", "<p:sp><p:nvSpPr><p:cNvPr id=\"9\" name=\"Eigene Form\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
                + "<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>bleibt</a:t></a:r></a:p></p:txBody></p:sp></p:spTree>")
            .replace(/(<p:ph type="body"[^>]*\/><\/p:nvPr><\/p:nvSpPr><p:spPr\/><p:txBody>)<a:bodyPr\/><a:lstStyle\/>/,
                "$1<a:bodyPr anchor=\"ctr\"/><a:lstStyle><a:lvl1pPr algn=\"r\"/></a:lstStyle>")
        expect(xml).toContain("anchor=\"ctr\"")
        zip.file(before.part, xml)
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))

        await apply(deck, [ { op: "slide.fill", slide: "0", notes: "neu" } ])
        const after = await notesOf(deck, 0)
        expect(after.part).toBe(before.part)
        expect(placeholders(after.doc)).toEqual([ "sldImg:2", "body:3", "sldNum:5" ])
        expect(bodyText(after.doc)).toEqual([ "neu" ])
        expect(elements(after.doc, "p:cNvPr").map((c) => c.getAttribute("name"))).toContain("Eigene Form")
        expect(elements(after.doc, "a:t").map((t) => t.textContent)).toContain("bleibt")
        const body = elements(after.doc, "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === "body") as Element
        expect(firstElement(body, "a:bodyPr")?.getAttribute("anchor")).toBe("ctr")
        expect(firstElement(body, "a:lvl1pPr")?.getAttribute("algn")).toBe("r")
        expect(after.rels).toBe(before.rels)
    })

    it("adds a notes body to an existing notes part without one", async () => {
        const deck = await freshDeck("notes-nobody.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", notes: "weg" } ])
        const before = await notesOf(deck, 0)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const xml = (await (zip.file(before.part) as JSZip.JSZipObject).async("string"))
            .replace(/<p:sp><p:nvSpPr><p:cNvPr id="3"[\s\S]*?<\/p:sp>/, "")
        zip.file(before.part, xml)
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
        expect(placeholders(parseXml(xml))).toEqual([ "sldImg:2", "sldNum:5" ])

        await apply(deck, [ { op: "slide.fill", slide: "0", notes: "wieder da" } ])
        const after = await notesOf(deck, 0)
        expect(after.part).toBe(before.part)
        expect(placeholders(after.doc)).toEqual([ "sldImg:2", "sldNum:5", "body:3" ])
        expect(bodyText(after.doc)).toEqual([ "wieder da" ])
        const ids = elements(after.doc, "p:cNvPr").map((c) => c.getAttribute("id"))
        expect(new Set(ids).size).toBe(ids.length)
    })
})
