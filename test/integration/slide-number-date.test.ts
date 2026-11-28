/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the fill fields `slideNumber` and `date` clone the
**  layout's slide-number and date placeholders with their field (true) or
**  remove them (false), like `footer` does for the footer; `pptc footer`
**  takes --number and --date. The clones go in front of the shape tree's
**  extension list and get relationships of their own for what the layout
**  content references.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import JSZip from "jszip"
import { cmdFooter } from "../../src/commands/sugar.js"
import { relTargets, relsPartOf } from "../../src/engine/parts.js"
import { elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { expectIntact } from "../util/integrity.js"
import { apply, craftSlide, failure, freshDeck, slideDoc, stateOf } from "../util/elements.js"

let deck: string

/**  the footer-band placeholder types on a slide, in tree order  */
const band = async (index: number): Promise<string[]> =>
    elements(await slideDoc(deck, index), "p:ph").map((ph) => ph.getAttribute("type") ?? "")
        .filter((t) => [ "ftr", "sldNum", "dt" ].includes(t))

/**  the extension list a PowerPoint slide carries in its shape tree  */
const EXT = "<p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\">"
    + "<p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"1\"/></p:ext></p:extLst>"

beforeAll(async () => {
    deck = await freshDeck("slide-number-date.pptx")
    await apply(deck, [
        { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Eins" } }, slideNumber: false, date: false },
        { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Zwei" } } }
    ])
})

describe("slide number and date", () => {
    it("removes them on false", async () => {
        expect(await band(0)).toEqual([ "ftr" ])
        expect((await band(1)).sort()).toEqual([ "dt", "ftr", "sldNum" ])
    })

    it("clones them with their field on true, in front of the extension list", async () => {
        await craftSlide(deck, 0, (xml) => xml.replace("</p:spTree>", `${EXT}</p:spTree>`))
        await apply(deck, [ { op: "slide.fill", slide: "0", slideNumber: true, date: true } ])
        expect(await band(0)).toEqual([ "ftr", "sldNum", "dt" ])
        const doc = await slideDoc(deck, 0)
        expect(elements(doc, "a:fld").map((f) => f.getAttribute("type"))).toEqual([ "slidenum", expect.stringMatching(/^datetime/) ])
        const spTree = firstElement(doc, "p:spTree")
        expect(spTree?.lastChild?.nodeName).toBe("p:extLst")
    })

    it("gives the clone relationships of its own", async () => {
        /*  a link in the layout's slide-number placeholder  */
        const layout = (await JSZip.loadAsync(readFileSync(deck)))
        const slidePart = (await stateOf(deck)).slides[1]?.part ?? ""
        const slideRels = relTargets(slidePart, parseXml(await (layout.file(relsPartOf(slidePart)) as JSZip.JSZipObject).async("string")))
        const layoutPart = slideRels.find((r) => r.kind === "slideLayout")?.target ?? ""
        const xml = await (layout.file(layoutPart) as JSZip.JSZipObject).async("string")
        const at = xml.indexOf("type=\"sldNum\"")
        const fld = xml.indexOf("<a:fld", at)
        layout.file(layoutPart, `${xml.slice(0, fld)}<a:r><a:rPr lang="de-DE"><a:hlinkClick r:id="rIdPptcTest"/></a:rPr><a:t>Seite</a:t></a:r>${xml.slice(fld)}`)
        const rels = await (layout.file(relsPartOf(layoutPart)) as JSZip.JSZipObject).async("string")
        layout.file(relsPartOf(layoutPart), rels.replace("</Relationships>",
            "<Relationship Id=\"rIdPptcTest\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
            + "Target=\"https://example.com/seite\" TargetMode=\"External\"/></Relationships>"))
        writeFileSync(deck, await layout.generateAsync({ type: "nodebuffer" }))

        const warnings = await apply(deck, [ { op: "slide.fill", slide: "1", slideNumber: true } ])
        expect(warnings).toEqual([])
        const doc = await slideDoc(deck, 1)
        const id = elements(doc, "a:hlinkClick")[0]?.getAttribute("r:id") ?? ""
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const own = relTargets(slidePart, parseXml(await (zip.file(relsPartOf(slidePart)) as JSZip.JSZipObject).async("string")))
        expect(own.find((r) => r.id === id)?.el.getAttribute("Target")).toBe("https://example.com/seite")
    })

    it("pptc footer takes --number and --date", async () => {
        await apply(deck, [ { op: "slide.fill", slide: "1", slideNumber: false, date: false } ])
        await cmdFooter([ deck, "--number", "--date" ])
        await expectIntact(deck)
        expect((await band(1)).sort()).toEqual([ "dt", "ftr", "sldNum" ])
        await cmdFooter([ deck, "Fuß", "--slide", "0" ])
        expect((await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.placeholderKind === "footer")?.text).toBe("Fuß")
        await expect(cmdFooter([ deck ])).rejects.toMatchObject({ code: "E_USAGE" })
    })

    it("rejects a non-boolean", async () => {
        expect(await failure(deck, [ { op: "slide.fill", slide: "0", slideNumber: "yes" } ])).toBe("E_SCHEMA")
    })
})
