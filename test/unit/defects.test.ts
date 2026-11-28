/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the repair-trigger rules of single parts (engine/defects).
*/

import { describe, expect, it } from "vitest"
import { relTargets } from "../../src/engine/parts.js"
import { notesOwner, refDefects, relsDefects, shapeIdDefects, type PartView } from "../../src/engine/defects.js"
import { parseXml } from "../../src/engine/xml.js"

const NS = "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""
    + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
    + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
const RELS = "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\""
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/**  a part view from XML texts  */
const view = (part: string, xml: string | null, rels: string): PartView => ({
    part, doc: xml === null ? null : parseXml(xml), rels: relTargets(part, parseXml(`<Relationships ${RELS}>${rels}</Relationships>`))
})

describe("repair-trigger rules", () => {
    it("ignores the repeated shape ids of an mc:Fallback", () => {
        const sp = "<p:sp><p:nvSpPr><p:cNvPr id=\"5\" name=\"x\"/></p:nvSpPr></p:sp>"
        const slide = (fallback: string): string =>
            `<p:sld ${NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld><p:spTree>`
            + `<mc:AlternateContent><mc:Choice Requires="a14">${sp}</mc:Choice>${fallback}</mc:AlternateContent></p:spTree></p:cSld></p:sld>`
        expect(shapeIdDefects(view("s.xml", slide(`<mc:Fallback>${sp}</mc:Fallback>`), ""))).toEqual([])
        expect(shapeIdDefects(view("s.xml", slide(sp), ""))).toHaveLength(1)
    })

    it("reports references to missing relationships, not empty ones", () => {
        const xml = `<p:sld ${NS}><a:hlinkClick r:id=""/><a:blip r:embed="rId1"/><a:blip r:embed="rId2"/></p:sld>`
        expect(refDefects(view("s.xml", xml, `<Relationship Id="rId1" Type="${T}/image" Target="i.png"/>`)).map((f) => f.text))
            .toEqual([ "s.xml: references missing relationship ids rId2" ])
    })

    it("reports duplicate ids and dead targets of paired Relationship tags", () => {
        const found = relsDefects(view("ppt/slides/slide1.xml", null,
            `<Relationship Id="a" Type="${T}/image" Target="../media/gone.png"></Relationship>`
            + `<Relationship Id="a" Type="${T}/image" Target="../media/ok.png"/>`), (p) => p === "ppt/media/ok.png")
        expect(found.map((f) => [ f.rule, f.text ])).toEqual([
            [ "rels-dup", "ppt/slides/slide1.xml: duplicate relationship ids a" ],
            [ "rels-dead", "ppt/slides/slide1.xml: dead relationship target ppt/media/gone.png" ]
        ])
    })

    it("finds the owner of a notes part past its slide-jump links", () => {
        const notes = view("ppt/notesSlides/notesSlide1.xml", `<p:notes ${NS}><a:hlinkClick r:id="rJump"/></p:notes>`,
            `<Relationship Id="rJump" Type="${T}/slide" Target="../slides/slide9.xml"/>`
            + `<Relationship Id="rOwn" Type="${T}/slide" Target="../slides/slide2.xml"/>`)
        expect(notesOwner(notes)).toBe("ppt/slides/slide2.xml")
    })
})
