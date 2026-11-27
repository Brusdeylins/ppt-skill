/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the OPC package layer (allocator, part names, relationships), the
**  repair-trigger rules of single parts and the engine's byte guard.
*/

import { describe, expect, it } from "vitest"
import JSZip from "jszip"
import { PartAllocator, findPart, relTargets, relsOwner } from "../../src/engine/parts.js"
import { notesOwner, refDefects, relsDefects, shapeIdDefects, type PartView } from "../../src/engine/defects.js"
import { byteGuard } from "../../src/engine/inplace/engine.js"
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

/**  a zip reloaded from its bytes  */
const reload = async (zip: JSZip): Promise<JSZip> =>
    await JSZip.loadAsync(await zip.generateAsync({ type: "nodebuffer" }))

describe("PartAllocator", () => {
    it("continues above every name of a family in any case", async () => {
        const zip = new JSZip()
        zip.file("ppt/slides/Slide3.XML", "<x/>")
        zip.file("ppt/slides/slide1.xml", "<x/>")
        const alloc = await PartAllocator.create(zip, [ 256, 300 ])
        expect(alloc.part("ppt/slides", "slide", ".xml")).toBe("ppt/slides/slide4.xml")
        expect(alloc.part("ppt/slides", "slide", ".xml")).toBe("ppt/slides/slide5.xml")
        expect(alloc.like("ppt/media/photo.png")).toBe("ppt/media/image1.png")
        expect(alloc.slideId()).toBe(301)
    })

    it("continues above relationship ids in any quoting", async () => {
        const zip = new JSZip()
        zip.file("ppt/_rels/presentation.xml.rels", `<Relationships ${RELS}><Relationship Id='rIdPptc9005' Type='x' Target='y'/></Relationships>`)
        expect((await PartAllocator.create(zip, [])).relId()).toBe("rIdPptc9006")
    })
})

describe("part names and relationships", () => {
    it("finds a part case-insensitively", () => {
        const zip = new JSZip()
        zip.file("ppt/media/Photo1.PNG", "x")
        expect(findPart(zip, "ppt/media/photo1.png")).toBe("ppt/media/Photo1.PNG")
        expect(findPart(zip, "ppt/media/photo2.png")).toBeNull()
    })

    it("names the owner of a rels part", () => {
        expect(relsOwner("_rels/.rels")).toBe("")
        expect(relsOwner("ppt/slides/_rels/slide1.xml.rels")).toBe("ppt/slides/slide1.xml")
        expect(relsOwner("ppt/_rels/presentation.xml.rels")).toBe("ppt/presentation.xml")
    })

    it("resolves relative, absolute and external targets", () => {
        const rels = view("ppt/slides/slide1.xml", null,
            `<Relationship Id="a" Type="${T}/image" Target="../media/i.png"/>`
            + `<Relationship Id="b" Type="${T}/image" Target="/ppt/media/j.png"/>`
            + `<Relationship Id="c" Type="${T}/hyperlink" Target="https://x" TargetMode="External"/>`).rels
        expect(rels.map((r) => r.target)).toEqual([ "ppt/media/i.png", "ppt/media/j.png", null ])
    })
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
        expect(refDefects(view("s.xml", xml, `<Relationship Id="rId1" Type="${T}/image" Target="i.png"/>`)))
            .toEqual([ "s.xml: references missing relationship ids rId2" ])
    })

    it("reports duplicate ids and dead targets of paired Relationship tags", () => {
        const found = relsDefects(view("ppt/slides/slide1.xml", null,
            `<Relationship Id="a" Type="${T}/image" Target="../media/gone.png"></Relationship>`
            + `<Relationship Id="a" Type="${T}/image" Target="../media/ok.png"/>`), (p) => p === "ppt/media/ok.png")
        expect(found).toEqual([
            "ppt/slides/slide1.xml: duplicate relationship ids a",
            "ppt/slides/slide1.xml: dead relationship target ../media/gone.png"
        ])
    })

    it("finds the owner of a notes part past its slide-jump links", () => {
        const notes = view("ppt/notesSlides/notesSlide1.xml", `<p:notes ${NS}><a:hlinkClick r:id="rJump"/></p:notes>`,
            `<Relationship Id="rJump" Type="${T}/slide" Target="../slides/slide9.xml"/>`
            + `<Relationship Id="rOwn" Type="${T}/slide" Target="../slides/slide2.xml"/>`)
        expect(notesOwner(notes)).toBe("ppt/slides/slide2.xml")
    })
})

describe("byte guard", () => {
    it("fires when an untouched part changes or is rewritten", async () => {
        const src = new JSZip()
        src.file("ppt/slides/slide1.xml", "<a/>")
        const pkg = await reload(src)
        const untouched = new Map([ [ "ppt/slides/slide1.xml", pkg.file("ppt/slides/slide1.xml") ] ])
        expect(await byteGuard(pkg, untouched, await reload(pkg))).toEqual([])

        const other = new JSZip()
        other.file("ppt/slides/slide1.xml", "<b/>")
        expect(await byteGuard(pkg, untouched, await reload(other))).toHaveLength(1)

        pkg.file("ppt/slides/slide1.xml", "<a/>")
        expect(await byteGuard(pkg, untouched, await reload(pkg))).toHaveLength(1)
    })
})
