/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the OPC package layer -- allocator, part names, relationships.
*/

import { describe, expect, it } from "vitest"
import JSZip from "jszip"
import { PartAllocator, findPart, relTargets, relsOwner } from "../../src/engine/parts.js"
import type { PartView } from "../../src/engine/defects.js"
import { parseXml } from "../../src/engine/xml.js"

const RELS = "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\""
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/**  a part view from XML texts  */
const view = (part: string, xml: string | null, rels: string): PartView => ({
    part, doc: xml === null ? null : parseXml(xml), rels: relTargets(part, parseXml(`<Relationships ${RELS}>${rels}</Relationships>`))
})

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
