/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: input decks carrying PowerPoint repair triggers stay
**  editable. A slide with a trigger (duplicate shape or relationship ids,
**  dead relationships, references to missing relationships, a notes part of
**  another slide) is repaired and reported with W_REPAIRED; triggers in
**  layouts are dropped package-wide; single-quoted XML attributes are no
**  trigger at all.
*/

import { describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions, type ExecuteResult } from "../../src/commands/apply.js"
import { expectIntact } from "../util/integrity.js"
import { verifyFile } from "../../src/engine/verify.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp", "repair")

const opts: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/**  smallest valid 1x1 PNG with a tag after IEND  */
const png = (tag: string): Buffer => Buffer.concat([ Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64"), Buffer.from(tag) ])

/**  a picture shape referencing a relationship id  */
const pic = (id: number, name: string, rid: string): string =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
    + "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>"

/**  a deck with slides A, B, C (with notes), then patched by `craft`  */
const craftDeck = async (name: string, craft: (zip: JSZip, slides: string[]) => Promise<void>): Promise<string> => {
    mkdirSync(TMP, { recursive: true })
    const deck = path.join(TMP, `${name}.pptx`)
    writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
    await executeOps(deck, { ops: [ "A", "B", "C" ].map((t) =>
        ({ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: t } }, notes: `notes ${t}` })) } as never, opts)
    const slides = (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.part)
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const ct = await text(zip, "[Content_Types].xml")
    if (!ct.includes("Extension=\"png\""))
        zip.file("[Content_Types].xml", ct.replace("<Default ", "<Default Extension=\"png\" ContentType=\"image/png\"/><Default "))
    await craft(zip, slides)
    writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
    return deck
}
const text = async (zip: JSZip, part: string): Promise<string> =>
    await (zip.file(part) as JSZip.JSZipObject).async("string")
const relsOf = (part: string): string =>
    `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`
const addRel = async (zip: JSZip, part: string, id: string, type: string, target: string): Promise<void> =>
    void zip.file(relsOf(part), (await text(zip, relsOf(part))).replace("</Relationships>",
        `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/></Relationships>`))
const addShape = async (zip: JSZip, slide: string, xml: string): Promise<void> =>
    void zip.file(slide, (await text(zip, slide)).replace("</p:spTree>", `${xml}</p:spTree>`))
const apply = async (deck: string, ops: unknown[]): Promise<ExecuteResult> =>
    await executeOps(deck, { ops } as never, opts)
const repaired = (res: ExecuteResult): string[] =>
    res.warnings.filter((w) => w.code === "W_REPAIRED").map((w) => w.message)
const notesOf = async (deck: string): Promise<(string | null)[]> =>
    (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.notes)

describe("repair triggers in the input deck", () => {
    it("repairs a duplicate relationship id that the slide uses", async () => {
        const deck = await craftDeck("dup-rel", async (zip, [ a ]) => {
            zip.file("ppt/media/dup.png", png("dup"))
            await addRel(zip, a as string, "rIdDup", "image", "../media/dup.png")
            await addRel(zip, a as string, "rIdDup", "image", "../media/dup.png")
            await addShape(zip, a as string, pic(90, "Bild", "rIdDup"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("duplicate relationship id")
        await expectIntact(deck)
        const state = await readDeckState(await DeckArchive.open(deck))
        expect(state.slides[0]?.shapes.some((s) => s.name === "Bild")).toBe(true)
    })

    it("drops a link to a missing slide together with its hyperlink", async () => {
        const deck = await craftDeck("dead-link", async (zip, [ a ]) => {
            await addRel(zip, a as string, "rIdGone", "slide", "slide99.xml")
            zip.file(a as string, (await text(zip, a as string)).replace(/<a:t>A<\/a:t>/,
                "<a:t>A</a:t></a:r><a:r><a:rPr><a:hlinkClick r:id=\"rIdGone\" action=\"ppaction://hlinksldjump\"/></a:rPr><a:t>!</a:t>"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("dead relationship target")
        await expectIntact(deck)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const slide = (await readDeckState(await DeckArchive.open(deck))).slides[0]?.part as string
        expect(await text(zip, slide)).not.toContain("rIdGone")
    })

    it("drops a placeholder picture fill whose relationship is missing (1.0.11 layout clones)", async () => {
        const deck = await craftDeck("dangling-fill", async (zip, [ a ]) => {
            zip.file(a as string, (await text(zip, a as string)).replace("<p:spPr/>", "<p:spPr></p:spPr>").replace("</p:spPr>",
                "<a:blipFill><a:blip r:embed=\"rIdMissing\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>"))
        })
        expect((await readDeckState(await DeckArchive.open(deck))).slides.length).toBe(3)
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("references missing relationship id rIdMissing")
        await expectIntact(deck)
        expect((await readDeckState(await DeckArchive.open(deck))).slides[0]?.title).toBe("A")
    })

    it("gives a slide its own notes when the owner of a shared notes part is removed", async () => {
        const deck = await craftDeck("shared-notes", async (zip, [ a, b ]) => {
            const relsB = relsOf(b as string)
            const notesA = /Target="(\.\.\/notesSlides\/[^"]+)"/.exec(await text(zip, relsOf(a as string)))?.[1] as string
            zip.file(relsB, (await text(zip, relsB)).replace(/Target="\.\.\/notesSlides\/[^"]+"/, `Target="${notesA}"`))
        })
        const res = await apply(deck, [ { op: "slide.rm", slide: "title:A" } ])
        expect(repaired(res).join(" ")).toContain("belongs to another slide")
        await expectIntact(deck)
        expect(await notesOf(deck)).toEqual([ "notes A", "notes C" ])
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const b = (await readDeckState(await DeckArchive.open(deck))).slides[0]?.part as string
        const notes = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(await text(zip, relsOf(b)))?.[1] as string
        expect(await text(zip, `ppt/notesSlides/_rels/${notes}.rels`)).toContain(`Target="../slides/${path.posix.basename(b)}"`)
    })

    it("repairs a dead relationship in the notes of an untouched slide", async () => {
        const deck = await craftDeck("dead-notes-rel", async (zip, [ a ]) => {
            const notes = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(await text(zip, relsOf(a as string)))?.[1] as string
            await addRel(zip, `ppt/notesSlides/${notes}`, "rIdDead", "image", "../media/none.png")
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("dead relationship target")
        await expectIntact(deck)
        expect(await notesOf(deck)).toEqual([ "notes A", "notes B", "notes C" ])
    })

    it("drops a dead relationship of a layout package-wide", async () => {
        const deck = await craftDeck("dead-layout-rel", async (zip) => {
            const layout = Object.keys(zip.files).find((f) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(f)) as string
            await addRel(zip, layout, "rIdDead", "image", "../media/none.png")
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("slideLayouts")
        await expectIntact(deck)
    })

    it("renumbers duplicate shape ids", async () => {
        const deck = await craftDeck("dup-shape-ids", async (zip, [ a ]) => {
            zip.file("ppt/media/d.png", png("d"))
            await addRel(zip, a as string, "rIdPic", "image", "../media/d.png")
            await addShape(zip, a as string, pic(2, "Doppelt", "rIdPic"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("duplicate cNvPr id")
        await expectIntact(deck)
    })

    it("reads single-quoted attributes without seeing triggers", async () => {
        const deck = await craftDeck("single-quotes", async (zip, [ a ]) => {
            const notes = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(await text(zip, relsOf(a as string)))?.[1] as string
            for (const part of [ `ppt/notesSlides/_rels/${notes}.rels`, relsOf(a as string), "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels" ])
                zip.file(part, (await text(zip, part)).replace(/(<[^?!][^>]*?)>/g, (tag: string) =>
                    tag.replace(/="([^"]*)"/g, "='$1'")))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res)).toEqual([])
        await expectIntact(deck)
        expect(await notesOf(deck)).toEqual([ "notes A", "notes B", "notes C" ])
        const added = await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "D" } }, notes: "notes D" } ])
        expect(added.result.slideCount).toBe(4)
        await expectIntact(deck)
    })

    it("treats show=\"false\" as hidden and keeps it hidden on an edit", async () => {
        const deck = await craftDeck("show-false", async (zip, [ a ]) => {
            zip.file(a as string, (await text(zip, a as string)).replace("<p:sld ", "<p:sld show=\"false\" "))
        })
        expect((await readDeckState(await DeckArchive.open(deck))).slides[0]?.hidden).toBe(true)
        await apply(deck, [ { op: "slide.fill", slide: "title:A", notes: "edited" } ])
        expect((await readDeckState(await DeckArchive.open(deck))).slides[0]?.hidden).toBe(true)
        await expectIntact(deck)
    })

    it("names a missing notes master instead of writing a dead relationship", async () => {
        mkdirSync(TMP, { recursive: true })
        const deck = path.join(TMP, "no-notes-master.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const rels = await text(zip, "ppt/_rels/presentation.xml.rels")
        zip.file("ppt/_rels/presentation.xml.rels", rels.replace(/<Relationship [^>]*\/notesMaster"[^>]*\/>/, ""))
        zip.file("ppt/presentation.xml", (await text(zip, "ppt/presentation.xml")).replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/, ""))
        for (const f of Object.keys(zip.files).filter((n) => n.startsWith("ppt/notesMasters/")))
            zip.remove(f)
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
        const ops = [ { op: "slide.add", layout: "CONTENT", notes: "n" } ]
        await expect(executeOps(deck, { ops } as never, { ...opts, dryRun: true })).rejects.toMatchObject({ code: "E_TEMPLATE" })
        await expect(apply(deck, ops)).rejects.toMatchObject({ code: "E_TEMPLATE" })
    })

    it("removes a SmartArt together with its drawing on el.rm", async () => {
        const DGM = "application/vnd.openxmlformats-officedocument.drawingml.diagram"
        const deck = await craftDeck("smartart-rm", async (zip, [ a ]) => {
            const parts: [ string, string, string ][] = [
                [ "data", `${DGM}Data+xml`, "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                    + "<dgm:ptLst/><dgm:extLst><a:ext uri=\"http://schemas.microsoft.com/office/drawing/2008/diagram\">"
                    + "<dsp:dataModelExt xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" relId=\"rIdDgD\"/></a:ext></dgm:extLst></dgm:dataModel>" ],
                [ "layout", `${DGM}Layout+xml`, "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>" ],
                [ "quickStyle", `${DGM}Style+xml`, "<dgm:styleDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>" ],
                [ "colors", `${DGM}Colors+xml`, "<dgm:colorsDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>" ],
                [ "drawing", "application/vnd.ms-office.drawingml.diagramDrawing+xml", "<dsp:drawing xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\"/>" ]
            ]
            let ct = await text(zip, "[Content_Types].xml")
            for (const [ kind, type, xml ] of parts) {
                zip.file(`ppt/diagrams/${kind}9.xml`, xml)
                ct = ct.replace("</Types>", `<Override PartName="/ppt/diagrams/${kind}9.xml" ContentType="${type}"/></Types>`)
            }
            zip.file("[Content_Types].xml", ct)
            await addRel(zip, a as string, "rIdDgM", "diagramData", "../diagrams/data9.xml")
            await addRel(zip, a as string, "rIdDgL", "diagramLayout", "../diagrams/layout9.xml")
            await addRel(zip, a as string, "rIdDgQ", "diagramQuickStyle", "../diagrams/quickStyle9.xml")
            await addRel(zip, a as string, "rIdDgC", "diagramColors", "../diagrams/colors9.xml")
            zip.file(relsOf(a as string), (await text(zip, relsOf(a as string))).replace("</Relationships>",
                "<Relationship Id=\"rIdDgD\" Type=\"http://schemas.microsoft.com/office/2007/relationships/diagramDrawing\" Target=\"../diagrams/drawing9.xml\"/></Relationships>"))
            await addShape(zip, a as string, "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"95\" name=\"Smart\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
                + "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">"
                + "<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" r:dm=\"rIdDgM\" r:lo=\"rIdDgL\" r:qs=\"rIdDgQ\" r:cs=\"rIdDgC\"/>"
                + "</a:graphicData></a:graphic></p:graphicFrame>")
        })
        await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        await expectIntact(deck)
        await apply(deck, [ { op: "el.rm", slide: "title:A", name: "Smart" } ])
        await expectIntact(deck)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        expect(Object.keys(zip.files).filter((f) => /^ppt\/diagrams\/[^/]+\.xml$/.test(f))).toEqual([])
    })

    it("never names a new part like an existing one in another case", async () => {
        const deck = await craftDeck("case-names", async (zip, [ a ]) => {
            zip.file("ppt/media/IMAGE1.PNG", png("upper"))
            const ct = await text(zip, "[Content_Types].xml")
            zip.file("[Content_Types].xml", ct.replace("<Default ", "<Default Extension=\"PNG\" ContentType=\"image/png\"/><Default "))
            await addRel(zip, a as string, "rIdUpper", "image", "../media/IMAGE1.PNG")
            await addShape(zip, a as string, pic(91, "Gross", "rIdUpper"))
        })
        mkdirSync(TMP, { recursive: true })
        const file = path.join(TMP, "case-names-new.png")
        writeFileSync(file, png("new"))
        await apply(deck, [ { op: "el.add", slide: "title:B", elements: [ { type: "image", name: "Neu", frame: { x: 1, y: 1, w: 1, h: 1 }, path: file } ] } ])
        await expectIntact(deck)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const media = Object.keys(zip.files).filter((f) => f.startsWith("ppt/media/") && !f.endsWith("/"))
        expect(new Set(media.map((m) => m.toLowerCase())).size).toBe(media.length)
        expect((await (zip.file("ppt/media/IMAGE1.PNG") as JSZip.JSZipObject).async("nodebuffer")).toString().endsWith("upper")).toBe(true)
    })
})

/**  the bytes of a part of a deck file  */
const bytesIn = async (deck: string, part: string): Promise<Buffer> =>
    await ((await JSZip.loadAsync(readFileSync(deck))).file(part) as JSZip.JSZipObject).async("nodebuffer")

/**  the notes part a slide's relationships name  */
const notesPartOf = async (zip: JSZip, slide: string): Promise<string> =>
    path.posix.join("ppt/slides", /Target="(\.\.\/notesSlides\/[^"]+)"/.exec(await text(zip, relsOf(slide)))?.[1] as string)

describe("repair triggers beyond the slide itself", () => {
    it("allocates above relationship ids of a single-quoted deck", async () => {
        const deck = await craftDeck("single-quoted-pres-rels", async (zip) => {
            zip.file("ppt/_rels/presentation.xml.rels", (await text(zip, "ppt/_rels/presentation.xml.rels")).replace(/"/g, "'"))
        })
        const res = await apply(deck, [ { op: "slide.copy", slide: "index:0" }, { op: "slide.add", layout: "CONTENT" } ])
        expect(repaired(res)).toEqual([])
        await expectIntact(deck)
        const parts = (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.part)
        expect(parts).toHaveLength(5)
        expect(new Set(parts).size).toBe(5)
    })

    it("drops a master id of a dead presentation relationship", async () => {
        const deck = await craftDeck("dead-handout", async (zip) => {
            await addRel(zip, "ppt/presentation.xml", "rIdHM", "handoutMaster", "handoutMasters/none.xml")
            zip.file("ppt/presentation.xml", (await text(zip, "ppt/presentation.xml")).replace("<p:sldIdLst>",
                "<p:handoutMasterIdLst><p:handoutMasterId r:id=\"rIdHM\"/></p:handoutMasterIdLst><p:sldIdLst>"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("ppt/presentation.xml")
        await expectIntact(deck)
        expect((await bytesIn(deck, "ppt/presentation.xml")).toString()).not.toContain("rIdHM")
    })

    it("drops a layout picture whose relationship is dead", async () => {
        let layout = ""
        const deck = await craftDeck("dead-layout-pic", async (zip) => {
            layout = Object.keys(zip.files).find((f) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(f)) as string
            await addRel(zip, layout, "rIdDead", "image", "../media/none.png")
            await addShape(zip, layout, pic(90, "Deko", "rIdDead"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain(layout)
        await expectIntact(deck)
        expect((await bytesIn(deck, layout)).toString()).not.toContain("rIdDead")
    })

    it("repairs the notes of a copied slide", async () => {
        const deck = await craftDeck("copy-defective-notes", async (zip, [ a ]) => {
            const notes = await notesPartOf(zip, a as string)
            zip.file("ppt/media/n.png", png("n"))
            await addRel(zip, notes, "rIdDup", "image", "../media/n.png")
            await addRel(zip, notes, "rIdDup", "image", "../media/n.png")
        })
        await apply(deck, [ { op: "slide.copy", slide: "index:0" } ])
        await expectIntact(deck)
        expect((await notesOf(deck)).filter((n) => n === "notes A")).toHaveLength(2)
    })

    it("gives a copy of a slide with foreign notes its own notes", async () => {
        const deck = await craftDeck("copy-foreign-notes", async (zip, [ a, b ]) => {
            const notesRels = relsOf(await notesPartOf(zip, a as string))
            zip.file(notesRels, (await text(zip, notesRels)).replace(/\.\.\/slides\/[^"]+/, `../slides/${path.posix.basename(b as string)}`))
        })
        await apply(deck, [ { op: "slide.copy", slide: "index:0" } ])
        await expectIntact(deck)
        expect((await notesOf(deck)).filter((n) => n === "notes A")).toHaveLength(2)
    })

    it("drops only the dead picture of a group", async () => {
        let slide = ""
        const deck = await craftDeck("group-dangling", async (zip, [ a ]) => {
            slide = a as string
            zip.file("ppt/media/ok.png", png("ok"))
            await addRel(zip, slide, "rIdOk", "image", "../media/ok.png")
            await addShape(zip, slide, "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"80\" name=\"Gruppe\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
                + `<p:grpSpPr/>${pic(81, "Gut", "rIdOk")}${pic(82, "Kaputt", "rIdMissing")}</p:grpSp>`)
        })
        await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        await expectIntact(deck)
        const xml = (await bytesIn(deck, slide)).toString()
        expect(xml).toContain("name=\"Gut\"")
        expect(xml).not.toContain("name=\"Kaputt\"")
    })

    it("keeps a picture whose target differs in letter case", async () => {
        let slide = ""
        const deck = await craftDeck("case-target", async (zip, [ a ]) => {
            slide = a as string
            zip.file("ppt/media/Photo1.PNG", png("case"))
            await addRel(zip, slide, "rIdCase", "image", "../media/photo1.png")
            await addShape(zip, slide, pic(90, "Foto", "rIdCase"))
        })
        const before = await bytesIn(deck, slide)
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        expect(repaired(res)).toEqual([])
        await expectIntact(deck)
        expect((await bytesIn(deck, slide)).equals(before)).toBe(true)
        expect((await bytesIn(deck, "ppt/media/Photo1.PNG")).toString().endsWith("case")).toBe(true)
    })

    it("accepts the repeated shape ids of an mc:Fallback", async () => {
        let slide = ""
        const sp = "<p:sp><p:nvSpPr><p:cNvPr id=\"70\" name=\"Formel\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>"
        const deck = await craftDeck("mc-fallback", async (zip, [ a ]) => {
            slide = a as string
            await addShape(zip, slide, "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
                + `<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">${sp}</mc:Choice>`
                + `<mc:Fallback>${sp}</mc:Fallback></mc:AlternateContent>`)
        })
        const before = await bytesIn(deck, slide)
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        expect(repaired(res)).toEqual([])
        await expectIntact(deck)
        expect((await bytesIn(deck, slide)).equals(before)).toBe(true)
    })

    it("drops a background picture whose relationship is missing", async () => {
        let slide = ""
        const deck = await craftDeck("bg-dangling", async (zip, [ a ]) => {
            slide = a as string
            zip.file(slide, (await text(zip, slide)).replace("<p:cSld>", "<p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed=\"rIdGone\"/>"
                + "<a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>"))
        })
        await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        await expectIntact(deck)
        expect((await bytesIn(deck, slide)).toString()).not.toContain("<p:bg>")
    })

    it("writes a title with replacement patterns into the app properties", async () => {
        const deck = await craftDeck("dollar-title", async () => {})
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Revenue in US$'000 $&" } } } ])
        await expectIntact(deck)
        expect((await bytesIn(deck, "docProps/app.xml")).toString()).toContain("Revenue in US$'000 $&amp;")
    })

    it("collects the ink and ActiveX parts of a removed slide", async () => {
        const deck = await craftDeck("ink-activex", async (zip, [ a ]) => {
            zip.file("ppt/ink/ink1.xml", "<inkml:ink xmlns:inkml=\"http://www.w3.org/2003/InkML\"/>")
            zip.file("ppt/activeX/activeX1.xml", "<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\"/>")
            zip.file("[Content_Types].xml", (await text(zip, "[Content_Types].xml")).replace("</Types>",
                "<Override PartName=\"/ppt/ink/ink1.xml\" ContentType=\"application/inkml+xml\"/>"
                + "<Override PartName=\"/ppt/activeX/activeX1.xml\" ContentType=\"application/vnd.ms-office.activeX+xml\"/></Types>"))
            await addRel(zip, a as string, "rIdInk", "customXml", "../ink/ink1.xml")
            await addRel(zip, a as string, "rIdAx", "control", "../activeX/activeX1.xml")
        })
        await apply(deck, [ { op: "slide.rm", slide: "index:0" } ])
        await expectIntact(deck)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        expect(zip.file("ppt/ink/ink1.xml")).toBeNull()
        expect(zip.file("ppt/activeX/activeX1.xml")).toBeNull()
    })
})

/**  the warnings of a code  */
const warned = (res: ExecuteResult, code: string): string[] =>
    res.warnings.filter((w) => w.code === code).map((w) => w.message)

describe("input defects beyond repair and repair granularity", () => {
    it("keeps a slide listed twice as it was and can remove the duplicate", async () => {
        const deck = await craftDeck("listed-twice", async (zip, [ a ]) => {
            await addRel(zip, "ppt/presentation.xml", "rIdTwice", "slide", `slides/${path.posix.basename(a as string)}`)
            zip.file("ppt/presentation.xml", (await text(zip, "ppt/presentation.xml"))
                .replace("</p:sldIdLst>", "<p:sldId id=\"900\" r:id=\"rIdTwice\"/></p:sldIdLst>"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(warned(res, "W_INPUT_DEFECT").join(" ")).toContain("a second time")
        await apply(deck, [ { op: "slide.rm", slide: "id:900" } ])
        await expectIntact(deck)
        expect(await notesOf(deck)).toEqual([ "notes A", "notes B", "notes C" ])
    })

    it("gives duplicate slide ids fresh ones and keeps every slide", async () => {
        const deck = await craftDeck("dup-sldid", async (zip) => {
            const pres = await text(zip, "ppt/presentation.xml")
            const ids = [ ...pres.matchAll(/<p:sldId id="(\d+)"/g) ].map((m) => m[1] as string)
            zip.file("ppt/presentation.xml", pres.replace(`<p:sldId id="${ids[1] as string}"`, `<p:sldId id="${ids[0] as string}"`))
        })
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:2", placeholders: { title: { text: "C2" } } } ])
        expect(repaired(res).join(" ")).toContain("duplicate slide ids")
        await expectIntact(deck)
        expect(await notesOf(deck)).toEqual([ "notes A", "notes B", "notes C" ])
    })

    it("gives notes without back-reference one", async () => {
        const deck = await craftDeck("notes-no-backref", async (zip, [ a ]) => {
            const notesRels = relsOf(await notesPartOf(zip, a as string))
            zip.file(notesRels, (await text(zip, notesRels)).replace(/<Relationship [^>]*\/slide"[^>]*\/>/, ""))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("own notes part")
        await expectIntact(deck)
        expect(await notesOf(deck)).toEqual([ "notes A", "notes B", "notes C" ])
    })

    it("accepts notes whose back-reference a jump link reuses", async () => {
        const deck = await craftDeck("notes-jump-reuse", async (zip, [ a ]) => {
            const notes = await notesPartOf(zip, a as string)
            const rid = /<Relationship [^>]*Id="([^"]+)"[^>]*\/slide"/.exec(await text(zip, relsOf(notes)))?.[1] as string
            zip.file(notes, (await text(zip, notes)).replace("<a:r><a:t>notes A",
                `<a:r><a:rPr><a:hlinkClick r:id="${rid}" action="ppaction://hlinksldjump"/></a:rPr><a:t>notes A`))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res)).toEqual([])
        await expectIntact(deck)
    })

    it("drops a dead relationship of the package", async () => {
        const deck = await craftDeck("dead-thumbnail", async (zip) => {
            zip.remove("docProps/thumbnail.jpeg")
            zip.file("_rels/.rels", (await text(zip, "_rels/.rels")).replace("</Relationships>",
                "<Relationship Id=\"rIdThumb\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\""
                + " Target=\"docProps/thumbnail.jpeg\"/></Relationships>"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("_rels/.rels")
        await expectIntact(deck)
    })

    it("keeps a content-type override that differs in letter case", async () => {
        let slide = ""
        const deck = await craftDeck("ct-case", async (zip, [ a ]) => {
            slide = a as string
            const upper = slide.replace("slides/slide", "slides/Slide")
            zip.file("[Content_Types].xml", (await text(zip, "[Content_Types].xml")).replace(`PartName="/${slide}"`, `PartName="/${upper}"`))
        })
        await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        await expectIntact(deck)
        expect((await bytesIn(deck, "[Content_Types].xml")).toString().toLowerCase()).toContain(`partname="/${slide.toLowerCase()}"`)
    })

    it("drops a dead picture bullet, group fill and picture extension, not their shapes", async () => {
        let slide = ""
        const deck = await craftDeck("granular", async (zip, [ a ]) => {
            slide = a as string
            zip.file("ppt/media/ok.png", png("ok"))
            await addRel(zip, slide, "rIdOk", "image", "../media/ok.png")
            await addShape(zip, slide, "<p:sp><p:nvSpPr><p:cNvPr id=\"60\" name=\"Liste\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr/>"
                + "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buBlip><a:blip r:embed=\"rIdGone1\"/></a:buBlip></a:pPr>"
                + "<a:r><a:t>Punkt</a:t></a:r></a:p></p:txBody></p:sp>"
                + "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"61\" name=\"Gruppe\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
                + "<p:grpSpPr><a:blipFill><a:blip r:embed=\"rIdGone2\"/></a:blipFill></p:grpSpPr>"
                + pic(62, "Innen", "rIdOk") + "</p:grpSp>"
                + pic(63, "Svg", "rIdOk").replace("<a:blip r:embed=\"rIdOk\"/>",
                    "<a:blip r:embed=\"rIdOk\"><a:extLst><a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">"
                    + "<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rIdGone3\"/>"
                    + "</a:ext></a:extLst></a:blip>"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        const message = repaired(res).join(" ")
        expect(message).toContain("a picture bullet")
        expect(message).toContain("a picture fill")
        expect(message).toContain("an extension")
        await expectIntact(deck)
        const xml = (await bytesIn(deck, slide)).toString()
        for (const name of [ "Liste", "Gruppe", "Innen", "Svg" ])
            expect(xml).toContain(`name="${name}"`)
        expect(xml).toContain("Punkt")
    })

    it("keeps a master placeholder whose picture bullet is dead", async () => {
        let master = ""
        const deck = await craftDeck("master-bullet", async (zip) => {
            master = Object.keys(zip.files).find((f) => /^ppt\/slideMasters\/[^/]+\.xml$/.test(f)) as string
            const xml = await text(zip, master)
            expect(xml).toContain("<a:lstStyle/>")
            zip.file(master, xml.replace("<a:lstStyle/>", "<a:lstStyle><a:lvl1pPr><a:buBlip><a:blip r:embed=\"rIdGone\"/></a:buBlip></a:lvl1pPr></a:lstStyle>"))
        })
        const shapes = (xml: string): number => [ ...xml.matchAll(/<p:sp>/g) ].length
        const before = shapes(await text(await JSZip.loadAsync(readFileSync(deck)), master))
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("a picture bullet")
        await expectIntact(deck)
        expect(shapes((await bytesIn(deck, master)).toString())).toBe(before)
    })

    it("drops a dead transition sound with its action, a SmartArt with any dead part as a whole", async () => {
        let slide = ""
        const deck = await craftDeck("sound-smartart", async (zip, [ a ]) => {
            slide = a as string
            zip.file("ppt/media/ok.png", png("ok"))
            await addRel(zip, slide, "rIdOk", "image", "../media/ok.png")
            await addShape(zip, slide, "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"64\" name=\"Diagramm\"/><p:cNvGraphicFramePr/><p:nvPr/>"
                + "</p:nvGraphicFramePr><p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></p:xfrm><a:graphic>"
                + "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">"
                + "<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" r:dm=\"rIdOk\" r:lo=\"rIdGone1\" r:qs=\"rIdOk\" r:cs=\"rIdOk\"/>"
                + "</a:graphicData></a:graphic></p:graphicFrame>")
            zip.file(slide, (await text(zip, slide)).replace("</p:cSld>",
                "</p:cSld><p:transition><p:fade/><p:sndAc><p:stSnd><p:snd r:embed=\"rIdGone2\" name=\"x.wav\"/></p:stSnd></p:sndAc></p:transition>"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("a transition sound")
        await expectIntact(deck)
        const xml = (await bytesIn(deck, slide)).toString()
        expect(xml).toContain("<p:transition>")
        expect(xml).not.toContain("p:sndAc")
        expect(xml).not.toContain("name=\"Diagramm\"")
    })

    it("reads a duplicate presentation relationship id like the engine", async () => {
        const deck = await craftDeck("dup-pres-rel", async (zip) => {
            const pres = await text(zip, "ppt/presentation.xml")
            const rids = [ ...pres.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g) ].map((m) => m[1] as string)
            const rels = await text(zip, "ppt/_rels/presentation.xml.rels")
            zip.file("ppt/_rels/presentation.xml.rels", rels.replace(`Id="${rids[1] as string}"`, `Id="${rids[0] as string}"`))
            zip.file("ppt/presentation.xml", pres.replace(`r:id="${rids[1] as string}"`, `r:id="${rids[0] as string}"`))
        })
        const titles = (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.title)
        expect(titles).toEqual([ "A", "A", "C" ])
        await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect((await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.title)).toEqual(titles)
    })
})

describe("repairs keep what they may not remove", () => {
    it("drops animations with a dead sound without blocking the deck", async () => {
        const deck = await craftDeck("timing-sound", async (zip, [ a ]) => {
            zip.file("ppt/media/a.wav", Buffer.from("RIFF"))
            zip.file("[Content_Types].xml", (await text(zip, "[Content_Types].xml")).replace("<Default ", "<Default Extension=\"wav\" ContentType=\"audio/wav\"/><Default "))
            await addRel(zip, a as string, "rIdSnd", "audio", "../media/a.wav")
            const snd = (rid: string): string => `<p:par><p:cTn id="1"><p:stCondLst><p:cond delay="0"><p:tgtEl><p:sndTgt r:embed="${rid}" name="a.wav"/></p:tgtEl></p:cond></p:stCondLst></p:cTn></p:par>`
            zip.file(a as string, (await text(zip, a as string)).replace("</p:cSld>", `</p:cSld><p:timing><p:tnLst>${snd("rIdSnd")}${snd("rIdGoneSnd")}</p:tnLst></p:timing>`))
        })
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        expect(repaired(res).join(" ")).toContain("the animations")
        await expectIntact(deck)
    })

    it("never names a new part like a dead relationship target", async () => {
        let slide = ""
        const deck = await craftDeck("dead-target-name", async (zip, [ a ]) => {
            slide = a as string
            await addRel(zip, slide, "rIdJump", "slide", "slide4.xml")
            await addShape(zip, slide, "<p:sp><p:nvSpPr><p:cNvPr id=\"70\" name=\"Sprung\"><a:hlinkClick r:id=\"rIdJump\" action=\"ppaction://hlinksldjump\"/></p:cNvPr>"
                + "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>")
        })
        const res = await apply(deck, [ { op: "slide.add", layout: "CONTENT" } ])
        expect(repaired(res).join(" ")).toContain("dead relationship target")
        await expectIntact(deck)
        const parts = (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.part)
        expect(parts).not.toContain("ppt/slides/slide4.xml")
        expect((await bytesIn(deck, relsOf(slide))).toString()).not.toContain("rIdJump")
    })

    it("tolerates malformed custom XML and reads UTF-16 custom XML", async () => {
        const deck = await craftDeck("custom-xml", async (zip) => {
            zip.file("customXml/item1.xml", "<a><b></a>")
            zip.file("customXml/item2.xml", Buffer.concat([ Buffer.from([ 0xFF, 0xFE ]),
                Buffer.from("<?xml version=\"1.0\" encoding=\"UTF-16\"?><root/>", "utf16le") ]))
        })
        const before = await verifyFile(deck)
        expect(before.join(" ")).toContain("customXml/item1.xml: malformed XML")
        expect(before.join(" ")).not.toContain("item2")
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        expect(res.warnings.map((w) => w.code)).toContain("W_INPUT_DEFECT")
        expect(await verifyFile(deck)).toEqual(before)
    })

    it("keeps the slide count of app properties without title list", async () => {
        const deck = await craftDeck("app-no-titles", async (zip) => {
            zip.file("docProps/app.xml", (await text(zip, "docProps/app.xml"))
                .replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, "").replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, ""))
        })
        await apply(deck, [ { op: "slide.add", layout: "CONTENT" } ])
        await expectIntact(deck)
        expect((await bytesIn(deck, "docProps/app.xml")).toString()).toContain("<Slides>4</Slides>")
    })

    it("turns a dead picture fill of a theme style list into a plain fill", async () => {
        let theme = ""
        const deck = await craftDeck("theme-fill", async (zip) => {
            theme = Object.keys(zip.files).find((f) => /^ppt\/theme\/theme\d+\.xml$/.test(f)) as string
            const xml = await text(zip, theme)
            const lst = /<a:fillStyleLst>([\s\S]*?)<\/a:fillStyleLst>/.exec(xml)?.[0] as string
            zip.file(theme, xml.replace(lst, lst.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, `<a:blipFill><a:blip xmlns:r="${REL}" r:embed="rIdGone"/></a:blipFill>`)))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("a picture fill style")
        await expectIntact(deck)
        const lst = /<a:fillStyleLst>([\s\S]*?)<\/a:fillStyleLst>/.exec((await bytesIn(deck, theme)).toString())?.[1] as string
        expect([ ...lst.matchAll(/<a:(solidFill|gradFill|blipFill|pattFill|noFill)\b/g) ].length).toBe(3)
    })

    it("edits a slide listed twice", async () => {
        const deck = await craftDeck("listed-twice-edit", async (zip, [ a ]) => {
            await addRel(zip, "ppt/presentation.xml", "rIdTwice", "slide", `slides/${path.posix.basename(a as string)}`)
            zip.file("ppt/presentation.xml", (await text(zip, "ppt/presentation.xml"))
                .replace("</p:sldIdLst>", "<p:sldId id=\"900\" r:id=\"rIdTwice\"/></p:sldIdLst>"))
        })
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:0", placeholders: { title: { text: "A2" } } } ])
        expect(warned(res, "W_INPUT_DEFECT").join(" ")).toContain("a second time")
        expect((await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.title)).toEqual([ "A2", "B", "C", "A2" ])
    })

    it("drops a slide id pointing at an external relationship", async () => {
        const deck = await craftDeck("sldid-external", async (zip) => {
            zip.file("ppt/_rels/presentation.xml.rels", (await text(zip, "ppt/_rels/presentation.xml.rels")).replace("</Relationships>",
                `<Relationship Id="rIdExt" Type="${REL}/slide" Target="https://example.com/x.xml" TargetMode="External"/></Relationships>`))
            zip.file("ppt/presentation.xml", (await text(zip, "ppt/presentation.xml"))
                .replace("</p:sldIdLst>", "<p:sldId id=\"901\" r:id=\"rIdExt\"/></p:sldIdLst>"))
        })
        const res = await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        expect(repaired(res).join(" ")).toContain("without a slide part")
        await expectIntact(deck)
    })

    it("resolves a percent-encoded target", async () => {
        let slide = ""
        const deck = await craftDeck("iri-target", async (zip, [ a ]) => {
            slide = a as string
            zip.file("ppt/media/my image.png", png("iri"))
            await addRel(zip, slide, "rIdIri", "image", "../media/my%20image.png")
            await addShape(zip, slide, pic(90, "Leer", "rIdIri"))
        })
        const before = await bytesIn(deck, slide)
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        expect(repaired(res)).toEqual([])
        expect((await bytesIn(deck, slide)).equals(before)).toBe(true)
        expect((await JSZip.loadAsync(readFileSync(deck))).file("ppt/media/my image.png")).not.toBeNull()
    })
})

/**  a slide with a chart whose XML references a relationship `craft` breaks  */
const chartDeck = async (name: string, rels: string): Promise<string> =>
    await craftDeck(name, async (zip, [ a ]) => {
        zip.file("ppt/charts/chart91.xml", `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL}">`
            + "<c:chart><c:plotArea><c:layout/></c:plotArea></c:chart><c:externalData r:id=\"rIdWb\"/></c:chartSpace>")
        if (rels !== "")
            zip.file("ppt/charts/_rels/chart91.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`)
        zip.file("[Content_Types].xml", (await text(zip, "[Content_Types].xml")).replace("</Types>",
            "<Override PartName=\"/ppt/charts/chart91.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/></Types>"))
        await addRel(zip, a as string, "rIdChart", "chart", "../charts/chart91.xml")
        await addShape(zip, a as string, "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"95\" name=\"Diagramm\"/><p:cNvGraphicFramePr/><p:nvPr/>"
            + "</p:nvGraphicFramePr><p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></p:xfrm><a:graphic>"
            + "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
            + "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rIdChart\"/></a:graphicData></a:graphic></p:graphicFrame>")
    })

describe("repairs inside charts and package metadata", () => {
    it("repairs a chart without its workbook relationship, also in a copy", async () => {
        const deck = await chartDeck("chart-no-wb-rel", "")
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        expect(repaired(res).join(" ")).toContain("ppt/charts/chart91.xml")
        await expectIntact(deck)
        const copy = await chartDeck("chart-no-wb-rel-copy", "")
        await apply(copy, [ { op: "slide.copy", slide: "index:0" } ])
        await expectIntact(copy)
    })

    it("repairs a chart whose workbook is missing", async () => {
        const deck = await chartDeck("chart-dead-wb",
            `<Relationship Id="rIdWb" Type="${REL}/package" Target="../embeddings/none.xlsx"/>`)
        await apply(deck, [ { op: "slide.fill", slide: "index:1", placeholders: { title: { text: "B2" } } } ])
        await expectIntact(deck)
        expect((await bytesIn(deck, "ppt/charts/chart91.xml")).toString()).not.toContain("externalData")
    })

    it("writes UTF-16 app properties back as declared UTF-8", async () => {
        const deck = await craftDeck("app-utf16", async (zip) => {
            const xml = (await text(zip, "docProps/app.xml")).replace(/encoding="UTF-8"/, "encoding=\"UTF-16\"")
            zip.file("docProps/app.xml", Buffer.concat([ Buffer.from([ 0xFF, 0xFE ]), Buffer.from(xml, "utf16le") ]))
        })
        await apply(deck, [ { op: "slide.add", layout: "CONTENT" } ])
        await expectIntact(deck)
        const app = (await bytesIn(deck, "docProps/app.xml")).toString("utf8")
        expect(app).toContain("encoding=\"UTF-8\"")
        expect(app).toContain("<Slides>4</Slides>")
    })

    it("keeps the custom-show titles after the slide titles", async () => {
        const deck = await craftDeck("custom-show-titles", async (zip) => {
            const pres = await text(zip, "ppt/presentation.xml")
            const rid = /<p:sldId [^>]*r:id="([^"]+)"/.exec(pres)?.[1] as string
            zip.file("ppt/presentation.xml", pres.replace(/(<p:notesSz[^>]*\/>)/,
                `$1<p:custShowLst><p:custShow name="Kurz" id="0"><p:sldLst><p:sld r:id="${rid}"/></p:sldLst></p:custShow></p:custShowLst>`))
            const app = await text(zip, "docProps/app.xml")
            const hp = /<HeadingPairs><vt:vector size="(\d+)"[^>]*>([\s\S]*?)<\/vt:vector><\/HeadingPairs>/.exec(app) as RegExpExecArray
            const tp = /<TitlesOfParts><vt:vector size="(\d+)"[^>]*>([\s\S]*?)<\/vt:vector><\/TitlesOfParts>/.exec(app) as RegExpExecArray
            zip.file("docProps/app.xml", app
                .replace(hp[0], `<HeadingPairs><vt:vector size="${Number(hp[1]) + 2}" baseType="variant">${hp[2] as string}`
                    + "<vt:variant><vt:lpstr>Zielgruppenorientierte Präsentationen</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>")
                .replace(tp[0], `<TitlesOfParts><vt:vector size="${Number(tp[1]) + 1}" baseType="lpstr">${tp[2] as string}<vt:lpstr>Kurz</vt:lpstr></vt:vector></TitlesOfParts>`))
        })
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "D" } } } ])
        await expectIntact(deck)
        const titles = [ ...(await bytesIn(deck, "docProps/app.xml")).toString().matchAll(/<vt:lpstr>([^<]*)<\/vt:lpstr>/g) ].map((m) => m[1])
        expect(titles.slice(-5)).toEqual([ "A", "B", "C", "D", "Kurz" ])
    })

    it("leaves an untouched slide with malformed relationships alone", async () => {
        let slide = ""
        const deck = await craftDeck("malformed-rels", async (zip, [ , b ]) => {
            slide = b as string
            zip.file(relsOf(slide), (await text(zip, relsOf(slide))).replace("</Relationships>", "<Relationship></Relationships>"))
        })
        const before = await bytesIn(deck, relsOf(slide))
        const res = await apply(deck, [ { op: "slide.fill", slide: "index:0", placeholders: { title: { text: "A2" } } } ])
        expect(res.warnings.map((w) => w.code)).toContain("W_INPUT_DEFECT")
        expect((await bytesIn(deck, relsOf(slide))).equals(before)).toBe(true)
    })

    it("writes a percent-escaped target of a copy escaped", async () => {
        const deck = await craftDeck("iri-copy", async (zip, [ a ]) => {
            zip.file("ppt/media/my image.png", png("iri"))
            await addRel(zip, a as string, "rIdIri", "image", "../media/my%20image.png")
            await addShape(zip, a as string, pic(90, "Leer", "rIdIri"))
        })
        await apply(deck, [ { op: "slide.copy", slide: "index:0" } ])
        await expectIntact(deck)
        const parts = (await readDeckState(await DeckArchive.open(deck))).slides.map((s) => s.part)
        const copyRels = (await bytesIn(deck, relsOf(parts[1] as string))).toString()
        expect(copyRels).toContain("my%20image.png")
    })
})
