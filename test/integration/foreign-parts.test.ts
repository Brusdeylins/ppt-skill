/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: parts that PowerPoint-authored decks carry and pptc never
**  creates itself -- embedded OLE objects, slide comments, theme overrides
**  of charts and slides. An apply that does not touch them must keep each
**  of them with its own slide, byte for byte.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { expectIntact } from "../util/integrity.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

const apply = async (deck: string, ops: unknown[]): Promise<void> => {
    await executeOps(deck, { ops } as never, opts)
}
const text = async (zip: JSZip, part: string): Promise<string> =>
    await (zip.file(part) as JSZip.JSZipObject).async("string")
const relsOf = (part: string): string =>
    `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`

/**  add a relationship to a part's rels file  */
const addRel = async (zip: JSZip, part: string, id: string, type: string, target: string): Promise<void> => {
    const rels = relsOf(part)
    zip.file(rels, (await text(zip, rels)).replace("</Relationships>",
        `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/></Relationships>`))
}

/**  register a content-type override  */
const addOverride = async (zip: JSZip, part: string, type: string): Promise<void> => {
    zip.file("[Content_Types].xml", (await text(zip, "[Content_Types].xml")).replace("</Types>",
        `<Override PartName="/${part}" ContentType="${type}"/></Types>`))
}

/**  the part a slide (by presentation order) relates to via a rel type  */
const relatedPart = async (deck: string, index: number, type: string): Promise<{ part: string, bytes: Buffer } | null> => {
    const slide = (await readDeckState(await DeckArchive.open(deck))).slides[index] as { part: string }
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const rel = [ ...(await text(zip, relsOf(slide.part))).matchAll(/<Relationship\b[^>]*>/g) ]
        .map((m) => m[0]).find((r) => r.includes(`/${type}"`))
    if (rel === undefined)
        return null
    const part = path.posix.normalize(path.posix.join("ppt/slides", /Target="([^"]+)"/.exec(rel)?.[1] ?? ""))
    return { part, bytes: await (zip.file(part) as JSZip.JSZipObject).async("nodebuffer") }
}

/**  a deck with two titled slides, then patched by `craft`  */
const craftDeck = async (file: string, craft: (zip: JSZip, slides: string[]) => Promise<void>): Promise<void> => {
    writeFileSync(file, await buildEmptyDeck(TEMPLATE))
    await apply(file, [
        { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "A" } } },
        { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "B" } } }
    ])
    const slides = (await readDeckState(await DeckArchive.open(file))).slides.map((s) => (s as { part: string }).part)
    const zip = await JSZip.loadAsync(readFileSync(file))
    await craft(zip, slides)
    writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }))
}

beforeAll(() => {
    mkdirSync(TMP, { recursive: true })
})

describe("parts of PowerPoint-authored decks", () => {
    it("keeps each slide's embedded OLE object", async () => {
        const deck = path.join(TMP, "foreign-ole.pptx")
        await craftDeck(deck, async (zip, slides) => {
            for (const [ i, slide ] of slides.entries()) {
                const bin = `ppt/embeddings/oleObject${i + 1}.bin`
                zip.file(bin, Buffer.from(`payload-${i}`))
                await addOverride(zip, bin, "application/vnd.openxmlformats-officedocument.oleObject")
                await addRel(zip, slide, "rId90", "oleObject", `../embeddings/oleObject${i + 1}.bin`)
                zip.file(slide, (await text(zip, slide)).replace("</p:spTree>",
                    "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"90\" name=\"Objekt\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
                    + "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></p:xfrm>"
                    + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\">"
                    + "<p:oleObj progId=\"Package\" r:id=\"rId90\"><p:embed/></p:oleObj>"
                    + "</a:graphicData></a:graphic></p:graphicFrame></p:spTree>"))
            }
        })
        for (let round = 0; round < 2; round++) {
            await apply(deck, [ { op: "meta.props", set: { title: `Runde ${round}` } } ])
            for (const i of [ 0, 1 ])
                expect((await relatedPart(deck, i, "oleObject"))?.bytes.toString()).toBe(`payload-${i}`)
        }
        await expectIntact(deck)
    })

    it("keeps slide comments and theme overrides on every apply", async () => {
        const deck = path.join(TMP, "foreign-implicit.pptx")
        await craftDeck(deck, async (zip, slides) => {
            const b = slides[1] as string
            zip.file("ppt/comments/comment1.xml",
                "<p:cmLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>")
            await addOverride(zip, "ppt/comments/comment1.xml", "application/vnd.openxmlformats-officedocument.presentationml.comments+xml")
            await addRel(zip, b, "rId91", "comments", "../comments/comment1.xml")
            zip.file("ppt/theme/themeOverride2.xml",
                "<a:themeOverride xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><!-- slide B --></a:themeOverride>")
            await addOverride(zip, "ppt/theme/themeOverride2.xml", "application/vnd.openxmlformats-officedocument.themeOverride+xml")
            await addRel(zip, b, "rId92", "themeOverride", "../theme/themeOverride2.xml")
        })
        await apply(deck, [ { op: "slide.fill", slide: "title:A", notes: "x" } ])
        await apply(deck, [ { op: "meta.props", set: { title: "t" } } ])
        expect((await relatedPart(deck, 1, "comments"))?.part).toBe("ppt/comments/comment1.xml")
        expect((await relatedPart(deck, 1, "themeOverride"))?.bytes.toString()).toContain("slide B")
        await expectIntact(deck)

        /*  removing the commented slide also removes its comments part  */
        await apply(deck, [ { op: "slide.rm", slide: "title:B" } ])
        expect((await JSZip.loadAsync(readFileSync(deck))).file("ppt/comments/comment1.xml")).toBeNull()
        await expectIntact(deck)
    })

    it("keeps a chart's theme override apart from a slide's theme override", async () => {
        const deck = path.join(TMP, "foreign-chart-override.pptx")
        await craftDeck(deck, async (zip, slides) => {
            zip.file("ppt/theme/themeOverride2.xml",
                "<a:themeOverride xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><!-- slide B --></a:themeOverride>")
            await addOverride(zip, "ppt/theme/themeOverride2.xml", "application/vnd.openxmlformats-officedocument.themeOverride+xml")
            await addRel(zip, slides[1] as string, "rId92", "themeOverride", "../theme/themeOverride2.xml")
        })
        await apply(deck, [ { op: "el.add", slide: "title:A", elements: [ { type: "chart", name: "D", frame: { x: 1, y: 1.5, w: 4, h: 3 },
            data: { type: "column", categories: [ "Q1", "Q2" ], series: [ { name: "S", values: [ 1, 2 ] } ] } } ] } ])
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const chart = Object.keys(zip.files).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f)) as string
        zip.file("ppt/theme/themeOverride1.xml",
            "<a:themeOverride xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><!-- chart --></a:themeOverride>")
        await addOverride(zip, "ppt/theme/themeOverride1.xml", "application/vnd.openxmlformats-officedocument.themeOverride+xml")
        await addRel(zip, chart, "rId93", "themeOverride", "../theme/themeOverride1.xml")
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))

        for (let round = 0; round < 2; round++) {
            await apply(deck, [ { op: "meta.props", set: { title: `Runde ${round}` } } ])
            expect((await relatedPart(deck, 1, "themeOverride"))?.bytes.toString()).toContain("slide B")
        }
        await expectIntact(deck)
    })

    it.each([ false, true ])("fills the placeholder, not an earlier shape with the same name (creationIds: %s)", async (withIds) => {
        const deck = path.join(TMP, `foreign-names-${String(withIds)}.pptx`)
        const cid = (id: number): string => withIds
            ? `<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0000-0000-0000-0000000000${id}}"/></a:ext></a:extLst>`
            : ""
        const sp = (id: number, name: string, txt: string): string =>
            `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}">${cid(id)}</p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
            + "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm></p:spPr>"
            + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${txt}</a:t></a:r></a:p></p:txBody></p:sp>`
        let bodyName = ""
        await craftDeck(deck, async (zip, slides) => {
            const a = slides[0] as string
            const xml = await text(zip, a)
            bodyName = [ ...xml.matchAll(/<p:cNvPr id="\d+" name="([^"]+)"\/>(?:(?!<\/p:nvSpPr>)[\s\S])*<p:ph [^>]*idx="1"/g) ][0]?.[1] ?? ""

            /*  a same-named textbox and a group child named "Box" before the real top-level "Box";
                the placeholders get creation ids too, as PowerPoint writes them  */
            const withPhIds = withIds ? xml.replace(/<p:cNvPr id="(\d+)" name="([^"]+)"\/>/g,
                (_m, id: string, name: string) => `<p:cNvPr id="${id}" name="${name}">${cid(Number(id))}</p:cNvPr>`) : xml
            const group = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"80\" name=\"Gruppe\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
                + "<p:grpSpPr/>" + sp(81, "Box", "GROUP-CHILD") + "</p:grpSp>"
            zip.file(a, withPhIds.replace("</p:grpSpPr>", `</p:grpSpPr>${sp(78, bodyName, "KEEP-ME")}${group}${sp(82, "Box", "TOP")}`))
        })
        expect(bodyName).not.toBe("")
        await apply(deck, [
            { op: "slide.fill", slide: "title:A", placeholders: { body: { text: "NEW-BODY" } } },
            { op: "el.set", slide: "title:A", name: "Box", text: "TOP-NEW" }
        ])
        const slide = (await readDeckState(await DeckArchive.open(deck))).slides[0] as { part: string }
        const xml = await text(await JSZip.loadAsync(readFileSync(deck)), slide.part)
        expect(xml).toContain("KEEP-ME")
        expect(xml).toContain("NEW-BODY")
        expect(xml).toContain("GROUP-CHILD")
        expect(xml).toContain("TOP-NEW")
        await expectIntact(deck)
    })

    it("keeps SmartArt editable and its pictures on every apply", async () => {
        const deck = path.join(TMP, "foreign-smartart.pptx")
        const NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
        const CT = "application/vnd.openxmlformats-officedocument.drawingml."
        await craftDeck(deck, async (zip, slides) => {
            zip.file("[Content_Types].xml", (await text(zip, "[Content_Types].xml")).replace("<Default ",
                "<Default Extension=\"png\" ContentType=\"image/png\"/><Default "))
            for (const [ i, slide ] of slides.entries()) {
                const n = i + 1
                const part = (kind: string): string => `ppt/diagrams/${kind}${n}.xml`
                zip.file(`ppt/media/smart${n}.png`, Buffer.from(`smart-picture-${n}`))
                zip.file(part("data"), `<dgm:dataModel xmlns:dgm="${NS_DGM}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`
                    + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><dgm:ptLst/>"
                    + "<dgm:bg><a:blipFill><a:blip r:embed=\"rId1\"/></a:blipFill></dgm:bg><dgm:extLst><a:ext uri=\"x\">"
                    + "<dsp:dataModelExt xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" relId=\"rId85\" minVer=\"x\"/>"
                    + "</a:ext></dgm:extLst></dgm:dataModel>")
                zip.file(`ppt/diagrams/_rels/data${n}.xml.rels`, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + `<Relationship Id="rId1" Type="${REL}/image" Target="../media/smart${n}.png"/></Relationships>`)
                for (const kind of [ "layout", "quickStyle", "colors" ])
                    zip.file(part(kind), `<dgm:${kind === "layout" ? "layoutDef" : kind === "colors" ? "colorsDef" : "styleDef"} xmlns:dgm="${NS_DGM}"/>`)
                zip.file(part("drawing"), "<dsp:drawing xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\"><dsp:spTree/></dsp:drawing>")
                await addOverride(zip, part("data"), `${CT}diagramData+xml`)
                await addOverride(zip, part("layout"), `${CT}diagramLayout+xml`)
                await addOverride(zip, part("quickStyle"), `${CT}diagramStyle+xml`)
                await addOverride(zip, part("colors"), `${CT}diagramColors+xml`)
                await addOverride(zip, part("drawing"), "application/vnd.ms-office.drawingml.diagramDrawing+xml")
                await addRel(zip, slide, "rId81", "diagramData", `../diagrams/data${n}.xml`)
                await addRel(zip, slide, "rId82", "diagramLayout", `../diagrams/layout${n}.xml`)
                await addRel(zip, slide, "rId83", "diagramQuickStyle", `../diagrams/quickStyle${n}.xml`)
                await addRel(zip, slide, "rId84", "diagramColors", `../diagrams/colors${n}.xml`)
                const rels = relsOf(slide)
                zip.file(rels, (await text(zip, rels)).replace("</Relationships>",
                    `<Relationship Id="rId85" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing${n}.xml"/></Relationships>`))
                zip.file(slide, (await text(zip, slide)).replace("</p:spTree>",
                    "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"91\" name=\"Diagramm\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
                    + "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></p:xfrm>"
                    + `<a:graphic><a:graphicData uri="${NS_DGM}"><dgm:relIds xmlns:dgm="${NS_DGM}" r:dm="rId81" r:lo="rId82" r:qs="rId83" r:cs="rId84"/>`
                    + "</a:graphicData></a:graphic></p:graphicFrame></p:spTree>"))
            }
        })
        for (let round = 0; round < 3; round++) {
            await apply(deck, [ { op: "meta.props", set: { title: `Runde ${round}` } } ])
            const zip = await JSZip.loadAsync(readFileSync(deck))
            for (const i of [ 0, 1 ]) {
                const data = (await relatedPart(deck, i, "diagramData"))?.part as string
                const drawing = await relatedPart(deck, i, "diagramDrawing")
                expect(drawing).not.toBeNull()
                const relId = /relId="([^"]+)"/.exec(await text(zip, data))?.[1]
                const slide = (await readDeckState(await DeckArchive.open(deck))).slides[i] as { part: string }
                expect(await text(zip, relsOf(slide.part))).toContain(`Id="${relId as string}" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"`)
                const picture = /Target="\.\.\/media\/([^"]+)"/.exec(await text(zip, relsOf(data)))?.[1]
                expect(await (zip.file(`ppt/media/${picture as string}`) as JSZip.JSZipObject).async("string")).toBe(`smart-picture-${i + 1}`)
            }
            await expectIntact(deck)
        }
    })

    it("accepts a master written with PowerPoint's and the Open XML SDK's spellings", async () => {
        const deck = path.join(TMP, "foreign-master.pptx")
        await craftDeck(deck, async (zip) => {
            const master = "ppt/slideMasters/slideMaster1.xml"
            zip.file(master, (await text(zip, master))
                .replace("<p:cNvPr id=\"2\" name=\"Title Placeholder 1\"/>",
                    "<p:cNvPr id='2' name='Title Placeholder 1' descr=\"Bild&#xA;&#xA;Automatisch generierte Beschreibung &apos;x&apos;\" />")
                .replace("<a:lstStyle/>", "<a:lstStyle></a:lstStyle>"))
        })
        for (let round = 0; round < 2; round++)
            await apply(deck, [ { op: "slide.fill", slide: "title:A", notes: `r${round}` } ])
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "C" } } } ])
        await expectIntact(deck)
    })

    it("keeps slide-level and shape-level tags (think-cell) and picture bullets", async () => {
        const deck = path.join(TMP, "foreign-tags.pptx")
        await craftDeck(deck, async (zip, slides) => {
            const b = slides[1] as string
            for (const n of [ 1, 2 ]) {
                zip.file(`ppt/tags/tag${n}.xml`, "<p:tagLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
                    + `<p:tag name="THINKCELL" val="t${n}"/></p:tagLst>`)
                await addOverride(zip, `ppt/tags/tag${n}.xml`, "application/vnd.openxmlformats-officedocument.presentationml.tags+xml")
            }
            await addRel(zip, b, "rId93", "tags", "../tags/tag1.xml")
            await addRel(zip, b, "rId94", "tags", "../tags/tag2.xml")
            zip.file("ppt/media/bullet.png", Buffer.from("bullet-picture"))
            zip.file("[Content_Types].xml", (await text(zip, "[Content_Types].xml")).replace("<Default ",
                "<Default Extension=\"png\" ContentType=\"image/png\"/><Default "))
            await addRel(zip, b, "rId95", "image", "../media/bullet.png")
            zip.file(b, (await text(zip, b))
                .replace("<p:ph type=\"title\"/>", "<p:ph type=\"title\"/><p:custDataLst><p:tags r:id=\"rId94\"/></p:custDataLst>")
                .replace("</p:spTree>", "</p:spTree><p:custDataLst><p:tags r:id=\"rId93\"/></p:custDataLst>")
                .replace("<a:p><a:r><a:t>B</a:t>", "<a:p><a:pPr><a:buBlip><a:blip r:embed=\"rId95\"/></a:buBlip></a:pPr><a:r><a:t>B</a:t>"))
        })
        await apply(deck, [ { op: "meta.props", set: { title: "x" } } ])
        await apply(deck, [ { op: "slide.fill", slide: "title:A", notes: "n" } ])
        await apply(deck, [ { op: "slide.copy", slide: "title:B" } ])
        await expectIntact(deck)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        for (const i of [ 1, 2 ]) {
            const slide = (await readDeckState(await DeckArchive.open(deck))).slides[i] as { part: string }
            const xml = await text(zip, slide.part)
            const rels = await text(zip, relsOf(slide.part))
            const tagOf = async (rid: string | undefined): Promise<string> => {
                const target = new RegExp(`Id="${rid ?? ""}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ?? ""
                return await text(zip, path.posix.normalize(path.posix.join("ppt/slides", target)))
            }
            const shapeRid = /<p:ph type="title"\/><p:custDataLst><p:tags r:id="([^"]+)"/.exec(xml)?.[1]
            const slideRid = /<\/p:spTree><p:custDataLst><p:tags r:id="([^"]+)"/.exec(xml)?.[1]
            expect(await tagOf(shapeRid)).toContain("t2")
            expect(await tagOf(slideRid)).toContain("t1")
            const bulletRid = /<a:buBlip><a:blip r:embed="([^"]+)"/.exec(xml)?.[1]
            expect(await tagOf(bulletRid)).toBe("bullet-picture")
        }
    })
})
