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
})
