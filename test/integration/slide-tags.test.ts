/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: slide tags (<p:custDataLst> -> <p:tags r:id>, the carrier of
**  think-cell chart links and other add-in metadata) must survive an apply.
**  Regression guard for the defect where pptx-automizer's "unsupported tag"
**  cleaning silently stripped the element and its /tags relationship from
**  every re-imported slide -- unlinking think-cell charts deck-wide.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { executeOps } from "../../src/commands/apply.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "slide-tags.pptx")

const opts = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

const TAGS_CT = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
const TAGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"

/**  attach a tags part to the first slide -- emulating a deck whose slide
     carries think-cell/add-in metadata in p:custDataLst  */
const injectSlideTags = async (file: string): Promise<void> => {
    const zip = await JSZip.loadAsync(readFileSync(file))
    zip.file("ppt/tags/tag1.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + "<p:tagLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
        + "<p:tag name=\"THINKCELL\" val=\"marker\"/></p:tagLst>")
    const ct = await (zip.file("[Content_Types].xml") as JSZip.JSZipObject).async("string")
    zip.file("[Content_Types].xml", ct.replace("</Types>",
        `<Override PartName="/ppt/tags/tag1.xml" ContentType="${TAGS_CT}"/></Types>`))
    const part = Object.keys(zip.files)
        .find((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)) as string
    const relsPart = `ppt/slides/_rels/${path.basename(part)}.rels`
    const rels = await (zip.file(relsPart) as JSZip.JSZipObject).async("string")
    zip.file(relsPart, rels.replace("</Relationships>",
        `<Relationship Id="rId77" Type="${TAGS_REL}" Target="../tags/tag1.xml"/></Relationships>`))
    const slide = await (zip.file(part) as JSZip.JSZipObject).async("string")
    zip.file(part, slide.replace("</p:spTree>",
        "</p:spTree><p:custDataLst><p:tags r:id=\"rId77\"/></p:custDataLst>"))
    writeFileSync(file, await zip.generateAsync({ type: "nodebuffer" }))
}

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    await executeOps(DECK, {
        ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Tagged" } } }]
    }, opts)
    await injectSlideTags(DECK)
})

describe("slide tags survive an apply", () => {
    it("keeps <p:custDataLst> wired to the tags part after editing", async () => {
        await executeOps(DECK, {
            ops: [{ op: "slide.fill", slide: "title:Tagged", placeholders: { body: { text: "edited" } } }]
        }, opts)

        /*  automizer renumbers slide parts on re-import: locate the slide
            part carrying the restored list instead of assuming slide1.xml  */
        const zip = await JSZip.loadAsync(readFileSync(DECK))
        const parts = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
        expect(parts).toHaveLength(1)
        const part = parts[0] as string
        const slide = await (zip.file(part) as JSZip.JSZipObject).async("string")
        const m = /<p:custDataLst><p:tags r:id="(rIdPptc\d+)"\/><\/p:custDataLst>/.exec(slide)
        expect(m).not.toBeNull()
        const rels = await (zip.file(`ppt/slides/_rels/${path.basename(part)}.rels`) as JSZip.JSZipObject).async("string")
        expect(rels).toContain(`Id="${m?.[1]}" Type="${TAGS_REL}" Target="../tags/tag1.xml"`)
        expect(zip.file("ppt/tags/tag1.xml")).not.toBeNull()
        await expectIntact(DECK)
    })

    it("clones the tags part when a tagged slide is copied", async () => {
        /*  PowerPoint mints an independent tags part per duplicated slide
            (think-cell tags are per-instance); a shared part would silently
            alias both slides onto one think-cell chart instance  */
        await executeOps(DECK, {
            ops: [{ op: "slide.copy", slide: "title:Tagged" }]
        }, opts)

        const zip = await JSZip.loadAsync(readFileSync(DECK))
        const parts = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
        expect(parts).toHaveLength(2)
        const targets: string[] = []
        for (const part of parts) {
            const slide = await (zip.file(part) as JSZip.JSZipObject).async("string")
            const m = /<p:custDataLst><p:tags r:id="(rIdPptc\d+)"\/><\/p:custDataLst>/.exec(slide)
            expect(m).not.toBeNull()
            const rels = await (zip.file(`ppt/slides/_rels/${path.basename(part)}.rels`) as JSZip.JSZipObject).async("string")
            const rel = new RegExp(`Id="${m?.[1]}" Type="[^"]*" Target="([^"]+)"`).exec(rels)
            expect(rel).not.toBeNull()
            targets.push(rel?.[1] as string)
            expect(zip.file(path.posix.normalize(path.posix.join("ppt/slides", rel?.[1] as string)))).not.toBeNull()
        }
        expect(new Set(targets).size).toBe(2)
        await expectIntact(DECK)
    })
})

describe("presentation-level tags survive an apply", () => {
    it("keeps the presentation's p:custDataLst wired across a rebuild", async () => {
        /*  think-cell parks workbook-level data in the PRESENTATION's
            custDataLst; automizer happens not to clean presentation.xml
            today -- this guard turns a future regression into a test failure
            instead of silent data loss  */
        const zipIn = await JSZip.loadAsync(readFileSync(DECK))
        zipIn.file("ppt/tags/tagPres.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<p:tagLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
            + "<p:tag name=\"THINKCELL_WORKBOOK\" val=\"marker\"/></p:tagLst>")
        const ct = await (zipIn.file("[Content_Types].xml") as JSZip.JSZipObject).async("string")
        zipIn.file("[Content_Types].xml", ct.replace("</Types>",
            `<Override PartName="/ppt/tags/tagPres.xml" ContentType="${TAGS_CT}"/></Types>`))
        const presRels = await (zipIn.file("ppt/_rels/presentation.xml.rels") as JSZip.JSZipObject).async("string")
        zipIn.file("ppt/_rels/presentation.xml.rels", presRels.replace("</Relationships>",
            `<Relationship Id="rId88" Type="${TAGS_REL}" Target="tags/tagPres.xml"/></Relationships>`))
        const pres = await (zipIn.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
        zipIn.file("ppt/presentation.xml", pres.replace(/(<p:notesSz[^>]*\/>)/,
            "$1<p:custDataLst><p:tags r:id=\"rId88\"/></p:custDataLst>"))
        writeFileSync(DECK, await zipIn.generateAsync({ type: "nodebuffer" }))

        await executeOps(DECK, {
            ops: [{ op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Third" } } }]
        }, opts)

        const zip = await JSZip.loadAsync(readFileSync(DECK))
        const presOut = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
        expect(presOut).toContain("<p:custDataLst><p:tags r:id=\"rId88\"/></p:custDataLst>")
        const relsOut = await (zip.file("ppt/_rels/presentation.xml.rels") as JSZip.JSZipObject).async("string")
        expect(relsOut).toContain(`Id="rId88" Type="${TAGS_REL}" Target="tags/tagPres.xml"`)
        expect(zip.file("ppt/tags/tagPres.xml")).not.toBeNull()
        await expectIntact(DECK)
    })
})
