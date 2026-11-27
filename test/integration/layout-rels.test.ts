/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: a new slide clones the placeholders of its layout; a
**  placeholder referencing a layout relationship (a picture fill, as
**  corporate templates use for decorative circles) must reference a
**  relationship of the slide itself -- otherwise PowerPoint shows "the
**  picture cannot be displayed".
*/

import { describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

const opts: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  smallest valid 1x1 PNG  */
const PIXEL = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64")

describe("placeholders cloned from a layout", () => {
    it("keep their picture fill on a new slide", async () => {
        mkdirSync(TMP, { recursive: true })
        const deck = path.join(TMP, "layout-rels.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))

        /*  give the title placeholder of layout CONTENT a picture fill  */
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const text = async (part: string): Promise<string> => await (zip.file(part) as JSZip.JSZipObject).async("string")
        let layout = ""
        for (const f of Object.keys(zip.files).filter((n) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(n)))
            if ((await text(f)).includes("name=\"CONTENT\""))
                layout = f
        const relsPart = `ppt/slideLayouts/_rels/${path.posix.basename(layout)}.rels`
        zip.file("ppt/media/layoutFill.png", PIXEL)
        const ct = await text("[Content_Types].xml")
        if (!ct.includes("Extension=\"png\""))
            zip.file("[Content_Types].xml", ct.replace("<Default ", "<Default Extension=\"png\" ContentType=\"image/png\"/><Default "))
        zip.file(relsPart, (await text(relsPart)).replace("</Relationships>",
            "<Relationship Id=\"rIdFill\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\""
            + " Target=\"../media/layoutFill.png\"/></Relationships>"))
        const xml = await text(layout)
        const title = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?type="title"(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/.exec(xml)?.[0] as string
        const filled = title.replace(/<p:spPr\/>/, "<p:spPr></p:spPr>").replace("</p:spPr>",
            "<a:blipFill><a:blip r:embed=\"rIdFill\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>")
        expect(filled).toContain("rIdFill")
        zip.file(layout, xml.replace(title, filled))
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))

        await executeOps(deck, { ops: [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Mit Füllung" } } } ] } as never, opts)
        await expectIntact(deck)

        /*  every relationship id the new slide references resolves to an
            existing part (or is external)  */
        const slide = (await readDeckState(await DeckArchive.open(deck))).slides[0] as { part: string }
        const out = await JSZip.loadAsync(readFileSync(deck))
        const slideXml = await (out.file(slide.part) as JSZip.JSZipObject).async("string")
        const rels = await (out.file(`ppt/slides/_rels/${path.posix.basename(slide.part)}.rels`) as JSZip.JSZipObject).async("string")
        const used = [ ...slideXml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g) ].map((m) => m[1] as string)
        expect(used.length).toBeGreaterThan(0)
        for (const id of used) {
            const rel = new RegExp(`<Relationship [^>]*Id="${id}"[^>]*>`).exec(rels)?.[0]
            expect(rel, `relationship ${id}`).toBeDefined()
            const target = /Target="([^"]+)"/.exec(rel as string)?.[1] as string
            expect(out.file(path.posix.normalize(path.posix.join("ppt/slides", target))), target).not.toBeNull()
        }
    })
})
