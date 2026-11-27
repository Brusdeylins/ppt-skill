/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: images added by el.add in separate applies must each get
**  their own media file -- a later apply must never overwrite the media
**  of an earlier slide (regression: media12.svg reused across applies).
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { expectIntact } from "../util/integrity.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  smallest valid 1x1 PNG; a trailing marker after IEND makes each copy distinct  */
const PIXEL = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64")
const png = (tag: string): Buffer => Buffer.concat([ PIXEL, Buffer.from(tag) ])

/**  a tiny SVG with a distinct fill color  */
const svg = (color: string): Buffer => Buffer.from(
    `<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="${color}"/></svg>`)

/**  media files referenced by a slide (by presentation order), by name  */
const slideMedia = async (file: string, index: number): Promise<Map<string, Buffer>> => {
    const zip = await JSZip.loadAsync(readFileSync(file))
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const rels = await (zip.file("ppt/_rels/presentation.xml.rels") as JSZip.JSZipObject).async("string")
    const rids = [ ...pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g) ].map((m) => m[1])
    const relMap = new Map([ ...rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g) ].map((m) => [ m[1], m[2] ]))
    const slide = path.posix.basename(relMap.get(rids[index] as string) as string)
    const slideRels = await (zip.file(`ppt/slides/_rels/${slide}.rels`) as JSZip.JSZipObject).async("string")
    const media = new Map<string, Buffer>()
    for (const m of slideRels.matchAll(/Target="\.\.\/media\/([^"]+)"/g)) {
        const name = m[1] as string
        media.set(name, await (zip.file(`ppt/media/${name}`) as JSZip.JSZipObject).async("nodebuffer"))
    }
    return media
}

/**  add one slide with one image per apply, then check every slide still shows its own image  */
const addInSeparateApplies = async (deck: string, images: string[]): Promise<void> => {
    writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
    for (const [ i, image ] of images.entries()) {
        await executeOps(deck, {
            ops: [
                { op: "slide.add", ref: "s", layout: "CONTENT", placeholders: { title: { text: `Bild ${i + 1}` } } },
                { op: "el.add", slide: "$s", elements: [
                    { type: "image", name: `Bild${i + 1}`, frame: { x: 1, y: 1.5, w: 3, h: 3 }, path: image } ] }
            ]
        }, opts)
    }
    await expectIntact(deck)
}

beforeAll(() => {
    mkdirSync(TMP, { recursive: true })
})

describe("media names across applies", () => {
    it("keeps each SVG image on its own slide", async () => {
        const colors = [ "#FF0000", "#00FF00", "#0000FF" ]
        const files = colors.map((c, i) => {
            const f = path.join(TMP, `media-names-${i}.svg`)
            writeFileSync(f, svg(c))
            return f
        })
        const deck = path.join(TMP, "media-names-svg.pptx")
        await addInSeparateApplies(deck, files)
        for (const [ i, color ] of colors.entries()) {
            const svgs = [ ...(await slideMedia(deck, i)).entries() ].filter(([ n ]) => n.endsWith(".svg"))
            expect(svgs).toHaveLength(1)
            expect((svgs[0] as [ string, Buffer ])[1].toString()).toContain(color)
        }
    })

    it("keeps each PNG image on its own slide", async () => {
        const tags = [ "A", "B", "C" ]
        const files = tags.map((t) => {
            const f = path.join(TMP, `media-names-${t}.png`)
            writeFileSync(f, png(t))
            return f
        })
        const deck = path.join(TMP, "media-names-png.pptx")
        await addInSeparateApplies(deck, files)
        for (const [ i, tag ] of tags.entries()) {
            const pngs = [ ...(await slideMedia(deck, i)).values() ]
            expect(pngs).toHaveLength(1)
            expect((pngs[0] as Buffer).equals(png(tag))).toBe(true)
        }
    })

    it("keeps each picture placeholder image on its own slide", async () => {
        const tags = [ "P", "Q", "R" ]
        const deck = path.join(TMP, "media-names-ph.pptx")
        writeFileSync(deck, await buildEmptyDeck(TEMPLATE))
        for (const tag of tags) {
            const f = path.join(TMP, `media-names-${tag}.png`)
            writeFileSync(f, png(tag))
            await executeOps(deck, {
                ops: [ { op: "slide.add", layout: "PICTURE",
                    placeholders: { title: { text: `Bild ${tag}` }, image: { image: f } } } ]
            }, opts)
        }
        await expectIntact(deck)
        for (const [ i, tag ] of tags.entries()) {
            const pngs = [ ...(await slideMedia(deck, i)).values() ]
            expect(pngs).toHaveLength(1)
            expect((pngs[0] as Buffer).equals(png(tag))).toBe(true)
        }
    })
})
