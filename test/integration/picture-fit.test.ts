/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: images in picture placeholders. `fit: crop` (the default)
**  covers the frame and cuts the overhang off centered (`a:srcRect`),
**  `contain` shrinks the picture frame around the whole image, `stretch`
**  keeps the frame and distorts; an unknown image size stretches.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState, readTemplateInfo } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { PptcError } from "../../src/infra/errors.js"
import { inchToEmu, type Frame } from "../../src/core/model.js"
import { elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { expectIntact } from "../util/integrity.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")
const DECK = path.join(TMP, "picture-fit.pptx")

const opts: ExecuteOptions = { templatePath: TEMPLATE, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

/**  a PNG header of the given size (the engine reads only the header)  */
const png = (w: number, h: number): Buffer => {
    const b = Buffer.alloc(33)
    Buffer.from([ 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a ]).copy(b, 0)
    b.writeUInt32BE(13, 8)
    b.write("IHDR", 12, "latin1")
    b.writeUInt32BE(w, 16)
    b.writeUInt32BE(h, 20)
    return b
}

const LANDSCAPE = path.join(TMP, "fit-landscape.png")
const PORTRAIT = path.join(TMP, "fit-portrait.png")
const UNKNOWN = path.join(TMP, "fit-unknown.png")
const VECTOR = path.join(TMP, "fit-vector.svg")
let frame: Frame

/**  the picture of the last slide: its srcRect attributes and frame (EMU)  */
const lastPicture = async (): Promise<{ crop: Record<string, string> | null, box: number[] }> => {
    const state = await readDeckState(await DeckArchive.open(DECK))
    const part = (state.slides[state.slides.length - 1] as { part: string }).part
    const zip = await JSZip.loadAsync(readFileSync(DECK))
    const pic = elements(parseXml(await (zip.file(part) as JSZip.JSZipObject).async("string")), "p:pic")[0] as NonNullable<ReturnType<typeof firstElement>>
    const src = firstElement(pic, "a:srcRect")
    const crop = src === null ? null
        : Object.fromEntries([ "l", "t", "r", "b" ].filter((a) => src.hasAttribute(a)).map((a) => [ a, src.getAttribute(a) as string ]))
    const off = firstElement(pic, "a:off")
    const ext = firstElement(pic, "a:ext")
    return { crop, box: [ off?.getAttribute("x"), off?.getAttribute("y"), ext?.getAttribute("cx"), ext?.getAttribute("cy") ].map(Number) }
}

/**  add a picture slide with one image fill  */
const addPicture = async (image: string, fit?: string): Promise<void> => {
    await executeOps(DECK, { ops: [ { op: "slide.add", layout: "PICTURE",
        placeholders: { image: { image, ...(fit !== undefined && { fit }) } } } ] } as never, opts)
    await expectIntact(DECK)
}

beforeAll(async () => {
    mkdirSync(TMP, { recursive: true })
    writeFileSync(LANDSCAPE, png(1600, 900))
    writeFileSync(PORTRAIT, png(900, 1600))
    writeFileSync(UNKNOWN, Buffer.from("not really a png"))
    writeFileSync(VECTOR, "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 400 100\"><rect width=\"400\" height=\"100\"/></svg>")
    writeFileSync(DECK, await buildEmptyDeck(TEMPLATE))
    const layout = (await readTemplateInfo(await DeckArchive.open(TEMPLATE))).layouts.find((l) => l.name === "PICTURE")
    frame = layout?.placeholders.find((p) => p.kind === "picture")?.frame as Frame
})

describe("picture placeholder fit", () => {
    const layoutBox = (): number[] => [ inchToEmu(frame.x), inchToEmu(frame.y), inchToEmu(frame.w), inchToEmu(frame.h) ]
    const aspect = (): number => inchToEmu(frame.w) / inchToEmu(frame.h)

    it("crops a landscape image left and right by default", async () => {
        await addPicture(LANDSCAPE)
        const { crop, box } = await lastPicture()
        const side = Math.round((1 - aspect() / (16 / 9)) / 2 * 100000)
        expect(crop).toEqual({ l: String(side), r: String(side) })
        expect(box).toEqual(layoutBox())
    })

    it("crops a portrait image top and bottom", async () => {
        await addPicture(PORTRAIT, "crop")
        const { crop } = await lastPicture()
        const side = Math.round((1 - (9 / 16) / aspect()) / 2 * 100000)
        expect(crop).toEqual({ t: String(side), b: String(side) })
    })

    it("crops an SVG by its viewBox", async () => {
        await addPicture(VECTOR)
        const { crop } = await lastPicture()
        expect(Object.keys(crop ?? {})).toEqual([ "l", "r" ])
    })

    it("contains the whole image in a centered, shrunk frame", async () => {
        await addPicture(PORTRAIT, "contain")
        const { crop, box } = await lastPicture()
        const [ x, y, w, h ] = layoutBox() as [ number, number, number, number ]
        const cw = Math.round(h * 9 / 16)
        expect(crop).toBeNull()
        expect(box).toEqual([ x + Math.round((w - cw) / 2), y, cw, h ])
        const shapes = (await readDeckState(await DeckArchive.open(DECK))).slides.at(-1)?.shapes ?? []
        expect(shapes.find((sh) => sh.type === "picture")?.frame?.h).toBe(frame.h)
    })

    it("stretches on request and for an unknown image size", async () => {
        await addPicture(LANDSCAPE, "stretch")
        expect(await lastPicture()).toEqual({ crop: null, box: layoutBox() })
        await addPicture(UNKNOWN)
        expect(await lastPicture()).toEqual({ crop: null, box: layoutBox() })
    })

    it("rejects a fit without an image", async () => {
        await expect(executeOps(DECK, { ops: [ { op: "slide.add", layout: "PICTURE",
            placeholders: { 2: { text: "x", fit: "crop" } } } ] } as never, opts))
            .rejects.toSatisfy((err: unknown) => err instanceof PptcError && err.code === "E_SCHEMA")
    })

    it("rejects the fit 'shrink'", async () => {
        await expect(executeOps(DECK, { ops: [ { op: "slide.add", layout: "PICTURE",
            placeholders: { image: { image: LANDSCAPE, fit: "shrink" } } } ] } as never, opts))
            .rejects.toSatisfy((err: unknown) => err instanceof PptcError && err.code === "E_SCHEMA")
        await expect(executeOps(DECK, { ops: [ { op: "slide.fill", slide: "0",
            placeholders: { image: { image: LANDSCAPE, fit: "shrink" } } } ] } as never, opts))
            .rejects.toSatisfy((err: unknown) => err instanceof PptcError && err.code === "E_SCHEMA")
    })
})
