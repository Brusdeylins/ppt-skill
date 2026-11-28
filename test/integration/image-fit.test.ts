/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: el.add images take the fit of picture placeholders --
**  `stretch` (the default) keeps the frame, `crop` covers it with a
**  centered `a:srcRect`, `contain` shrinks the frame centered around the
**  whole image -- and a `hyperlink` (external relationship on the picture).
*/

import { beforeAll, describe, expect, it } from "vitest"
import { writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import { inchToEmu } from "../../src/core/model.js"
import { relTargets, relsPartOf } from "../../src/engine/parts.js"
import { NS_R, firstElement, parseXml } from "../../src/engine/xml.js"
import { TMP, apply, failure, freshDeck, png, shapeOf } from "../util/elements.js"

let deck: string
const LANDSCAPE = path.join(TMP, "el-fit-landscape.png")
const PORTRAIT = path.join(TMP, "el-fit-portrait.png")
const frame = { x: 1, y: 1, w: 4, h: 3 }

beforeAll(async () => {
    deck = await freshDeck("image-fit.pptx")
    writeFileSync(LANDSCAPE, png(1600, 900))
    writeFileSync(PORTRAIT, png(900, 1600))
    await apply(deck, [ { op: "slide.add", layout: "DEFAULT" } ])
})

/**  add an image and read its crop and frame (EMU)  */
const picture = async (name: string, extra: Record<string, unknown>): Promise<{ crop: Record<string, string> | null, box: number[] }> => {
    await apply(deck, [ { op: "el.add", slide: "0", elements: [ { type: "image", name, frame, ...extra } ] } ])
    const { shape } = await shapeOf(deck, name)
    const src = firstElement(shape, "a:srcRect")
    const off = firstElement(shape, "a:off")
    const ext = firstElement(shape, "a:ext")
    return {
        crop: src === null ? null : Object.fromEntries([ "l", "t", "r", "b" ].filter((a) => src.hasAttribute(a)).map((a) => [ a, src.getAttribute(a) as string ])),
        box: [ off?.getAttribute("x"), off?.getAttribute("y"), ext?.getAttribute("cx"), ext?.getAttribute("cy") ].map(Number)
    }
}

const box = [ inchToEmu(frame.x), inchToEmu(frame.y), inchToEmu(frame.w), inchToEmu(frame.h) ]

describe("el.add image fit", () => {
    it("stretches by default", async () => {
        expect(await picture("Stretched", { path: LANDSCAPE })).toEqual({ crop: null, box })
    })

    it("crops centered to the frame", async () => {
        const side = Math.round((1 - (4 / 3) / (16 / 9)) / 2 * 100000)
        expect(await picture("Cropped", { path: LANDSCAPE, fit: "crop" })).toEqual({ crop: { l: String(side), r: String(side) }, box })
    })

    it("contains the whole image in a centered, shrunk frame", async () => {
        const [ x, y, w, h ] = box as [ number, number, number, number ]
        const cw = Math.round(h * 9 / 16)
        expect(await picture("Contained", { path: PORTRAIT, fit: "contain" }))
            .toEqual({ crop: null, box: [ x + Math.round((w - cw) / 2), y, cw, h ] })
    })

    it("rejects the fit 'shrink'", async () => {
        expect(await failure(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "image", name: "Shrunk", frame, path: LANDSCAPE, fit: "shrink" } ] } ])).toBe("E_SCHEMA")
    })

    it("links the picture to a URL", async () => {
        await picture("Linked", { path: LANDSCAPE, hyperlink: "https://example.com/info" })
        const { shape, part, zip } = await shapeOf(deck, "Linked")
        const rid = firstElement(shape, "a:hlinkClick")?.getAttributeNS(NS_R, "id")
        const rels = parseXml(await (zip.file(relsPartOf(part)) as JSZip.JSZipObject).async("string"))
        const rel = relTargets(part, rels).find((r) => r.id === rid)
        expect(rel?.target).toBeNull()
        expect(rel?.el.getAttribute("Target")).toBe("https://example.com/info")
    })
})
