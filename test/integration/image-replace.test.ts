/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: el.set `image` replaces the image of a picture or a
**  filled picture placeholder in place -- frame and animations stay, the
**  image is cut centered to fit the frame (a crop is re-cut), the other
**  images of the slide stay, the old media part goes when nothing else
**  uses it. The content guard accepts exactly the declared swap: without
**  it the same output fails.
*/

import { beforeAll, describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import { executeOps } from "../../src/commands/apply.js"
import { digestBytes, verifyContent, type SlideExpect } from "../../src/engine/guard.js"
import { firstElement } from "../../src/engine/xml.js"
import { TMP, apply, craftSlide, failure, freshDeck, png, shapeNamed, slideDoc, stateOf } from "../util/elements.js"

const WIDE = path.join(TMP, "image-replace-wide.png")
const TALL = path.join(TMP, "image-replace-tall.png")
const OTHER = path.join(TMP, "image-replace-other.png")
const VECTOR = path.join(TMP, "image-replace.svg")
let deck: string

/**  an entrance animation of the shape with the given id  */
const timing = (spid: string): string => "<p:timing><p:tnLst><p:par><p:cTn id=\"1\" dur=\"indefinite\" restart=\"never\" nodeType=\"tmRoot\">"
    + "<p:childTnLst><p:seq concurrent=\"1\" nextAc=\"seek\"><p:cTn id=\"2\" dur=\"indefinite\" nodeType=\"mainSeq\"><p:childTnLst>"
    + "<p:par><p:cTn id=\"3\" fill=\"hold\"><p:stCondLst><p:cond delay=\"indefinite\"/></p:stCondLst><p:childTnLst>"
    + "<p:par><p:cTn id=\"4\" presetID=\"1\" presetClass=\"entr\" presetSubtype=\"0\" fill=\"hold\" nodeType=\"clickEffect\">"
    + "<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id=\"5\" dur=\"1\" fill=\"hold\">"
    + "<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>"
    + `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>`
    + "</p:cBhvr><p:to><p:strVal val=\"visible\"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"
    + "</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"

/**  the srcRect attributes and the id of a picture  */
const pictureFacts = async (index: number, name: string): Promise<{ crop: string, id: string | null, box: string }> => {
    const pic = shapeNamed(await slideDoc(deck, index), name)
    const src = firstElement(pic, "a:srcRect")
    const off = firstElement(pic, "a:off")
    const ext = firstElement(pic, "a:ext")
    return {
        crop: src === null ? "none" : [ "l", "t", "r", "b" ].map((a) => `${a}=${src.getAttribute(a) ?? ""}`).join(","),
        id: firstElement(pic, "p:cNvPr")?.getAttribute("id") ?? null,
        box: [ off?.getAttribute("x"), off?.getAttribute("y"), ext?.getAttribute("cx"), ext?.getAttribute("cy") ].join(",")
    }
}

/**  the media file names of a slide's pictures, by shape name prefix  */
const images = async (index: number): Promise<Record<string, string | undefined>> =>
    Object.fromEntries(((await stateOf(deck)).slides[index]?.shapes ?? []).filter((sh) => sh.type === "picture")
        .map((sh) => [ sh.name.replace(/-[0-9a-f-]{36}$/, ""), sh.image ]))

beforeAll(async () => {
    writeFileSync(WIDE, png(1600, 900))
    writeFileSync(TALL, png(900, 1600))
    writeFileSync(OTHER, Buffer.concat([ png(100, 100), Buffer.from("other") ]))
    writeFileSync(VECTOR, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40\" height=\"20\"><rect width=\"40\" height=\"20\"/></svg>")
    deck = await freshDeck("image-replace.pptx")
    await apply(deck, [
        { op: "slide.add", layout: "Title Only", ref: "s", placeholders: { title: { text: "Bilder" } } },
        { op: "el.add", slide: "$s", elements: [
            { type: "image", name: "A", frame: { x: 1, y: 2, w: 3, h: 2 }, path: WIDE, fit: "crop" },
            { type: "image", name: "B", frame: { x: 5, y: 2, w: 2, h: 2 }, path: OTHER } ] },
        { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: WIDE } } }
    ])
    const id = (await pictureFacts(0, "A")).id ?? ""
    await craftSlide(deck, 0, (xml) => xml.replace("</p:sld>", `${timing(id)}</p:sld>`))
})

describe("el.set image", () => {
    it("replaces a picture's image, re-crops it and keeps frame, id, animation and the other image", async () => {
        const before = await pictureFacts(0, "A")
        const was = await images(0)
        await apply(deck, [ { op: "el.set", slide: "0", name: "A", image: TALL } ])
        const after = await pictureFacts(0, "A")
        const now = await images(0)
        expect(after.box).toBe(before.box)
        expect(after.id).toBe(before.id)
        expect(before.crop).toMatch(/^l=\d+,t=,r=\d+,b=$/)
        expect(after.crop).toMatch(/^l=,t=\d+,r=,b=\d+$/)
        expect(now["A"]).not.toBe(was["A"])
        expect(now["B"]).toBe(was["B"])
        expect((await slideDoc(deck, 0)).getElementsByTagName("p:spTgt").item(0)?.getAttribute("spid")).toBe(after.id)
        const zip = await JSZip.loadAsync(readFileSync(deck))
        expect(zip.file(`ppt/media/${was["A"] ?? ""}`)).toBeNull()
    })

    it("replaces the image of a filled picture placeholder, cropped anew", async () => {
        const name = ((await stateOf(deck)).slides[1]?.shapes.find((sh) => sh.type === "picture"))?.name ?? ""
        const before = await pictureFacts(1, name)
        await apply(deck, [ { op: "el.set", slide: "1", name, image: TALL, altText: "hochkant" } ])
        const after = await pictureFacts(1, name)
        expect(after.box).toBe(before.box)
        expect(after.crop).toMatch(/^l=,t=\d+,r=,b=\d+$/)
        expect((await stateOf(deck)).slides[1]?.shapes.find((sh) => sh.name === name)?.altText).toBe("hochkant")
    })

    it("cuts an image of another aspect to fit an uncropped picture's frame, centered; the frame stays", async () => {
        await apply(deck, [ { op: "el.add", slide: "0", elements: [
            { type: "image", name: "D", frame: { x: 1, y: 5, w: 4, h: 2 }, path: WIDE, fit: "stretch" } ] } ])
        const before = await pictureFacts(0, "D")
        expect(before.crop).toBe("none")
        await apply(deck, [ { op: "el.set", slide: "0", name: "D", image: TALL } ])
        const after = await pictureFacts(0, "D")
        expect(after.box).toBe(before.box)
        expect(after.crop).toBe("l=,t=35938,r=,b=35938")

        /*  a frame set with the image: the crop fits the new frame  */
        await apply(deck, [ { op: "el.set", slide: "0", name: "D", frame: { w: 2, h: 2 }, image: WIDE } ])
        expect((await pictureFacts(0, "D")).crop).toBe("l=21875,t=,r=21875,b=")

        /*  an image of unknown size stretches  */
        const unknown = path.join(TMP, "image-replace-unknown.png")
        writeFileSync(unknown, Buffer.from("not really a png"))
        await apply(deck, [ { op: "el.set", slide: "0", name: "D", image: unknown } ])
        expect((await pictureFacts(0, "D")).crop).toBe("l=,t=,r=,b=")
    })

    it("swaps one picture twice in one document", async () => {
        await apply(deck, [
            { op: "el.set", slide: "0", name: "A", image: WIDE },
            { op: "el.set", slide: "0", name: "A", image: TALL, title: "zweimal" }
        ])
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const shown = (await images(0))["A"] ?? ""
        expect(digestBytes(await (zip.file(`ppt/media/${shown}`) as JSZip.JSZipObject).async("nodebuffer"))).toBe(digestBytes(readFileSync(TALL)))
    })

    it("shows an SVG through the SVG extension", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "B", image: VECTOR } ])
        const pic = shapeNamed(await slideDoc(deck, 0), "B")
        const svg = pic.getElementsByTagName("asvg:svgBlip").item(0)
        expect(svg?.getAttribute("r:embed")).toBe(firstElement(pic, "a:blip")?.getAttribute("r:embed"))
        await apply(deck, [ { op: "el.set", slide: "0", name: "B", image: OTHER } ])
        expect(shapeNamed(await slideDoc(deck, 0), "B").getElementsByTagName("asvg:svgBlip").length).toBe(0)
    })

    it("patches images planned in the same document", async () => {
        await apply(deck, [
            { op: "slide.add", layout: "PICTURE", ref: "p", placeholders: { image: { image: WIDE } } },
            { op: "el.set", slide: "$p", name: "PptcPh-1", image: OTHER },
            { op: "el.add", slide: "$p", elements: [ { type: "image", name: "C", frame: { x: 1, y: 1, w: 1, h: 1 }, path: WIDE } ] },
            { op: "el.set", slide: "$p", name: "C", image: TALL }
        ])
        const digests = new Set([ OTHER, TALL ].map((f) => digestBytes(readFileSync(f))))
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const shown = await Promise.all(Object.values(await images(2)).map(async (m) =>
            digestBytes(await (zip.file(`ppt/media/${m ?? ""}`) as JSZip.JSZipObject).async("nodebuffer"))))
        expect(new Set(shown)).toEqual(digests)
    })

    it("rejects an image on a non-picture, an empty picture placeholder and an unsupported file", async () => {
        await apply(deck, [ { op: "slide.add", layout: "PICTURE", placeholders: { title: { text: "leer" } } } ])
        const empty = (await stateOf(deck)).slides[3]?.shapes.find((sh) => sh.placeholderKind === "picture")?.name ?? ""
        expect(await failure(deck, [ { op: "el.set", slide: "3", name: empty, image: WIDE } ])).toBe("E_SCHEMA")
        const title = (await stateOf(deck)).slides[0]?.shapes.find((sh) => sh.placeholderKind === "title")?.name ?? ""
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: title, image: WIDE } ])).toBe("E_SCHEMA")
        const txt = path.join(TMP, "image-replace.txt")
        writeFileSync(txt, "x")
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "A", image: txt } ])).toBe("E_SCHEMA")
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "A", image: path.join(TMP, "missing.png") } ])).toBe("E_FILE")
    })
})

describe("content guard of an image swap", () => {
    it("fails the same output without the declared swap", async () => {
        const before = await JSZip.loadAsync(readFileSync(deck))
        const out = path.join(TMP, "image-replace-out.pptx")
        writeFileSync(out, readFileSync(deck))
        await executeOps(deck, { ops: [ { op: "el.set", slide: "0", name: "A", image: WIDE } ] } as never,
            { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: out, minFontPt: 8 })
        const after = await JSZip.loadAsync(readFileSync(out))
        const part = (await stateOf(deck)).slides[0]?.part ?? ""
        const old = (await images(0))["A"] ?? ""
        const base: SlideExpect = { source: part, label: "#0", notes: false, assets: "same" }
        const all = async (exp: SlideExpect): Promise<SlideExpect[]> => (await stateOf(deck)).slides.map((s, i) => i === 0 ? exp
            : { source: s.part, label: `#${i}`, notes: false, assets: "same" })
        const undeclared = await verifyContent(before, after, await all(base))
        expect(undeclared.blocked.length + undeclared.lost.length).toBeGreaterThan(0)
        const oldDigest = digestBytes(await (before.file(`ppt/media/${old}`) as JSZip.JSZipObject).async("nodebuffer"))
        const declared = await verifyContent(before, after, await all({ ...base,
            replaced: [ { from: [ oldDigest ], to: digestBytes(readFileSync(WIDE)) } ] }))
        expect(declared).toEqual({ blocked: [], lost: [] })
        const wrong = await verifyContent(before, after, await all({ ...base,
            replaced: [ { from: [ oldDigest ], to: digestBytes(readFileSync(OTHER)) } ] }))
        expect(wrong.blocked.length).toBeGreaterThan(0)
    })
})
