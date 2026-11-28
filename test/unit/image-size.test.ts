/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the image header reader (PNG, JPEG incl. EXIF rotation, GIF,
**  WebP, BMP, SVG) and the fit of an image into a frame (crop, contain,
**  stretch).
*/

import { describe, expect, it } from "vitest"
import { fitImage, imageSize } from "../../src/engine/image-size.js"

/**  a PNG header with IHDR  */
const png = (w: number, h: number): Buffer => {
    const b = Buffer.alloc(33)
    Buffer.from([ 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a ]).copy(b, 0)
    b.writeUInt32BE(13, 8)
    b.write("IHDR", 12, "latin1")
    b.writeUInt32BE(w, 16)
    b.writeUInt32BE(h, 20)
    return b
}

/**  a JPEG with an optional EXIF orientation, a DQT segment and SOF0  */
const jpeg = (w: number, h: number, orientation?: number): Buffer => {
    const parts: Buffer[] = [ Buffer.from([ 0xff, 0xd8 ]) ]
    if (orientation !== undefined) {
        const tiff = Buffer.alloc(26)
        tiff.write("MM", 0, "latin1")
        tiff.writeUInt16BE(42, 2)
        tiff.writeUInt32BE(8, 4)
        tiff.writeUInt16BE(1, 8)
        tiff.writeUInt16BE(0x0112, 10)
        tiff.writeUInt16BE(3, 12)
        tiff.writeUInt32BE(1, 14)
        tiff.writeUInt16BE(orientation, 18)
        const body = Buffer.concat([ Buffer.from("Exif\0\0", "latin1"), tiff ])
        const head = Buffer.from([ 0xff, 0xe1, 0, 0 ])
        head.writeUInt16BE(body.length + 2, 2)
        parts.push(head, body)
    }
    parts.push(Buffer.from([ 0xff, 0xdb, 0x00, 0x04, 0x00, 0x00 ]))
    const sof = Buffer.from([ 0xff, 0xc0, 0x00, 0x11, 0x08, 0, 0, 0, 0, 0x03 ])
    sof.writeUInt16BE(h, 5)
    sof.writeUInt16BE(w, 7)
    parts.push(sof, Buffer.alloc(12))
    return Buffer.concat(parts)
}

/**  a GIF header  */
const gif = (w: number, h: number): Buffer => {
    const b = Buffer.alloc(13)
    b.write("GIF89a", 0, "latin1")
    b.writeUInt16LE(w, 6)
    b.writeUInt16LE(h, 8)
    return b
}

/**  a BMP header with BITMAPINFOHEADER (negative height: top-down)  */
const bmp = (w: number, h: number): Buffer => {
    const b = Buffer.alloc(54)
    b.write("BM", 0, "latin1")
    b.writeUInt32LE(40, 14)
    b.writeInt32LE(w, 18)
    b.writeInt32LE(h, 22)
    return b
}

/**  a WebP header of the given bitstream chunk  */
const webp = (chunk: "VP8 " | "VP8L" | "VP8X", w: number, h: number): Buffer => {
    const b = Buffer.alloc(40)
    b.write("RIFF", 0, "latin1")
    b.write("WEBP", 8, "latin1")
    b.write(chunk, 12, "latin1")
    if (chunk === "VP8 ") {
        b.writeUInt16LE(w, 26)
        b.writeUInt16LE(h, 28)
    }
    else if (chunk === "VP8L") {
        b[20] = 0x2f
        b.writeUInt32LE(((w - 1) & 0x3fff) | (((h - 1) & 0x3fff) << 14), 21)
    }
    else {
        b.writeUIntLE(w - 1, 24, 3)
        b.writeUIntLE(h - 1, 27, 3)
    }
    return b
}

const svg = (attrs: string): Buffer =>
    Buffer.from(`<?xml version="1.0"?>\n<svg xmlns="http://www.w3.org/2000/svg" ${attrs}><rect/></svg>`)

describe("image size", () => {
    it("reads PNG, GIF, BMP and JPEG headers", () => {
        expect(imageSize(png(1600, 900))).toEqual({ w: 1600, h: 900 })
        expect(imageSize(gif(320, 200))).toEqual({ w: 320, h: 200 })
        expect(imageSize(bmp(640, -480))).toEqual({ w: 640, h: 480 })
        expect(imageSize(jpeg(4000, 3000))).toEqual({ w: 4000, h: 3000 })
    })

    it("swaps a JPEG's sides for an EXIF rotation by 90 degrees", () => {
        expect(imageSize(jpeg(4000, 3000, 6))).toEqual({ w: 3000, h: 4000 })
        expect(imageSize(jpeg(4000, 3000, 3))).toEqual({ w: 4000, h: 3000 })
    })

    it("reads the three WebP bitstreams", () => {
        expect(imageSize(webp("VP8 ", 800, 600))).toEqual({ w: 800, h: 600 })
        expect(imageSize(webp("VP8L", 1024, 512))).toEqual({ w: 1024, h: 512 })
        expect(imageSize(webp("VP8X", 3000, 2000))).toEqual({ w: 3000, h: 2000 })
    })

    it("reads SVG width/height with units, else the viewBox", () => {
        expect(imageSize(svg("width=\"200\" height=\"100\""))).toEqual({ w: 200, h: 100 })
        expect(imageSize(svg("width='2in' height='96px'"))).toEqual({ w: 192, h: 96 })
        expect(imageSize(svg("width=\"100%\" height=\"100%\" viewBox=\"0 0 300 150\""))).toEqual({ w: 300, h: 150 })
        expect(imageSize(svg("viewBox=\"0,0,90,160\""))).toEqual({ w: 90, h: 160 })
        expect(imageSize(svg("viewBox=\"0 0 0 10\""))).toBeNull()
    })

    it("returns null for unknown or truncated data", () => {
        expect(imageSize(Buffer.from("hello"))).toBeNull()
        expect(imageSize(png(10, 10).subarray(0, 18))).toBeNull()
        expect(imageSize(jpeg(10, 10).subarray(0, 12))).toBeNull()
        expect(imageSize(Buffer.alloc(0))).toBeNull()
    })
})

describe("image fit", () => {
    const box = { x: 1000, y: 2000, w: 4000, h: 3000 }

    it("crops a wider image left and right, a taller one top and bottom", () => {
        expect(fitImage(box, { w: 1600, h: 900 }, "crop"))
            .toEqual({ box, crop: { l: 12500, t: 0, r: 12500, b: 0 } })
        expect(fitImage(box, { w: 900, h: 1600 }, "crop"))
            .toEqual({ box, crop: { l: 0, t: 28906, r: 0, b: 28906 } })
    })

    it("does not crop an image of the frame's aspect", () => {
        expect(fitImage(box, { w: 800, h: 600 }, "crop")).toEqual({ box, crop: null })
    })

    it("contains the image in a centered, shrunk frame", () => {
        expect(fitImage(box, { w: 1600, h: 900 }, "contain"))
            .toEqual({ box: { x: 1000, y: 2375, w: 4000, h: 2250 }, crop: null })
        expect(fitImage(box, { w: 900, h: 1600 }, "contain"))
            .toEqual({ box: { x: 2156, y: 2000, w: 1688, h: 3000 }, crop: null })
    })

    it("stretches on request or when the size is unknown", () => {
        expect(fitImage(box, { w: 1600, h: 900 }, "stretch")).toEqual({ box, crop: null })
        expect(fitImage(box, null, "crop")).toEqual({ box, crop: null })
        expect(fitImage(box, null, "contain")).toEqual({ box, crop: null })
    })
})
