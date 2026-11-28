/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/image-size: the pixel size of an image from its file header --
**  PNG, JPEG, GIF, WebP, BMP and SVG (width/height or viewBox). A small
**  own reader instead of a dependency: image fits only need the aspect
**  ratio, and an unknown size simply falls back to stretching.
*/

/**  pixel size of an image  */
export interface ImageSize {
    /**  width in pixels (SVG: user units)  */
    w: number
    /**  height in pixels (SVG: user units)  */
    h: number
}

/**  a size only when both sides are positive and finite  */
const sized = (w: number, h: number): ImageSize | null =>
    Number.isFinite(w) && Number.isFinite(h) && w > 0 && h > 0 ? { w, h } : null

/**  PNG: the IHDR chunk follows the 8-byte signature  */
const pngSize = (b: Buffer): ImageSize | null =>
    b.length >= 24 && b.toString("latin1", 12, 16) === "IHDR"
        ? sized(b.readUInt32BE(16), b.readUInt32BE(20)) : null

/**  GIF: logical screen size after the 6-byte signature  */
const gifSize = (b: Buffer): ImageSize | null =>
    b.length >= 10 ? sized(b.readUInt16LE(6), b.readUInt16LE(8)) : null

/**  BMP: BITMAPINFOHEADER width and (possibly negative, top-down) height,
     or the 16-bit sizes of an old BITMAPCOREHEADER  */
const bmpSize = (b: Buffer): ImageSize | null => {
    if (b.length < 26)
        return null
    if (b.readUInt32LE(14) === 12)
        return sized(b.readUInt16LE(18), b.readUInt16LE(20))
    return sized(Math.abs(b.readInt32LE(18)), Math.abs(b.readInt32LE(22)))
}

/**  WebP: lossy (VP8), lossless (VP8L) or extended (VP8X) bitstream  */
const webpSize = (b: Buffer): ImageSize | null => {
    if (b.length < 30)
        return null
    const chunk = b.toString("latin1", 12, 16)
    if (chunk === "VP8 ")
        return sized(b.readUInt16LE(26) & 0x3fff, b.readUInt16LE(28) & 0x3fff)
    if (chunk === "VP8L") {
        const bits = b.readUInt32LE(21)
        return sized((bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1)
    }
    if (chunk === "VP8X")
        return sized(b.readUIntLE(24, 3) + 1, b.readUIntLE(27, 3) + 1)
    return null
}

/**  the EXIF orientation of a JPEG APP1 segment, 1 when absent  */
const exifOrientation = (b: Buffer, start: number, end: number): number => {
    if (b.toString("latin1", start, start + 6) !== "Exif\0\0")
        return 1
    const tiff = start + 6
    if (tiff + 8 > end)
        return 1
    const le = b.toString("latin1", tiff, tiff + 2) === "II"
    const u16 = (o: number): number => le ? b.readUInt16LE(o) : b.readUInt16BE(o)
    const u32 = (o: number): number => le ? b.readUInt32LE(o) : b.readUInt32BE(o)
    const ifd = tiff + u32(tiff + 4)
    if (ifd + 2 > end)
        return 1
    const count = u16(ifd)
    for (let i = 0; i < count; i++) {
        const entry = ifd + 2 + i * 12
        if (entry + 12 > end)
            break
        if (u16(entry) === 0x0112)
            return u16(entry + 8)
    }
    return 1
}

/**  JPEG: the first start-of-frame marker; an EXIF orientation of 5 to 8
     (rotated by 90 degrees) swaps the displayed width and height  */
const jpegSize = (b: Buffer): ImageSize | null => {
    let orientation = 1
    let i = 2
    while (i + 9 < b.length) {
        if (b[i] !== 0xff) {
            i++
            continue
        }
        const marker = b[i + 1] as number
        if (marker === 0xff) {
            i++
            continue
        }
        if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
            i += 2
            continue
        }
        const length = b.readUInt16BE(i + 2)
        if (marker === 0xe1)
            orientation = exifOrientation(b, i + 4, Math.min(b.length, i + 2 + length))
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
            const h = b.readUInt16BE(i + 5)
            const w = b.readUInt16BE(i + 7)
            return orientation >= 5 && orientation <= 8 ? sized(h, w) : sized(w, h)
        }
        if (marker === 0xd9 || marker === 0xda)
            return null
        i += 2 + length
    }
    return null
}

/**  CSS length units in user units (px) of an SVG width/height  */
const SVG_UNITS: Record<string, number> = { "": 1, px: 1, pt: 4 / 3, pc: 16, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 }

/**  an absolute SVG length in user units, null for percentages and the rest  */
const svgLength = (value: string | undefined): number | null => {
    const m = /^\s*([0-9]*\.?[0-9]+(?:e[+-]?\d+)?)\s*([a-z]*)\s*$/i.exec(value ?? "")
    const unit = SVG_UNITS[(m?.[2] ?? "").toLowerCase()]
    return m === null || unit === undefined ? null : Number(m[1]) * unit
}

/**  SVG: width/height of the root element, else its viewBox  */
const svgSize = (b: Buffer): ImageSize | null => {
    const root = /<svg\b([^>]*)>/i.exec(b.toString("utf8", 0, Math.min(b.length, 65536)))
    if (root === null)
        return null
    const attr = (name: string): string | undefined =>
        new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(root[1] as string)?.slice(1).find((v) => v !== undefined)
    const w = svgLength(attr("width"))
    const h = svgLength(attr("height"))
    if (w !== null && h !== null)
        return sized(w, h)
    const box = (attr("viewBox") ?? "").trim().split(/[\s,]+/).map(Number)
    return box.length === 4 ? sized(box[2] as number, box[3] as number) : null
}

/**
 *  Read the pixel size of an image from its header bytes.
 *
 *  @param bytes - the image file content
 *  @returns width and height, null when the format or size is unknown
 */
export const imageSize = (bytes: Buffer): ImageSize | null => {
    try {
        if (bytes.length >= 8 && bytes.readUInt32BE(0) === 0x89504e47)
            return pngSize(bytes)
        if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8)
            return jpegSize(bytes)
        if (/^GIF8[79]a/.test(bytes.toString("latin1", 0, 6)))
            return gifSize(bytes)
        if (bytes.toString("latin1", 0, 4) === "RIFF" && bytes.toString("latin1", 8, 12) === "WEBP")
            return webpSize(bytes)
        if (bytes.toString("latin1", 0, 2) === "BM")
            return bmpSize(bytes)
        return svgSize(bytes)
    }
    catch {
        /*  a truncated header reads past the end: size unknown  */
        return null
    }
}

/**  how an image takes a frame of another aspect ratio: `crop` covers the
     frame and cuts the overhang off centered (PowerPoint's picture
     placeholder behaviour), `contain` shows the whole image centered in a
     shrunk frame, `stretch` distorts it to the frame  */
export type ImageFit = "crop" | "contain" | "stretch"

/**  a box on the slide, in any one unit (EMU in the engine)  */
export interface Box {
    x: number
    y: number
    w: number
    h: number
}

/**  the outcome of fitting an image into a box  */
export interface Fitted {
    /**  the picture's frame (the box itself, or shrunk for `contain`)  */
    box: Box
    /**  the `a:srcRect` insets in 1/1000 percent (100000 = 100%), null
         when nothing is cut off  */
    crop: { l: number, t: number, r: number, b: number } | null
}

/**
 *  Fit an image into a box.
 *
 *  @param box - the frame the image is to fill
 *  @param size - the image's pixel size, null when unknown (stretches)
 *  @param fit - the fit mode
 *  @returns the picture frame and its crop
 */
export const fitImage = (box: Box, size: ImageSize | null, fit: ImageFit): Fitted => {
    if (size === null || fit === "stretch" || box.w <= 0 || box.h <= 0)
        return { box, crop: null }
    const image = size.w / size.h
    const frame = box.w / box.h
    if (fit === "contain") {
        const w = image >= frame ? box.w : Math.round(box.h * image)
        const h = image >= frame ? Math.round(box.w / image) : box.h
        return { box: { x: box.x + Math.round((box.w - w) / 2), y: box.y + Math.round((box.h - h) / 2), w, h }, crop: null }
    }
    const side = Math.round((1 - (image > frame ? frame / image : image / frame)) / 2 * 100000)
    if (side <= 0)
        return { box, crop: null }
    return { box, crop: image > frame ? { l: side, t: 0, r: side, b: 0 } : { l: 0, t: side, r: 0, b: side } }
}
