/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/edits: the in-place changes of existing elements on one
**  slide DOM -- el.set (text with the kept format, frame, fill, border,
**  font color, alt text and title, the image of a picture, table cells and
**  rows) and text.replace on the content the slide carried before the
**  apply (shapes, placeholders, table cells, notes; never fields).
*/

import { readFileSync } from "node:fs"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../../infra/errors.js"
import { inchToEmu } from "../../core/model.js"
import type { PlannedEdit, PlannedReplace } from "../../core/ops/registry.js"
import type { RichText } from "../../schema/payloads.js"
import { digestBytes } from "../guard.js"
import { fitImage, imageSize } from "../image-size.js"
import { REL } from "../parts.js"
import { replaceInBody, setBodyText, setClickLink, setShapeText, type ListStyle } from "../text.js"
import { NS_A, NS_R, colorElement, elements, firstElement, relRefs } from "../xml.js"
import type { Package } from "./package.js"

/**  DrawingML fill elements (one of them per fill slot)  */
const FILLS = [ "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill" ]

/**  the namespace of the SVG picture extension  */
const NS_ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"

/**  the URI of the SVG picture extension  */
const SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

/**  the direct children of an element with one of the given names  */
const children = (parent: Element | null, names: string[]): Element[] => {
    const out: Element[] = []
    for (let n = parent?.firstChild ?? null; n !== null; n = n.nextSibling)
        if (n.nodeType === 1 && names.includes(n.nodeName))
            out.push(n as Element)
    return out
}

/**  insert a child before the first direct child named in `followers`
     (schema order), at the end when there is none  */
const insertOrdered = (parent: Element, child: Element, followers: string[]): void => {
    parent.insertBefore(child, children(parent, followers)[0] ?? null)
}

/**  a direct child, created in schema order when missing  */
const ensureChild = (parent: Element, name: string, followers: string[]): Element => {
    const found = children(parent, [ name ])[0]
    if (found !== undefined)
        return found
    const el = (parent.ownerDocument as Document).createElementNS(NS_A, name)
    insertOrdered(parent, el, followers)
    return el
}

/**  replace the fill of a property element by a solid color  */
const setSolidFill = (props: Element, color: string, followers: string[]): void => {
    for (const old of children(props, FILLS))
        props.removeChild(old)
    const fill = (props.ownerDocument as Document).createElementNS(NS_A, "a:solidFill")
    fill.appendChild(colorElement(props.ownerDocument as Document, color))
    insertOrdered(props, fill, followers)
}

/**  DrawingML color elements (one of them per color)  */
const COLORS = [ "a:scrgbClr", "a:srgbClr", "a:hslClr", "a:sysClr", "a:schemeClr", "a:prstClr" ]

/**  set the transparency of a color in percent: its `a:alpha` (none at 0)  */
const setAlpha = (color: Element | null, transparency: number): void => {
    if (color === null)
        return
    for (const old of children(color, [ "a:alpha" ]))
        color.removeChild(old)
    if (transparency <= 0)
        return
    const alpha = (color.ownerDocument as Document).createElementNS(NS_A, "a:alpha")
    alpha.setAttribute("val", String(Math.round((100 - transparency) * 1000)))
    color.appendChild(alpha)
}

/**  the shape properties of an element (`p:spPr`, a group's `p:grpSpPr`)  */
const shapeProps = (shape: Element): Element | null =>
    children(shape, [ "p:spPr", "p:grpSpPr" ])[0] ?? null

/**  the transform of an element: `p:xfrm` of a graphic frame, else the
     `a:xfrm` of its shape properties  */
const transformOf = (shape: Element): Element | null =>
    shape.nodeName === "p:graphicFrame" ? children(shape, [ "p:xfrm" ])[0] ?? null
        : children(shapeProps(shape), [ "a:xfrm" ])[0] ?? null

/**  the box of a transform in EMU, null when incomplete  */
const boxOf = (xfrm: Element | null): { x: number, y: number, w: number, h: number } | null => {
    const off = children(xfrm, [ "a:off" ])[0]
    const ext = children(xfrm, [ "a:ext" ])[0]
    if (off === undefined || ext === undefined)
        return null
    const box = { x: Number(off.getAttribute("x")), y: Number(off.getAttribute("y")),
        w: Number(ext.getAttribute("cx")), h: Number(ext.getAttribute("cy")) }
    return Object.values(box).every(Number.isFinite) ? box : null
}

/**  the rows of a table and the cells of a row (direct children)  */
const rowsOf = (tbl: Element): Element[] => children(tbl, [ "a:tr" ])
const cellsOf = (tr: Element): Element[] => children(tr, [ "a:tc" ])

/**  scale the column widths (`w`) or row heights (`h`) of a table to a
     new total; returns the total written  */
const scaleTable = (tbl: Element, axis: "w" | "h", total: number): number => {
    const items = axis === "w" ? children(children(tbl, [ "a:tblGrid" ])[0] ?? null, [ "a:gridCol" ]) : rowsOf(tbl)
    const attr = axis === "w" ? "w" : "h"
    const sizes = items.map((el) => Number(el.getAttribute(attr) ?? "0") || 0)
    const sum = sizes.reduce((a, b) => a + b, 0)
    if (sum <= 0)
        return total
    let written = 0
    items.forEach((el, i) => {
        const size = Math.round((sizes[i] as number) * total / sum)
        el.setAttribute(attr, String(size))
        written += size
    })
    return written
}

/**  the sum of a table's row heights  */
const tableHeight = (tbl: Element): number =>
    rowsOf(tbl).reduce((sum, tr) => sum + (Number(tr.getAttribute("h") ?? "0") || 0), 0)

/**  set the transform of an element: the given keys of the (complete,
     when the element has no transform yet) frame patch, in inches and
     degrees; a table's columns and rows follow its new size  */
const setFrame = (shape: Element, frame: NonNullable<PlannedEdit["set"]["frame"]>): void => {
    const doc = shape.ownerDocument as Document
    let xfrm = transformOf(shape)
    if (xfrm === null) {
        const props = shapeProps(shape)
        if (props === null)
            throw new PptcError("E_ENGINE", "element without shape properties")
        xfrm = doc.createElementNS(NS_A, "a:xfrm")
        props.insertBefore(xfrm, props.firstChild)
    }
    const off = ensureChild(xfrm, "a:off", [ "a:ext", "a:chOff", "a:chExt" ])
    const ext = ensureChild(xfrm, "a:ext", [ "a:chOff", "a:chExt" ])
    if (frame.x !== undefined)
        off.setAttribute("x", String(inchToEmu(frame.x)))
    if (frame.y !== undefined)
        off.setAttribute("y", String(inchToEmu(frame.y)))
    for (const [ key, attr ] of [ [ "x", "x" ], [ "y", "y" ] ] as const)
        if (!off.hasAttribute(attr))
            off.setAttribute(attr, String(inchToEmu(frame[key] ?? 0)))
    const tbl = firstElement(shape, "a:tbl")
    if (frame.w !== undefined)
        ext.setAttribute("cx", String(tbl === null ? inchToEmu(frame.w) : scaleTable(tbl, "w", inchToEmu(frame.w))))
    if (frame.h !== undefined)
        ext.setAttribute("cy", String(tbl === null ? inchToEmu(frame.h) : scaleTable(tbl, "h", inchToEmu(frame.h))))
    if (frame.rotation !== undefined) {
        const rot = Math.round((((frame.rotation % 360) + 360) % 360) * 60000)
        if (rot === 0)
            xfrm.removeAttribute("rot")
        else
            xfrm.setAttribute("rot", String(rot))
    }
}

/**  the outline of an element, created in schema order when missing  */
const outlineOf = (shape: Element): Element => {
    const props = shapeProps(shape)
    if (props === null)
        throw new PptcError("E_ENGINE", "element without shape properties")
    return ensureChild(props, "a:ln", [ "a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst" ])
}

/**  the run property elements of a shape's text (runs, fields, paragraph
     ends); a run without properties gets them  */
const runProps = (shape: Element): Element[] => {
    const txBody = children(shape, [ "p:txBody" ])[0] ?? null
    const out: Element[] = []
    for (const p of children(txBody, [ "a:p" ])) {
        for (const r of children(p, [ "a:r", "a:fld" ])) {
            let rPr = children(r, [ "a:rPr" ])[0]
            if (rPr === undefined) {
                rPr = (shape.ownerDocument as Document).createElementNS(NS_A, "a:rPr")
                r.insertBefore(rPr, r.firstChild)
            }
            out.push(rPr)
        }
        out.push(...children(p, [ "a:endParaRPr" ]))
    }
    return out
}

/**  set or clear (empty value) an attribute of `p:cNvPr`  */
const setDescription = (shape: Element, attr: "descr" | "title", value: string): void => {
    const cNvPr = firstElement(shape, "p:cNvPr")
    if (cNvPr === null)
        return
    if (value === "")
        cNvPr.removeAttribute(attr)
    else
        cNvPr.setAttribute(attr, value)
}

/**  a cell's text body, created when the cell has none  */
const cellBody = (tc: Element): Element => {
    const found = children(tc, [ "a:txBody" ])[0]
    if (found !== undefined)
        return found
    const doc = tc.ownerDocument as Document
    const txBody = doc.createElementNS(NS_A, "a:txBody")
    txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"))
    txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"))
    txBody.appendChild(doc.createElementNS(NS_A, "a:p"))
    tc.insertBefore(txBody, tc.firstChild)
    return txBody
}

/**  the text of a row cell value  */
const cellText = (value: RichText | number | boolean): RichText =>
    typeof value === "number" || typeof value === "boolean" ? String(value) : value

/**  remove rows, add rows in the format of the neighbouring body row, set cells;
     the frame's height follows the rows  */
const editTable = (shape: Element, set: PlannedEdit["set"], lang: string): void => {
    const tbl = firstElement(shape, "a:tbl")
    if (tbl === null)
        throw new PptcError("E_ENGINE", "table element without a:tbl")
    const before = tableHeight(tbl)
    for (const row of [ ...new Set(set.removeRows ?? []) ].sort((a, b) => b - a)) {
        const tr = rowsOf(tbl)[row]
        if (tr !== undefined)
            tbl.removeChild(tr)

        /*  without its header row the table has none  */
        if (row === 0)
            children(tbl, [ "a:tblPr" ])[0]?.removeAttribute("firstRow")
    }
    if (set.addRows !== undefined) {
        const rows = rowsOf(tbl)
        const at = set.addRows.at ?? rows.length

        /*  the row above is the format model -- but the header row 0 only
            when no body row follows: at 1 the first body row is the model  */
        const model = at === 1 && rows.length > 1 ? rows[1] : rows[at > 0 ? at - 1 : 0]
        if (model === undefined)
            throw new PptcError("E_ENGINE", "table without rows")
        const anchor = rows[at] ?? children(tbl, [ "a:extLst" ])[0] ?? null
        for (const values of set.addRows.rows) {
            const tr = model.cloneNode(true) as Element

            /*  the copy is a row of its own: no merges, no row id  */
            for (const ext of children(tr, [ "a:extLst" ]))
                tr.removeChild(ext)
            cellsOf(tr).forEach((tc, i) => {
                for (const attr of [ "gridSpan", "rowSpan", "hMerge", "vMerge" ])
                    tc.removeAttribute(attr)
                setBodyText(cellBody(tc), cellText(values[i] ?? ""), false, lang, true)
                for (const link of [ ...elements(tc, "a:hlinkClick"), ...elements(tc, "a:hlinkMouseOver") ])
                    link.parentNode?.removeChild(link)
            })
            tbl.insertBefore(tr, anchor)
        }
    }
    for (const cell of set.cells ?? []) {
        const tc = cellsOf(rowsOf(tbl)[cell.row] as Element)[cell.col]
        if (tc === undefined)
            throw new PptcError("E_ENGINE", `table cell ${cell.row},${cell.col} not found`)
        setBodyText(cellBody(tc), cell.text, false, lang, true)
    }
    const after = tableHeight(tbl)
    const ext = children(transformOf(shape), [ "a:ext" ])[0]
    if (after !== before && ext !== undefined && after > 0)
        ext.setAttribute("cy", String(after))
}

/**  the content hash of a part, null when it is missing  */
const partDigest = async (pkg: Package, part: string | null): Promise<string | null> => {
    const file = part === null ? null : pkg.zip.file(part)
    return file === null || file === undefined ? null : digestBytes(await file.async("nodebuffer"))
}

/**  an image an op replaced: the content hashes of the image parts the
     element showed and of the new one (null: the element went without a
     successor, a removed footer-band placeholder)  */
export interface ImageSwap {
    from: string[]
    to: string | null
}

/**  replace the image of a picture: a new media part, the blip (and an SVG
     extension) pointing at it, a crop cut to fit the frame centered  */
const replaceImage = async (pkg: Package, slidePart: string, rels: Document, shape: Element, file: string): Promise<ImageSwap> => {
    const doc = shape.ownerDocument as Document
    const blipFill = children(shape, [ "p:blipFill" ])[0] ?? null
    const blip = children(blipFill, [ "a:blip" ])[0] ?? null
    if (blipFill === null || blip === null)
        throw new PptcError("E_ADDR_NOTFOUND", `slide ${slidePart}: the picture has no image`)
    const targets = new Map(pkg.targets(slidePart, rels).map((t) => [ t.id, t.target ]))
    const svgBlips = elements(blip, "asvg:svgBlip")
    const from: string[] = []
    for (const id of [ blip.getAttributeNS(NS_R, "embed"), ...svgBlips.map((s) => s.getAttributeNS(NS_R, "embed")) ]) {
        const hash = await partDigest(pkg, targets.get(id ?? "") ?? null)
        if (hash !== null)
            from.push(hash)
    }

    const bytes = readFileSync(file)
    const ext = path.extname(file).slice(1).toLowerCase() || "png"
    const media = pkg.alloc.part("ppt/media", "image", `.${ext}`)
    pkg.write(media, bytes)
    await pkg.typePart(media, "", media)
    const rid = pkg.addRel(rels, `${REL}/image`, `../media/${path.posix.basename(media)}`)
    blip.setAttributeNS(NS_R, "r:embed", rid)
    blip.removeAttributeNS(NS_R, "link")

    /*  an SVG shows through the SVG extension (the blip carries the same
        bytes, as el.add's generator writes them); other images drop it  */
    for (const svg of svgBlips) {
        const extEl = svg.parentNode as Element
        const extLst = extEl.parentNode as Element
        extLst.removeChild(extEl)
        if (children(extLst, [ "a:ext" ]).length === 0)
            extLst.parentNode?.removeChild(extLst)
    }
    if (ext === "svg") {
        const extLst = ensureChild(blip, "a:extLst", [])
        const extEl = doc.createElementNS(NS_A, "a:ext")
        extEl.setAttribute("uri", SVG_EXT)
        const svg = doc.createElementNS(NS_ASVG, "asvg:svgBlip")
        svg.setAttributeNS(NS_R, "r:embed", rid)
        extEl.appendChild(svg)
        extLst.appendChild(extEl)
    }

    /*  the frame stays; the image covers it, cut off centered to the
        frame's aspect (an old crop is cut anew) -- an unknown image size
        or frame stretches, a tiled fill stays as it is  */
    const box = boxOf(transformOf(shape))
    if (box !== null && children(blipFill, [ "a:tile" ]).length === 0) {
        const crop = fitImage(box, imageSize(bytes), "crop").crop
        let srcRect = children(blipFill, [ "a:srcRect" ])[0]
        if (srcRect === undefined && crop !== null)
            srcRect = ensureChild(blipFill, "a:srcRect", [ "a:tile", "a:stretch" ])
        if (srcRect !== undefined) {
            for (const a of [ "l", "t", "r", "b" ])
                srcRect.removeAttribute(a)
            for (const [ side, value ] of Object.entries(crop ?? {}))
                if (value !== 0)
                    srcRect.setAttribute(side, String(value))
        }
    }
    return { from, to: digestBytes(bytes) }
}

/**
 *  Apply one el.set to an existing top-level element.
 *
 *  @param pkg - the deck package
 *  @param slidePart - the slide part
 *  @param rels - the slide's relationships
 *  @param shape - the element
 *  @param edit - the planned edit
 *  @param lang - the language of runs that state none
 *  @param listStyle - the list style the element inherits when it is a
 *                     placeholder (bullets by level), null when unknown
 *  @returns the image the edit replaced, null when it replaced none
 */
export const applyEdit = async (pkg: Package, slidePart: string, rels: Document, shape: Element,
    edit: PlannedEdit, lang: string, listStyle: ListStyle | null = null): Promise<ImageSwap | null> => {
    const { set } = edit
    if (set.text !== undefined)
        setShapeText(shape, set.text, false, lang, true, listStyle)
    if (set.fontColor !== undefined)
        for (const rPr of runProps(shape))
            setSolidFill(rPr, set.fontColor, [ "a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn",
                "a:uFillTx", "a:uFill", "a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver",
                "a:rtl", "a:extLst" ])

    /*  the frame first: a new image is cut to fit the final frame  */
    if (set.frame !== undefined)
        setFrame(shape, set.frame)
    const swap = set.image === undefined ? null : await replaceImage(pkg, slidePart, rels, shape, set.image)
    if (set.fill !== undefined || set.transparency !== undefined) {
        const props = shapeProps(shape)
        const color = (): Element | null =>
            props === null ? null : children(children(props, [ "a:solidFill" ])[0] ?? null, COLORS)[0] ?? null

        /*  a new fill keeps the element's transparency  */
        const alpha = children(color(), [ "a:alpha" ])[0]?.getAttribute("val") ?? null
        if (props !== null && set.fill !== undefined)
            setSolidFill(props, set.fill, [ "a:ln", "a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst" ])
        const transparency = set.transparency ?? (alpha === null ? 0 : (100000 - Number(alpha)) / 1000)
        setAlpha(color(), transparency)
    }
    if (set.border !== undefined)
        setSolidFill(outlineOf(shape), set.border, [ "a:prstDash", "a:custDash", "a:round", "a:bevel", "a:miter",
            "a:headEnd", "a:tailEnd", "a:extLst" ])
    if (set.borderPt !== undefined)
        outlineOf(shape).setAttribute("w", String(Math.round(set.borderPt * 12700)))
    if (set.altText !== undefined)
        setDescription(shape, "descr", set.altText)
    if (set.title !== undefined)
        setDescription(shape, "title", set.title)
    if (set.hyperlink !== undefined)
        setClickLink(shape, set.hyperlink)
    if (set.cells !== undefined || set.addRows !== undefined || set.removeRows !== undefined)
        editTable(shape, set, lang)
    return swap
}

/**
 *  Apply a text.replace to the content a slide carried before the apply:
 *  every text body of its shapes (group members included) and every table
 *  cell, but what an earlier op of the document replaced or removed.
 *
 *  @param slide - the slide DOM (changed in place)
 *  @param replace - the planned replace
 *  @returns the number of replacements and the relationship ids the
 *           removed runs referenced
 */
export const replaceExisting = (slide: Document, replace: PlannedReplace): { count: number, refs: string[] } => {
    const spTree = firstElement(slide, "p:spTree")
    let count = 0
    const refs: string[] = []
    const take = (done: { count: number, dropped: Element[] }): void => {
        count += done.count
        for (const run of done.dropped)
            refs.push(...relRefs(run).map((ref) => ref.attr.value))
    }
    for (let n = spTree?.firstChild ?? null; n !== null; n = n.nextSibling) {
        if (n.nodeType !== 1)
            continue
        const shape = n as Element
        const name = firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? ""
        const phType = shape.nodeName === "p:sp" ? firstElement(shape, "p:ph")?.getAttribute("type") ?? null : null
        if (replace.skip.includes(name) || (phType !== null && replace.skip.includes(`@${phType}`)))
            continue
        for (const txBody of elements(shape, "p:txBody"))
            take(replaceInBody(txBody, replace))
        for (const tbl of elements(shape, "a:tbl"))
            rowsOf(tbl).forEach((tr, row) => {
                cellsOf(tr).forEach((tc, col) => {
                    const body = children(tc, [ "a:txBody" ])[0]
                    if (body !== undefined && !replace.skipCells.some((c) => c.name === name && c.row === row
                        && (c.col === col || c.col === -1)))
                        take(replaceInBody(body, replace))
                })
            })
    }
    return { count, refs }
}

/**
 *  Apply a text.replace to speaker notes: the text of the notes body
 *  placeholder.
 *
 *  @param notes - the notes DOM (changed in place)
 *  @param replace - the planned replace
 *  @returns the number of replacements
 */
export const replaceNotes = (notes: Document, replace: PlannedReplace): number => {
    let count = 0
    for (const sp of elements(notes, "p:sp"))
        if (firstElement(sp, "p:ph")?.getAttribute("type") === "body")
            for (const txBody of elements(sp, "p:txBody"))
                count += replaceInBody(txBody, replace).count
    return count
}
