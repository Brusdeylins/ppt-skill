/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/reader: read-only OOXML access. Opens .pptx/.potx archives and
**  extracts the domain read models: TemplateInfo (layouts, theme, capacity)
**  and DeckState (slides, shapes, rev token). This module never writes.
*/

import JSZip from "jszip"
import { readFileSync } from "node:fs"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../infra/errors.js"
import { contentHash, requireFile } from "../infra/fs.js"
import { estimateCapacity } from "../core/describe/capacity.js"
import { regionWithin, coverageFraction } from "../core/describe/position.js"
import { elements, firstElement, drawingText, parseXml } from "./xml.js"
import { emuToInch, isFooterKind } from "../core/model.js"
import type {
    DeckState, Frame, Layout, ParaInfo, Placeholder, PlaceholderKind,
    RunInfo, ShapeInfo, SlideInfo, TemplateInfo
} from "../core/model.js"

/**  An opened OOXML archive with cached XML parsing.  */
export class DeckArchive {
    private readonly cache = new Map<string, Document>()

    private constructor(
        /**  underlying zip archive  */
        readonly zip: JSZip,
        /**  absolute source file path  */
        readonly file: string
    ) {}

    /**
     *  Open a .pptx or .potx file.
     *
     *  @param file - path of the archive
     *  @returns the opened archive
     *  @throws PptcError E_FILE when missing or not a zip
     */
    static async open(file: string): Promise<DeckArchive> {
        requireFile(file, "presentation file")
        try {
            const zip = await JSZip.loadAsync(readFileSync(file))
            return new DeckArchive(zip, path.resolve(file))
        }
        catch {
            throw new PptcError("E_FILE", `'${file}' is not a readable pptx/potx archive`)
        }
    }

    /**
     *  Raw string content of an archive part.
     *
     *  @param part - part path inside the archive
     *  @returns content, or null when the part does not exist
     */
    async text(part: string): Promise<string | null> {
        const f = this.zip.file(part)
        return f === null ? null : await f.async("string")
    }

    /**
     *  Parsed XML of an archive part (cached).
     *
     *  @param part - part path inside the archive
     *  @returns DOM document
     *  @throws PptcError E_TEMPLATE when the part is missing
     */
    async xml(part: string): Promise<Document> {
        const cached = this.cache.get(part)
        if (cached !== undefined)
            return cached
        const text = await this.text(part)
        if (text === null)
            throw new PptcError("E_TEMPLATE", `archive part missing: ${part}`)
        const doc = parseXml(text)
        this.cache.set(part, doc)
        return doc
    }
}

/**  normalize an OOXML placeholder `type` attribute to a PlaceholderKind  */
const phKind = (type: string | null): PlaceholderKind => {
    switch (type) {
        case "title":
        case "ctrTitle": return "title"
        case "subTitle": return "subtitle"
        case "pic": return "picture"
        case "ftr": return "footer"
        case "sldNum": return "slideNumber"
        case "dt": return "date"
        case null:
        case "body": return "body"
        default: return "other"
    }
}

/**  extract the frame of a shape from its `a:xfrm` (autoshapes/placeholders)
     or `p:xfrm` (graphicFrame: tables, charts), null when inherited  */
const shapeFrame = (sp: Element): Frame | null => {
    const xfrm = firstElement(sp, "a:xfrm") ?? firstElement(sp, "p:xfrm")
    if (xfrm === null)
        return null
    const off = firstElement(xfrm, "a:off")
    const ext = firstElement(xfrm, "a:ext")
    if (off === null || ext === null)
        return null
    return {
        x: emuToInch(Number(off.getAttribute("x"))),
        y: emuToInch(Number(off.getAttribute("y"))),
        w: emuToInch(Number(ext.getAttribute("cx"))),
        h: emuToInch(Number(ext.getAttribute("cy")))
    }
}

/**  read `p:ph` facts of a shape, null when the shape is no placeholder  */
const phFacts = (sp: Element): { idx: number, kind: PlaceholderKind } | null => {
    const ph = firstElement(sp, "p:ph")
    if (ph === null)
        return null
    return {
        idx: Number(ph.getAttribute("idx") ?? "0"),
        kind: phKind(ph.getAttribute("type"))
    }
}

/**  resolve the layout part paths of all masters, in master order  */
const layoutPartsInOrder = async (archive: DeckArchive): Promise<{ part: string, master: string }[]> => {
    const presRels = await archive.xml("ppt/_rels/presentation.xml.rels")
    const pres = await archive.xml("ppt/presentation.xml")
    const relTarget = (rels: Document, rid: string): string | null => {
        for (const rel of elements(rels, "Relationship"))
            if (rel.getAttribute("Id") === rid)
                return rel.getAttribute("Target")
        return null
    }
    const parts: { part: string, master: string }[] = []
    for (const masterId of elements(pres, "p:sldMasterId")) {
        const target = relTarget(presRels, masterId.getAttribute("r:id") ?? "")
        if (target === null)
            continue
        const masterPart = path.posix.join("ppt", target)
        const master = await archive.xml(masterPart)
        const masterRels = await archive.xml(
            path.posix.join(path.posix.dirname(masterPart), "_rels", `${path.posix.basename(masterPart)}.rels`))
        for (const layoutId of elements(master, "p:sldLayoutId")) {
            const lt = relTarget(masterRels, layoutId.getAttribute("r:id") ?? "")
            if (lt !== null)
                parts.push({
                    part: path.posix.normalize(path.posix.join("ppt/slideMasters", lt)),
                    master: masterPart
                })
        }
    }
    return parts
}

/**  placeholder frames of a master, keyed by "kind:idx" and "kind"  */
const masterFrameMap = async (archive: DeckArchive, masterPart: string): Promise<Map<string, Frame>> => {
    const map = new Map<string, Frame>()
    const master = await archive.xml(masterPart)
    for (const sp of elements(master, "p:sp")) {
        const ph = phFacts(sp)
        const frame = ph === null ? null : shapeFrame(sp)
        if (ph === null || frame === null)
            continue
        map.set(`${ph.kind}:${ph.idx}`, frame)
        if (!map.has(ph.kind))
            map.set(ph.kind, frame)
    }
    return map
}

/**  read default font sizes (pt) for title and body from the master text styles  */
const masterFontSizes = async (archive: DeckArchive): Promise<{ title: number, body: number }> => {
    const fallback = { title: 36, body: 18 }
    try {
        const master = await archive.xml("ppt/slideMasters/slideMaster1.xml")
        const styleSize = (styleTag: string): number | null => {
            const style = firstElement(master, styleTag)
            if (style === null)
                return null
            const lvl1 = firstElement(style, "a:lvl1pPr")
            const rPr = lvl1 === null ? null : firstElement(lvl1, "a:defRPr")
            const sz = rPr?.getAttribute("sz")
            return sz == null ? null : Number(sz) / 100
        }
        return {
            title: styleSize("p:titleStyle") ?? fallback.title,
            body: styleSize("p:bodyStyle") ?? fallback.body
        }
    }
    catch {
        return fallback
    }
}

/**  read the slide dimensions from presentation.xml (EMU to inch)  */
const readSlideSize = async (archive: DeckArchive): Promise<{ w: number, h: number }> => {
    const pres = await archive.xml("ppt/presentation.xml")
    const sldSz = firstElement(pres, "p:sldSz")
    return {
        w: emuToInch(Number(sldSz?.getAttribute("cx") ?? "12192000")),
        h: emuToInch(Number(sldSz?.getAttribute("cy") ?? "6858000"))
    }
}

/**  read theme fonts (major/minor) and the color scheme from the first theme part  */
const readTheme = async (
    archive: DeckArchive
): Promise<{ fonts: { major: string, minor: string }, colors: Record<string, string> }> => {
    const themePart = Object.keys(archive.zip.files)
        .find((f) => /^ppt\/theme\/theme\d+\.xml$/.test(f)) ?? "ppt/theme/theme1.xml"
    const theme = await archive.xml(themePart)
    const fontOf = (tag: string): string =>
        firstElement(firstElement(theme, tag) ?? theme, "a:latin")?.getAttribute("typeface") ?? "Calibri"
    const colors: Record<string, string> = {}
    const scheme = firstElement(theme, "a:clrScheme")
    if (scheme !== null && scheme.childNodes !== null)
        for (let i = 0; i < scheme.childNodes.length; i++) {
            const node = scheme.childNodes.item(i) as Element
            if (node.nodeType !== 1)
                continue
            const name = node.nodeName.replace("a:", "")
            const val = firstElement(node, "a:srgbClr")?.getAttribute("val")
                ?? firstElement(node, "a:sysClr")?.getAttribute("lastClr")
            if (val !== null && val !== undefined)
                colors[name] = val.toUpperCase()
        }
    return { fonts: { major: fontOf("a:majorFont"), minor: fontOf("a:minorFont") }, colors }
}

/**  read all layouts with resolved placeholder geometry, capacity and
     picture-overlay coverage; geometry is inherited from the master when a
     layout placeholder carries no own xfrm  */
const readLayouts = async (archive: DeckArchive): Promise<Layout[]> => {
    const sizes = await masterFontSizes(archive)
    const layouts: Layout[] = []
    const parts = await layoutPartsInOrder(archive)
    const masterFrames = new Map<string, Map<string, Frame>>()
    for (let index = 0; index < parts.length; index++) {
        const { part, master } = parts[index] as { part: string, master: string }
        if (!masterFrames.has(master))
            masterFrames.set(master, await masterFrameMap(archive, master))
        const inherited = masterFrames.get(master) as Map<string, Frame>
        const resolveFrame = (sp: Element, kind: PlaceholderKind, idx: number): Frame | null =>
            shapeFrame(sp) ?? inherited.get(`${kind}:${idx}`) ?? inherited.get(kind) ?? null
        const doc = await archive.xml(part)
        const name = firstElement(doc, "p:cSld")?.getAttribute("name") ?? `Layout ${index}`
        const placeholders: Placeholder[] = []
        const reserved: Frame[] = []
        for (const sp of elements(doc, "p:sp")) {
            const ph = phFacts(sp)
            if (ph === null)
                continue
            if (isFooterKind(ph.kind)) {
                const f = resolveFrame(sp, ph.kind, ph.idx)
                if (f !== null)
                    reserved.push(f)
                continue
            }
            const frame = resolveFrame(sp, ph.kind, ph.idx)
            const fontSize = ph.kind === "title" ? sizes.title : sizes.body
            placeholders.push({
                idx: ph.idx,
                kind: ph.kind,
                name: firstElement(sp, "p:cNvPr")?.getAttribute("name") ?? `Placeholder ${ph.idx}`,
                frame,
                capacity: frame !== null && ph.kind !== "picture"
                    ? estimateCapacity(frame, fontSize)
                    : null
            })
        }
        /*  picture placeholders: record which text shapes sit on top
            (title over a fullscreen background, footer over the image
            edge, ...) so image prompts keep those regions calm  */
        for (const pic of placeholders) {
            if (pic.kind !== "picture" || pic.frame === null)
                continue
            const overlays: { name: string, region: string }[] = []
            const covering: Frame[] = []
            for (const other of placeholders) {
                if (other === pic || other.kind === "picture" || other.frame === null)
                    continue
                const region = regionWithin(pic.frame, other.frame)
                if (region !== null) {
                    overlays.push({ name: other.name, region })
                    covering.push(other.frame)
                }
            }
            for (const f of reserved) {
                const region = regionWithin(pic.frame, f)
                if (region !== null) {
                    overlays.push({ name: "footer/slide-number", region })
                    covering.push(f)
                }
            }
            if (overlays.length > 0) {
                pic.overlays = overlays
                pic.coverage = coverageFraction(pic.frame, covering)
            }
        }
        layouts.push({ index, name, placeholders, reserved })
    }
    return layouts
}

/**  read PowerPoint p15 drawing guides from the master: `pos` is in 1/8 pt
     (inch = pos/576); orient="horz" is a horizontal line at a Y coordinate,
     otherwise a vertical line at an X coordinate  */
const readGuides = async (archive: DeckArchive): Promise<{ horizontal: number[], vertical: number[] }> => {
    const guides = { horizontal: [] as number[], vertical: [] as number[] }
    try {
        const master = await archive.xml("ppt/slideMasters/slideMaster1.xml")
        for (const g of elements(master, "p15:guide")) {
            const pos = Number(g.getAttribute("pos"))
            if (!Number.isFinite(pos))
                continue
            if (g.getAttribute("orient") === "horz")
                guides.horizontal.push(pos / 576)
            else
                guides.vertical.push(pos / 576)
        }
    }
    catch {
        /*  no master or no guides -- leave the lists empty  */
    }
    const uniqSort = (xs: number[]): number[] =>
        Array.from(new Set(xs.map((v) => Math.round(v * 100) / 100))).sort((a, b) => a - b)
    return { horizontal: uniqSort(guides.horizontal), vertical: uniqSort(guides.vertical) }
}

/**  derive the content area: the largest body placeholder with its edges
     snapped to the nearest guides -- the clean target for `el.add` on
     title-only layouts. Undefined when the template has no body placeholder  */
const deriveContentArea = (
    layouts: Layout[],
    guides: { horizontal: number[], vertical: number[] }
): Frame | undefined => {
    const bodies = layouts
        .flatMap((l) => l.placeholders)
        .filter((p) => p.kind === "body" && p.frame !== null)
        .map((p) => p.frame as Frame)
    if (bodies.length === 0)
        return undefined
    const biggest = bodies.reduce((a, b) => (a.w * a.h >= b.w * b.h ? a : b))
    const snap = (v: number, cand: number[]): number => {
        let best = v
        let bestD = 0.4
        for (const c of cand) {
            const d = Math.abs(c - v)
            if (d < bestD) {
                bestD = d
                best = c
            }
        }
        return best
    }
    const x1 = snap(biggest.x, guides.vertical)
    const x2 = snap(biggest.x + biggest.w, guides.vertical)
    const y1 = snap(biggest.y, guides.horizontal)
    const y2 = snap(biggest.y + biggest.h, guides.horizontal)
    return { x: x1, y: y1, w: Math.max(0, x2 - x1), h: Math.max(0, y2 - y1) }
}

/**
 *  Read template-wide data: slide size, theme, and all layouts with resolved
 *  placeholder geometry and capacity. Works on .potx and .pptx alike.
 *
 *  @param archive - opened archive
 *  @returns the resolved template info
 */
export const readTemplateInfo = async (archive: DeckArchive): Promise<TemplateInfo> => {
    const slideSize = await readSlideSize(archive)
    const { fonts, colors } = await readTheme(archive)
    const layouts = await readLayouts(archive)
    const guides = await readGuides(archive)
    const hasGuides = guides.horizontal.length + guides.vertical.length > 0
    const contentArea = deriveContentArea(layouts, guides)
    return {
        slideSize,
        fonts,
        colors,
        layouts,
        ...(hasGuides ? { guides } : {}),
        ...(contentArea !== undefined ? { contentArea } : {})
    }
}

/**  classify a non-placeholder shape element  */
const shapeType = (el: Element): ShapeInfo["type"] => {
    switch (el.nodeName) {
        case "p:pic": return "picture"
        case "p:cxnSp": return "connector"
        case "p:grpSp": return "group"
        case "p:graphicFrame": {
            const uri = firstElement(el, "a:graphicData")?.getAttribute("uri") ?? ""
            return uri.endsWith("/table") ? "table" : uri.endsWith("/chart") ? "chart" : "other"
        }
        default: return "other"
    }
}

/**  extract table cell texts from a graphicFrame  */
const tableCells = (frame: Element): string[][] =>
    elements(frame, "a:tr").map((tr) => elements(tr, "a:tc").map((tc) => drawingText(tc)))

/**  first DIRECT child element of `parent` whose tag is one of `names`  */
const directChild = (parent: Element | null, names: string[]): Element | null => {
    if (parent === null || parent.childNodes === null)
        return null
    for (let i = 0; i < parent.childNodes.length; i++) {
        const n = parent.childNodes.item(i) as Element
        if (n.nodeType === 1 && names.includes(n.nodeName))
            return n
    }
    return null
}

/**  resolve a DrawingML color container (a:solidFill, a:ln, ...) to an RRGGBB
     hex; `a:schemeClr` references are resolved against the theme color map  */
const colorOf = (el: Element | null, colors: Record<string, string>): string | undefined => {
    if (el === null)
        return undefined
    const srgb = firstElement(el, "a:srgbClr")?.getAttribute("val")
    if (srgb !== null && srgb !== undefined)
        return srgb.toUpperCase()
    const scheme = firstElement(el, "a:schemeClr")?.getAttribute("val")
    if (scheme !== null && scheme !== undefined) {
        const alias: Record<string, string> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" }
        return colors[alias[scheme] ?? scheme]
    }
    return undefined
}

/**  autoshape style of a `p:sp` (preset, fill, border, first-run font),
     theme colors resolved -- the read mirror of el.add's shape vocabulary  */
const shapeStyle = (sp: Element, colors: Record<string, string>): Partial<ShapeInfo> => {
    const out: Partial<ShapeInfo> = {}
    const spPr = firstElement(sp, "p:spPr")
    const prst = firstElement(sp, "a:prstGeom")?.getAttribute("prst")
    if (prst !== null && prst !== undefined)
        out.shape = prst
    const fill = colorOf(directChild(spPr, ["a:solidFill"]), colors)
    if (fill !== undefined)
        out.fill = fill
    const ln = directChild(spPr, ["a:ln"])
    const border = colorOf(directChild(ln, ["a:solidFill"]), colors)
    if (border !== undefined)
        out.border = border
    const lnW = ln?.getAttribute("w")
    if (lnW !== null && lnW !== undefined && lnW !== "")
        out.borderPt = Math.round((Number(lnW) / 12700) * 100) / 100
    const rPr = firstElement(sp, "a:rPr")
    const sz = rPr?.getAttribute("sz")
    if (sz !== null && sz !== undefined)
        out.fontSize = Number(sz) / 100
    const fontColor = colorOf(directChild(rPr, ["a:solidFill"]), colors)
    if (fontColor !== undefined)
        out.fontColor = fontColor
    const face = rPr === null ? null : firstElement(rPr, "a:latin")?.getAttribute("typeface")
    if (face !== null && face !== undefined)
        out.fontFace = face
    return out
}

/**  read one `a:r` run into the RunInfo shape (mirrors a slide.fill run)  */
const runStyle = (r: Element, colors: Record<string, string>): RunInfo => {
    const out: RunInfo = { text: firstElement(r, "a:t")?.textContent ?? "" }
    const rPr = firstElement(r, "a:rPr")
    if (rPr === null)
        return out
    const face = firstElement(rPr, "a:latin")?.getAttribute("typeface")
    if (face !== null && face !== undefined && face !== "")
        out.font = face
    const sz = rPr.getAttribute("sz")
    if (sz !== null && sz !== "")
        out.size = Number(sz) / 100
    if (rPr.getAttribute("b") === "1")
        out.bold = true
    if (rPr.getAttribute("i") === "1")
        out.italic = true
    const color = colorOf(directChild(rPr, ["a:solidFill"]), colors)
    if (color !== undefined)
        out.color = color
    return out
}

/**  Break a text body into paragraphs and runs, but only when it carries
     explicit run formatting (a per-run font, bold or italic) -- e.g. a
     monospace code block or an emphasized word. Returns undefined for plain
     uniform text, where the flat `text` already says everything and the
     extra structure would only be noise.  */
const richParagraphs = (txBody: Element, colors: Record<string, string>): ParaInfo[] | undefined => {
    const paras: ParaInfo[] = []
    let formatted = false
    for (const p of elements(txBody, "a:p")) {
        const runs = elements(p, "a:r").map((r) => runStyle(r, colors))
        if (runs.some((rn) => rn.font !== undefined || rn.bold === true || rn.italic === true))
            formatted = true
        const para: ParaInfo = { runs }
        const pPr = firstElement(p, "a:pPr")
        if (pPr !== null) {
            const lvl = pPr.getAttribute("lvl")
            if (lvl !== null && lvl !== "" && Number(lvl) > 0)
                para.level = Number(lvl)
            if (directChild(pPr, ["a:buChar", "a:buAutoNum"]) !== null)
                para.bullet = true
            else if (directChild(pPr, ["a:buNone"]) !== null)
                para.bullet = false
        }
        paras.push(para)
    }
    return formatted ? paras : undefined
}

/**
 *  Read the full deck state: slides, shapes, notes and the rev token.
 *
 *  @param archive - opened archive of a .pptx deck
 *  @returns the deck read model (the agent's source of truth)
 */
export const readDeckState = async (archive: DeckArchive): Promise<DeckState> => {
    const info = await readTemplateInfo(archive)
    const layoutIndexByName = new Map(info.layouts.map((l) => [l.name, l.index]))
    const pres = await archive.xml("ppt/presentation.xml")
    const presRels = await archive.xml("ppt/_rels/presentation.xml.rels")
    const relMap = new Map(elements(presRels, "Relationship")
        .map((r) => [r.getAttribute("Id") ?? "", r.getAttribute("Target") ?? ""]))

    const slides: SlideInfo[] = []
    const hashParts: string[] = []
    const sldIds = elements(pres, "p:sldId")
    for (let index = 0; index < sldIds.length; index++) {
        const sldId = sldIds[index] as Element
        const id = Number(sldId.getAttribute("id"))
        const target = relMap.get(sldId.getAttribute("r:id") ?? "") ?? ""
        const part = path.posix.normalize(path.posix.join("ppt", target))
        const partBase = path.posix.basename(part)
        const slideXmlText = await archive.text(part) ?? ""
        hashParts.push(`${id}:${slideXmlText}`)
        const doc = parseXml(slideXmlText)

        /*  layout of this slide  */
        const slideRels = await archive.xml(`ppt/slides/_rels/${partBase}.rels`)
        const slideRelMap = new Map(elements(slideRels, "Relationship")
            .map((r) => [r.getAttribute("Id") ?? "", r.getAttribute("Target") ?? ""]))
        const layoutTarget = elements(slideRels, "Relationship")
            .find((r) => (r.getAttribute("Type") ?? "").endsWith("/slideLayout"))?.getAttribute("Target")
        let layoutName = ""
        let layoutIndex = -1
        if (layoutTarget !== null && layoutTarget !== undefined) {
            const layoutPart = path.posix.normalize(path.posix.join("ppt/slides", layoutTarget))
            const layoutDoc = await archive.xml(layoutPart)
            layoutName = firstElement(layoutDoc, "p:cSld")?.getAttribute("name") ?? ""
            layoutIndex = layoutIndexByName.get(layoutName) ?? -1
        }

        /*  shapes  */
        const shapes: ShapeInfo[] = []
        let title: string | null = null
        const spTree = firstElement(doc, "p:spTree")
        if (spTree !== null && spTree.childNodes !== null)
            for (let i = 0; i < spTree.childNodes.length; i++) {
                const node = spTree.childNodes.item(i) as Element
                if (node.nodeType !== 1)
                    continue
                if (!["p:sp", "p:pic", "p:graphicFrame", "p:cxnSp", "p:grpSp"].includes(node.nodeName))
                    continue
                const cNvPr = firstElement(node, "p:cNvPr")
                const ph = node.nodeName === "p:sp" || node.nodeName === "p:pic" ? phFacts(node) : null
                const txBody = node.nodeName === "p:sp" ? firstElement(node, "p:txBody") : null
                const text = txBody === null ? null : drawingText(txBody) || null
                if (ph !== null && (ph.kind === "title" || ph.idx === 0))
                    title = text
                const shape: ShapeInfo = {
                    name: cNvPr?.getAttribute("name") ?? "",
                    id: Number(cNvPr?.getAttribute("id") ?? "0"),
                    type: node.nodeName === "p:sp"
                        ? (ph !== null ? "placeholder" : (firstElement(node, "a:prstGeom") !== null ? "shape" : "textbox"))
                        : shapeType(node),
                    placeholderIdx: ph?.idx ?? null,
                    placeholderKind: ph?.kind ?? null,
                    frame: shapeFrame(node),
                    text
                }
                if (node.nodeName === "p:sp")
                    Object.assign(shape, shapeStyle(node, info.colors))
                if (txBody !== null) {
                    const paras = richParagraphs(txBody, info.colors)
                    if (paras !== undefined)
                        shape.paragraphs = paras
                }
                if (node.nodeName === "p:pic") {
                    const embed = firstElement(node, "a:blip")?.getAttribute("r:embed")
                    const tgt = embed === null || embed === undefined ? undefined : slideRelMap.get(embed)
                    if (tgt !== undefined && tgt !== "")
                        shape.image = path.posix.basename(tgt)
                }
                if (shape.type === "table") {
                    shape.table = tableCells(node)
                    const grid = firstElement(node, "a:tblGrid")
                    if (grid !== null)
                        shape.colWidths = elements(grid, "a:gridCol")
                            .map((c) => emuToInch(Number(c.getAttribute("w"))))
                }
                shapes.push(shape)
            }

        /*  speaker notes  */
        const notesTarget = elements(slideRels, "Relationship")
            .find((r) => (r.getAttribute("Type") ?? "").endsWith("/notesSlide"))?.getAttribute("Target")
        let notes: string | null = null
        if (notesTarget !== null && notesTarget !== undefined) {
            const notesPart = path.posix.normalize(path.posix.join("ppt/slides", notesTarget))
            const notesText = await archive.text(notesPart)
            if (notesText !== null) {
                const notesDoc = parseXml(notesText)
                const body = elements(notesDoc, "p:sp")
                    .find((sp) => phFacts(sp)?.kind === "body")
                notes = body === undefined ? null : (drawingText(body) || null)
            }
        }
        /*  slide visibility: `show="0"` on the <p:sld> root is PowerPoint's
            "Hide Slide". Absent (or "1") means visible.  */
        const hidden = doc.documentElement?.getAttribute("show") === "0"
        slides.push({ id, index, title, layoutName, layoutIndex, shapes, notes, hidden, part })
    }

    /*  custom document properties: the deck's self-describing memory  */
    const customProps: Record<string, string> = {}
    const customXml = await archive.text("docProps/custom.xml")
    if (customXml !== null)
        for (const p of elements(parseXml(customXml), "property")) {
            const name = p.getAttribute("name")
            if (name !== null && name !== "")
                customProps[name] = firstElement(p, "vt:lpwstr")?.textContent ?? ""
        }

    return {
        file: archive.file,
        rev: contentHash(...hashParts, JSON.stringify(slides.map((s) => s.id))),
        slideSize: info.slideSize,
        slides,
        customProps
    }
}
