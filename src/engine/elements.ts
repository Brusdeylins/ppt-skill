/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/elements: translate validated ElementSpec payloads into PptxGenJS
**  calls on a generator slide; engine/inplace/transplant moves the result
**  into the target slide. What PptxGenJS cannot write is finished on its
**  output here: theme colors (written as color tokens first), one
**  paragraph property set per paragraph, picture crops and the number
**  format of a scatter chart's x axis.
*/

import { readFileSync } from "node:fs"
import type { Document, Element } from "@xmldom/xmldom"
import { EMU_PER_INCH, inchToEmu } from "../core/model.js"
import { layoutTable } from "../core/table-grid.js"
import { isSchemeColor, type ChartData, type ElementSpec, type Paragraph, type RichText,
    type SchemeColor, type TableCell, type TextDirection } from "../schema/payloads.js"
import { fitImage, imageSize, type ImageSize } from "./image-size.js"
import { BASELINE, paragraphRun } from "./text.js"
import { NS_A, elements, firstElement } from "./xml.js"

/*  the PptxGenJS surface used by pptc, typed structurally  */

/**  the subset of the PptxGenJS slide API used by pptc  */
export interface GenSlide {
    addText(text: unknown, opts: Record<string, unknown>): void
    addTable(rows: unknown[][], opts: Record<string, unknown>): void
    addChart(type: unknown, data: unknown[], opts: Record<string, unknown>): void
    addShape(type: unknown, opts: Record<string, unknown>): void
    addImage(opts: Record<string, unknown>): void
}

/**  the subset of the PptxGenJS root API used by pptc  */
export interface GenRoot {
    ChartType: Record<string, unknown>
    ShapeType: Record<string, unknown>
}

/**  OOXML `vert` value of each text direction (bodyPr and tcPr)  */
export const TEXT_DIRECTION_VERT: Record<TextDirection, string> = {
    horizontal: "horz",
    rotate90: "vert",
    rotate270: "vert270",
    stacked: "wordArtVert"
}

/**
 *  Theme colors for PptxGenJS. PptxGenJS writes only some theme slots
 *  and chart series colors always as RGB, so every theme color goes in as
 *  a hex token no element uses literally and is turned into its
 *  `a:schemeClr` on the generator output (`finish`).
 */
export class GenColors {
    /**  token (RRGGBB) to theme slot  */
    readonly tokens = new Map<string, SchemeColor>()
    private readonly byName = new Map<SchemeColor, string>()
    private readonly base: number

    /**
     *  Choose the token range for a set of element specs.
     *
     *  @param specs - every element of the apply (their hex colors and
     *                 texts must not collide with a token)
     */
    constructor(specs: ElementSpec[]) {
        const used = new Set([ ...JSON.stringify(specs).matchAll(/[0-9A-Fa-f]{6}/g) ].map((m) => m[0].toUpperCase()))
        let base = 0x0C0FE0
        while ([ ...Array(16).keys() ].some((i) => used.has(GenColors.hex(base + i))))
            base += 0x10
        this.base = base
    }

    /**  RRGGBB of a number  */
    private static hex(n: number): string {
        return n.toString(16).toUpperCase().padStart(6, "0")
    }

    /**
     *  The color to hand to PptxGenJS.
     *
     *  @param color - a validated color (RRGGBB or theme slot)
     *  @returns the hex value, or the token of a theme slot
     */
    gen(color: string): string {
        if (!isSchemeColor(color))
            return color
        let token = this.byName.get(color)
        if (token === undefined) {
            token = GenColors.hex(this.base + this.byName.size)
            this.byName.set(color, token)
            this.tokens.set(token, color)
        }
        return token
    }
}

/**
 *  Finish a generated slide or chart part: color tokens become their
 *  `a:schemeClr`, and every paragraph keeps only its first `a:pPr` --
 *  PptxGenJS writes one per run, which is invalid after the first run.
 *
 *  @param doc - the part DOM (changed in place)
 *  @param tokens - the color tokens of the apply
 *  @returns true when the part changed
 */
export const finishGenerated = (doc: Document, tokens: Map<string, SchemeColor>): boolean => {
    let changed = false
    for (const clr of elements(doc, "a:srgbClr")) {
        const slot = tokens.get((clr.getAttribute("val") ?? "").toUpperCase())
        if (slot === undefined)
            continue
        const scheme = doc.createElementNS(NS_A, "a:schemeClr")
        scheme.setAttribute("val", slot)
        while (clr.firstChild !== null)
            scheme.appendChild(clr.firstChild)
        clr.parentNode?.replaceChild(scheme, clr)
        changed = true
    }
    for (const p of elements(doc, "a:p"))
        for (const pPr of elements(p, "a:pPr").filter((el) => el.parentNode === p).slice(1)) {
            p.removeChild(pPr)
            changed = true
        }
    return changed
}

/**  map pptc chart types to PptxGenJS chart types with stacking options  */
const chartType = (gen: GenRoot, type: ChartData["type"]): { type: unknown, extra: Record<string, unknown> } => {
    const map: Record<ChartData["type"], [string, Record<string, unknown>]> = {
        bar: ["bar", { barDir: "bar" }],
        barStacked: ["bar", { barDir: "bar", barGrouping: "stacked" }],
        column: ["bar", { barDir: "col" }],
        columnStacked: ["bar", { barDir: "col", barGrouping: "stacked" }],
        line: ["line", {}],
        lineMarkers: ["line", { lineDataSymbol: "circle" }],
        pie: ["pie", {}],
        doughnut: ["doughnut", {}],
        area: ["area", {}],
        areaStacked: ["area", { barGrouping: "stacked" }],
        radar: ["radar", {}],
        scatter: ["scatter", { lineSize: 0 }],
        bubble: ["bubble", {}]
    }
    const entry = map[type]
    return { type: gen.ChartType[entry[0]], extra: entry[1] }
}

/**  convert rich text to the PptxGenJS text-run array form; every run
     carries its language (run, else paragraph, else the element's).
     Paragraph properties go on every run of the paragraph (PptxGenJS
     starts a new paragraph when they change), the bullet on its first
     run only (PptxGenJS starts a new paragraph at every bulleted run).  */
const genTextRuns = (text: RichText, lang: string, colors: GenColors): unknown => {
    if (typeof text === "string")
        return text
    const out: Record<string, unknown>[] = []
    for (const p of text as Paragraph[]) {
        const runs = p.runs ?? [ paragraphRun(p) ]
        runs.forEach((r, i) => {
            out.push({
                text: r.text,
                options: {
                    breakLine: i === runs.length - 1,
                    lang: r.lang ?? p.lang ?? lang,
                    ...(r.bold !== undefined && { bold: r.bold }),
                    ...(r.italic !== undefined && { italic: r.italic }),
                    ...(r.underline === true && { underline: { style: "sng" } }),
                    ...(r.strike === true && { strike: "sngStrike" }),
                    /*  PptxGenJS writes baseline * 50 (1/1000 %)  */
                    ...(r.superscript === true && { baseline: BASELINE.superscript / 50 }),
                    ...(r.subscript === true && { baseline: BASELINE.subscript / 50 }),
                    ...(r.size !== undefined && { fontSize: r.size }),
                    ...(r.font !== undefined && { fontFace: r.font }),
                    ...(r.color !== undefined && { color: colors.gen(r.color) }),
                    ...(r.highlight !== undefined && { highlight: colors.gen(r.highlight) }),
                    ...("hyperlink" in r && r.hyperlink !== undefined && { hyperlink: { url: r.hyperlink } }),
                    ...(p.level !== undefined && { indentLevel: p.level }),
                    ...(p.alignment !== undefined && { align: p.alignment }),
                    ...(p.spaceBefore !== undefined && { paraSpaceBefore: p.spaceBefore }),
                    ...(p.spaceAfter !== undefined && { paraSpaceAfter: p.spaceAfter }),
                    ...(p.lineSpacing !== undefined && { lineSpacingMultiple: p.lineSpacing }),
                    ...(i === 0 && p.bullet !== undefined && { bullet: p.bullet === "number" ? { type: "number" } : p.bullet })
                }
            })
        })
    }
    return out
}

/**  PptxGenJS options of a textbox or shape body: anchoring, inner margin
     (PptxGenJS: points as [left, right, bottom, top]) and text direction  */
const boxTextOptions = (el: Extract<ElementSpec, { type: "textbox" | "shape" }>): Record<string, unknown> => {
    const out: Record<string, unknown> = {}
    if (el.valign !== undefined)
        out.valign = el.valign
    if (el.margin !== undefined) {
        const [ l, t, r, b ] = typeof el.margin === "number" ? [ el.margin, el.margin, el.margin, el.margin ] : el.margin
        out.margin = [ l * 72, r * 72, b * 72, t * 72 ]
    }
    if (el.textDirection !== undefined)
        out.vert = TEXT_DIRECTION_VERT[el.textDirection]
    return out
}

/**  build PptxGenJS table rows including header styling, column
     alignment and styled, merged cells; PptxGenJS does not pass a
     table's language on to its cells, so every cell carries it  */
const genTableRows = (el: Extract<ElementSpec, { type: "table" }>, lang: string, colors: GenColors): unknown[][] => {
    const data = el.data
    const header = data.headers !== undefined ? 1 : 0
    const rows: unknown[][] = [ ...(data.headers !== undefined ? [ data.headers ] : []), ...data.rows ]
    const grid = layoutTable(rows)
    return rows.map((row, r) => row.map((cell, k) => {
        const styled = typeof cell === "object" ? cell as TableCell : null
        const options: Record<string, unknown> = { lang }
        if (r < header) {
            options.bold = true
            options.color = colors.gen(data.style?.headerFg ?? "FFFFFF")
            options.fill = { color: colors.gen(data.style?.headerBg ?? "333333") }
        }
        else if (data.style?.altRowBg !== undefined && (r - header) % 2 === 1)
            options.fill = { color: colors.gen(data.style.altRowBg) }
        const col = typeof grid === "string" ? k : grid.cells[r]?.[k]?.col ?? k
        const align = styled?.align ?? el.colAlign?.[col]
        if (align !== undefined)
            options.align = align
        if (styled !== null) {
            if (styled.bold !== undefined)
                options.bold = styled.bold
            if (styled.color !== undefined)
                options.color = colors.gen(styled.color)
            if (styled.fill !== undefined)
                options.fill = { color: colors.gen(styled.fill) }
            if (styled.valign !== undefined)
                options.valign = styled.valign
            if (styled.colspan !== undefined && styled.colspan > 1)
                options.colspan = styled.colspan
            if (styled.rowspan !== undefined && styled.rowspan > 1)
                options.rowspan = styled.rowspan
            if (styled.textDirection !== undefined)
                options.textDirection = TEXT_DIRECTION_VERT[styled.textDirection]
        }
        return { text: styled !== null ? genTextRuns(styled.text, lang, colors) : String(cell), options }
    }))
}

/**  default series color cycle when the payload specifies none  */
const DEFAULT_COLORS = ["4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646"]

/**  OOXML data-label positions of the pptc names  */
const LABEL_POS: Record<string, string> = {
    center: "ctr", insideEnd: "inEnd", insideBase: "inBase", outsideEnd: "outEnd",
    bestFit: "bestFit", left: "l", right: "r", above: "t", below: "b"
}

/**  PptxGenJS options of a chart's legend, labels and axes  */
const chartOptions = (c: ChartData): Record<string, unknown> => {
    const out: Record<string, unknown> = {
        showLegend: typeof c.legend === "string" || (c.legend ?? !["pie", "doughnut"].includes(c.type)),
        legendPos: typeof c.legend === "string" ? { top: "t", bottom: "b", left: "l", right: "r" }[c.legend] : "b"
    }
    if (c.labels !== undefined) {
        const percent = c.labels.percent === true
        out.showValue = c.labels.value ?? !percent
        out.showPercent = percent
        if (c.labels.format !== undefined)
            out.dataLabelFormatCode = c.labels.format
        /*  outside end is PowerPoint's own default of clustered bars, and
            PptxGenJS 4.0.1 drops it for them: written by leaving it out  */
        const clustered = (c.type === "bar" || c.type === "column") && c.stacked === undefined
        if (c.labels.position !== undefined && !(clustered && c.labels.position === "outsideEnd"))
            out.dataLabelPosition = LABEL_POS[c.labels.position]
    }
    if (c.stacked === "percent") {
        out.barGrouping = "percentStacked"
        out.valAxisLabelFormatCode = "0%"
    }
    const { x, y } = c.axes ?? {}
    if (x?.title !== undefined)
        Object.assign(out, { showCatAxisTitle: true, catAxisTitle: x.title })
    if (x?.min !== undefined)
        out.catAxisMinVal = x.min
    if (x?.max !== undefined)
        out.catAxisMaxVal = x.max
    if (y?.title !== undefined)
        Object.assign(out, { showValAxisTitle: true, valAxisTitle: y.title })
    if (y?.min !== undefined)
        out.valAxisMinVal = y.min
    if (y?.max !== undefined)
        out.valAxisMaxVal = y.max
    if (y?.format !== undefined)
        out.valAxisLabelFormatCode = y.format
    if (c.smooth !== undefined)
        out.lineSmooth = c.smooth
    if (c.holeSize !== undefined)
        out.holeSize = c.holeSize
    return out
}

/**
 *  Finish a generated chart part: PptxGenJS writes the y axis number
 *  format onto the x value axis of a scatter or bubble chart as well; the
 *  x axis gets its own (`General` unless `axes.x.format` is given).
 *
 *  @param doc - the chart part DOM (changed in place)
 *  @param chart - the chart payload
 *  @returns true when the part changed
 */
export const finishChart = (doc: Document, chart: ChartData): boolean => {
    if (chart.type !== "scatter" && chart.type !== "bubble")
        return false
    /*  PptxGenJS writes the x axis first  */
    const numFmt = elements(doc, "c:valAx").map((ax) => firstElement(ax, "c:numFmt"))[0] ?? null
    if (numFmt === null)
        return false
    const format = chart.axes?.x?.format
    numFmt.setAttribute("formatCode", format ?? "General")
    numFmt.setAttribute("sourceLinked", format !== undefined ? "0" : "1")
    return true
}

/**  the pixel size of an image file, null when unreadable  */
const fileImageSize = (file: string): ImageSize | null => {
    try {
        return imageSize(readFileSync(file))
    }
    catch {
        return null
    }
}

/**  the frame of an image element in EMU, null without explicit size  */
const imageBox = (el: Extract<ElementSpec, { type: "image" }>): { x: number, y: number, w: number, h: number } | null =>
    el.frame.w === undefined || el.frame.h === undefined ? null
        : { x: inchToEmu(el.frame.x), y: inchToEmu(el.frame.y), w: inchToEmu(el.frame.w), h: inchToEmu(el.frame.h) }

/**
 *  Finish a generated shape after its move into the target slide: a
 *  picture with `fit: crop` gets its centered `a:srcRect` (the same fit
 *  as picture placeholders).
 *
 *  @param shape - the moved shape element (changed in place)
 *  @param el - the element payload it was generated from
 */
export const finishShape = (shape: Element, el: ElementSpec): void => {
    if (el.type !== "image" || el.fit !== "crop")
        return
    const box = imageBox(el)
    const crop = box === null ? null : fitImage(box, fileImageSize(el.path), "crop").crop
    const blipFill = firstElement(shape, "p:blipFill")
    if (crop === null || blipFill === null)
        return
    const srcRect = (shape.ownerDocument as Document).createElementNS(NS_A, "a:srcRect")
    for (const [ side, value ] of Object.entries(crop))
        if (value > 0)
            srcRect.setAttribute(side, String(value))
    for (const old of elements(blipFill, "a:srcRect"))
        blipFill.removeChild(old)
    blipFill.insertBefore(srcRect, firstElement(blipFill, "a:stretch") ?? firstElement(blipFill, "a:tile"))
}

/**
 *  Render one validated element onto a PptxGenJS slide.
 *
 *  @param slide - PptxGenJS generator slide
 *  @param gen - PptxGenJS root object (enums)
 *  @param el - validated element payload
 *  @param lang - the language of text that states none (the deck language);
 *                always passed, as PptxGenJS writes `en-US` otherwise
 *  @param colors - the color tokens of the apply (theme colors)
 */
export const addElement = (slide: GenSlide, gen: GenRoot, el: ElementSpec, lang: string, colors: GenColors): void => {
    if (el.type === "connector") {
        const x = Math.min(el.from[0], el.to[0])
        const y = Math.min(el.from[1], el.to[1])
        const arrow = (a: string | undefined): string | null => a === undefined || a === "none" ? null : a
        slide.addShape(gen.ShapeType["line"], {
            x, y,
            w: Math.abs(el.to[0] - el.from[0]),
            h: Math.abs(el.to[1] - el.from[1]),
            flipH: el.to[0] < el.from[0],
            flipV: el.to[1] < el.from[1],
            line: {
                color: colors.gen(el.color ?? "4A4A4A"),
                width: el.widthPt ?? 1.5,
                ...(el.dash !== undefined && el.dash !== "solid" && { dashType: el.dash === "longDash" ? "lgDash" : el.dash }),
                /*  the flips keep the path start at `from`: head = start  */
                ...(arrow(el.startArrow) !== null && { beginArrowType: arrow(el.startArrow) }),
                ...(arrow(el.endArrow) !== null && { endArrowType: arrow(el.endArrow) })
            }
        })
        return
    }
    const frame = {
        x: el.frame.x, y: el.frame.y,
        ...(el.frame.w !== undefined && { w: el.frame.w }),
        ...(el.frame.h !== undefined && { h: el.frame.h }),
        ...(el.frame.rotation !== undefined && { rotate: el.frame.rotation })
    }
    switch (el.type) {
        case "textbox":
            slide.addText(genTextRuns(el.text, el.lang ?? lang, colors), {
                ...frame,
                lang: el.lang ?? lang,
                valign: "top",
                ...boxTextOptions(el),
                ...(el.fill !== undefined && { fill: { color: colors.gen(el.fill) } }),
                ...(el.border !== undefined && { line: { color: colors.gen(el.border), width: el.borderPt ?? 1 } })
            })
            break
        case "table": {
            const st = el.data.style
            const opts: Record<string, unknown> = {
                ...frame,
                border: { type: "solid", color: colors.gen(st?.border ?? "ACACAC"), pt: 0.5 },
                valign: "middle"
            }
            if (st?.fontSize !== undefined)
                opts.fontSize = st.fontSize
            if (el.textDirection !== undefined)
                opts.textDirection = TEXT_DIRECTION_VERT[el.textDirection]
            /*  column widths in inches, scaled to an explicit frame width  */
            if (el.colWidths !== undefined) {
                const sum = el.colWidths.reduce((a, b) => a + b, 0)
                const w = el.frame.w ?? sum
                opts.w = w
                opts.colW = el.colWidths.map((c) => c * w / sum)
            }
            /*  fixed header/row heights -> uniform tables across the deck  */
            if (st?.headerHeight !== undefined || st?.rowHeight !== undefined) {
                const hh = st.headerHeight ?? st.rowHeight
                const rh = st.rowHeight ?? st.headerHeight
                const heights: number[] = []
                if (el.data.headers !== undefined)
                    heights.push(hh as number)
                for (let i = 0; i < el.data.rows.length; i++)
                    heights.push(rh as number)
                opts.rowH = heights
            }
            slide.addTable(genTableRows(el, el.lang ?? lang, colors), opts)
            break
        }
        case "chart": {
            const { type, extra } = chartType(gen, el.data.type)
            const isXY = el.data.type === "scatter" || el.data.type === "bubble"
            /*  scatter/bubble: PptxGenJS takes the shared x values as the
                first series, the y values (and bubble sizes) after it  */
            const data = isXY
                ? [ { name: "X", values: el.data.series[0]?.x ?? [] }, ...el.data.series.map((s) => ({
                    name: s.name,
                    values: s.y ?? [],
                    ...(s.size !== undefined && { sizes: s.size })
                })) ]
                : el.data.series.map((s) => ({
                    name: s.name,
                    labels: (el.data.categories ?? []).map(String),
                    values: s.values ?? []
                }))
            slide.addChart(type, data, {
                ...frame, ...extra,
                chartColors: (el.data.colors ?? DEFAULT_COLORS).map((c) => colors.gen(c)),
                ...chartOptions(el.data),
                ...(el.data.title !== undefined && { showTitle: true, title: el.data.title }),
                ...(el.data.fontSize !== undefined && { catAxisLabelFontSize: el.data.fontSize, valAxisLabelFontSize: el.data.fontSize })
            })
            break
        }
        case "shape":
            slide.addText(el.text === undefined ? "" : genTextRuns(el.text, el.lang ?? lang, colors), {
                ...frame,
                lang: el.lang ?? lang,
                shape: gen.ShapeType[el.shape],
                align: "center",
                ...boxTextOptions(el),
                ...(el.fill !== undefined && { fill: { color: colors.gen(el.fill) } }),
                ...(el.border !== undefined && { line: { color: colors.gen(el.border), width: el.borderPt ?? 1 } }),
                ...(el.fontColor !== undefined && { color: colors.gen(el.fontColor) }),
                ...(el.fontSize !== undefined && { fontSize: el.fontSize })
            })
            break
        case "image": {
            /*  contain: the whole image in a centered, shrunk frame (EMU math
                as for picture placeholders); crop: finishShape  */
            const box = el.fit === "contain" ? imageBox(el) : null
            const fitted = box === null ? null : fitImage(box, fileImageSize(el.path), "contain").box
            slide.addImage({
                path: el.path,
                ...frame,
                ...(fitted !== null && { x: fitted.x / EMU_PER_INCH, y: fitted.y / EMU_PER_INCH,
                    w: fitted.w / EMU_PER_INCH, h: fitted.h / EMU_PER_INCH }),
                ...(el.hyperlink !== undefined && { hyperlink: { url: el.hyperlink } })
            })
            break
        }
    }
}
