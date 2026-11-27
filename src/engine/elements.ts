/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/elements: translate validated ElementSpec payloads into PptxGenJS
**  calls on a generator slide; engine/inplace/transplant moves the result
**  into the target slide.
*/

import type { ChartData, ElementSpec, RichText, TableData } from "../schema/payloads.js"

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

/**  convert rich text to the PptxGenJS text-run array form  */
const genTextRuns = (text: RichText): unknown => {
    if (typeof text === "string")
        return text
    const out: Record<string, unknown>[] = []
    for (const p of text) {
        const runs = p.runs ?? [{ text: p.text ?? "", bold: p.bold, italic: p.italic,
            underline: p.underline, size: p.size, font: p.font, color: p.color }]
        runs.forEach((r, i) => {
            out.push({
                text: r.text,
                options: {
                    breakLine: i === runs.length - 1,
                    ...(r.bold !== undefined && { bold: r.bold }),
                    ...(r.italic !== undefined && { italic: r.italic }),
                    ...(r.underline !== undefined && { underline: { style: "sng" } }),
                    ...(r.size !== undefined && { fontSize: r.size }),
                    ...(r.font !== undefined && { fontFace: r.font }),
                    ...(r.color !== undefined && { color: r.color }),
                    ...("hyperlink" in r && r.hyperlink !== undefined && { hyperlink: { url: r.hyperlink } }),
                    ...(p.level !== undefined && { indentLevel: p.level }),
                    ...(p.alignment !== undefined && { align: p.alignment }),
                    ...(p.bullet !== undefined && { bullet: p.bullet })
                }
            })
        })
    }
    return out
}

/**  build PptxGenJS table rows including header styling  */
const genTableRows = (data: TableData): unknown[][] => {
    const rows: unknown[][] = []
    if (data.headers !== undefined)
        rows.push(data.headers.map((h) => ({
            text: h,
            options: {
                bold: true,
                color: data.style?.headerFg ?? "FFFFFF",
                fill: { color: data.style?.headerBg ?? "333333" }
            }
        })))
    data.rows.forEach((row, i) => {
        rows.push(row.map((cell) => ({
            text: String(cell),
            options: data.style?.altRowBg !== undefined && i % 2 === 1
                ? { fill: { color: data.style.altRowBg } }
                : {}
        })))
    })
    return rows
}

/**  default series color cycle when the payload specifies none  */
const DEFAULT_COLORS = ["4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646"]

/**
 *  Render one validated element onto a PptxGenJS slide.
 *
 *  @param slide - PptxGenJS generator slide
 *  @param gen - PptxGenJS root object (enums)
 *  @param el - validated element payload
 */
export const addElement = (slide: GenSlide, gen: GenRoot, el: ElementSpec): void => {
    if (el.type === "connector") {
        const x = Math.min(el.from[0], el.to[0])
        const y = Math.min(el.from[1], el.to[1])
        slide.addShape(gen.ShapeType["line"], {
            x, y,
            w: Math.abs(el.to[0] - el.from[0]),
            h: Math.abs(el.to[1] - el.from[1]),
            flipH: el.to[0] < el.from[0],
            flipV: el.to[1] < el.from[1],
            line: {
                color: el.color ?? "4A4A4A",
                width: el.widthPt ?? 1.5,
                ...(el.dash !== undefined && el.dash !== "solid" && { dashType: el.dash })
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
            slide.addText(genTextRuns(el.text), {
                ...frame,
                valign: "top",
                ...(el.fill !== undefined && { fill: { color: el.fill } }),
                ...(el.border !== undefined && { line: { color: el.border, width: el.borderPt ?? 1 } })
            })
            break
        case "table": {
            const st = el.data.style
            const opts: Record<string, unknown> = {
                ...frame,
                border: { type: "solid", color: st?.border ?? "ACACAC", pt: 0.5 },
                valign: "middle"
            }
            if (st?.fontSize !== undefined)
                opts.fontSize = st.fontSize
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
            slide.addTable(genTableRows(el.data), opts)
            break
        }
        case "chart": {
            const { type, extra } = chartType(gen, el.data.type)
            const isXY = el.data.type === "scatter" || el.data.type === "bubble"
            const data = isXY
                ? el.data.series.map((s) => ({
                    name: s.name,
                    values: s.y ?? [],
                    labels: (s.x ?? []).map(String),
                    ...(s.size !== undefined && { sizes: s.size })
                }))
                : el.data.series.map((s) => ({
                    name: s.name,
                    labels: (el.data.categories ?? []).map(String),
                    values: s.values ?? []
                }))
            slide.addChart(type, data, {
                ...frame, ...extra,
                chartColors: el.data.colors ?? DEFAULT_COLORS,
                showLegend: el.data.legend ?? !["pie", "doughnut"].includes(el.data.type),
                legendPos: "b",
                ...(el.data.title !== undefined && { showTitle: true, title: el.data.title }),
                ...(el.data.fontSize !== undefined && { catAxisLabelFontSize: el.data.fontSize, valAxisLabelFontSize: el.data.fontSize })
            })
            break
        }
        case "shape":
            slide.addText(el.text ?? "", {
                ...frame,
                shape: gen.ShapeType[el.shape],
                align: "center",
                ...(el.fill !== undefined && { fill: { color: el.fill } }),
                ...(el.border !== undefined && { line: { color: el.border, width: el.borderPt ?? 1 } }),
                ...(el.fontColor !== undefined && { color: el.fontColor }),
                ...(el.fontSize !== undefined && { fontSize: el.fontSize })
            })
            break
        case "image":
            slide.addImage({ path: el.path, ...frame })
            break
    }
}
