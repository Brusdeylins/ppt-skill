/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Golden: the slide features of release 1.2 on the neutral fixture, a
**  fixed op sequence pinned by its own snapshot (the 1.0.11-frozen engine
**  goldens stay untouched). Beyond the generic fingerprint, each slide
**  records what the features write: picture crops and frames, the
**  language of every text run and -- on the "Phase 3" slides -- the XML
**  facts of the el.add options (arrows, table grid and cells, text
**  direction and layout, theme colors, image fit and link, chart labels
**  and axes), on the "Phase 4" slide those of the in-place edits (el.set
**  frame, style, alt text, image and table, text.replace, slide number
**  and date), on the "Phase 5" slides the slide links (action,
**  relationship, target); the generic fingerprint carries the sections.
*/

import { describe, expect, it } from "vitest"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { buildEmptyDeck } from "../../src/engine/seed.js"
import { DeckArchive, readDeckState } from "../../src/engine/reader.js"
import { executeOps, type ExecuteOptions } from "../../src/commands/apply.js"
import { PptcError } from "../../src/infra/errors.js"
import type { Element } from "@xmldom/xmldom"
import { relTargets, relsPartOf } from "../../src/engine/parts.js"
import { NS_R, elements, firstElement, parseXml } from "../../src/engine/xml.js"
import { expectIntact } from "../util/integrity.js"
import { fingerprint } from "../util/fingerprint.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const FIXTURE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp", "golden-1-2")

const opts: ExecuteOptions = { templatePath: null, dryRun: false, strict: false, expectRev: null, outFile: null, minFontPt: 8 }

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

/**  mask the UUID suffix of generated shape names  */
const maskName = (name: string): string =>
    name.replace(/-?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/, "-*")

/**  `name=value` for each present attribute of an element  */
const attrList = (el: Element | null, names: string[]): string =>
    el === null ? "-" : names.filter((n) => el.hasAttribute(n)).map((n) => `${n}=${el.getAttribute(n) ?? ""}`).join(",")

/**  the spacing value of a paragraph property child  */
const spacingVal = (pPr: Element, tag: string): string | null => {
    const el = elements(pPr, tag)[0]
    return el === undefined ? null : `${tag.slice(2)}=${(firstElement(el, "a:spcPct") ?? firstElement(el, "a:spcPts"))?.getAttribute("val") ?? ""}`
}

/**  the XML facts of the el.add options of 1.2 phase 3 in one shape  */
const shapeFacts = (shape: Element): string[] => {
    const out: string[] = []
    const scheme = elements(shape, "a:schemeClr").map((c) => c.getAttribute("val"))
    if (scheme.length > 0)
        out.push(`scheme:${scheme.join(",")}`)
    /*  outlines of shapes (a table's cell borders carry their own)  */
    for (const tag of firstElement(shape, "a:tbl") === null ? [ "a:headEnd", "a:tailEnd", "a:prstDash" ] : [])
        for (const el of elements(shape, tag))
            out.push(`${tag}:${el.getAttribute("type") ?? el.getAttribute("val") ?? ""}`)
    const bodyPr = elements(shape, "a:bodyPr").find((b) => b.parentNode?.nodeName === "p:txBody") ?? null
    if (bodyPr !== null)
        out.push(`bodyPr:${attrList(bodyPr, [ "vert", "anchor", "lIns", "tIns", "rIns", "bIns" ])}`)
    for (const p of elements(shape, "a:p").filter((el) => el.parentNode?.nodeName === "p:txBody")) {
        const pPrs = elements(p, "a:pPr")
        const pPr = pPrs[0] ?? null
        const para = pPr === null ? [] : [ attrList(pPr, [ "algn", "lvl" ]),
            ...[ "a:lnSpc", "a:spcBef", "a:spcAft" ].map((t) => spacingVal(pPr, t)).filter((v) => v !== null),
            ...[ "a:buAutoNum", "a:buChar", "a:buNone" ].filter((t) => firstElement(pPr, t) !== null) ]
        const runs = elements(p, "a:rPr").map((r) => attrList(r, [ "b", "i", "strike", "baseline" ])
            + (firstElement(r, "a:highlight") !== null ? "+highlight" : ""))
        out.push(`p[${pPrs.length}]:${para.join(" ")} | ${runs.join(" ; ")}`)
    }
    const grid = elements(shape, "a:gridCol").map((g) => g.getAttribute("w"))
    if (grid.length > 0)
        out.push(`grid:${grid.join(",")}`)
    for (const [ r, tr ] of elements(shape, "a:tr").entries())
        out.push(`tr${r}:${elements(tr, "a:tc").map((tc) => [ attrList(tc, [ "gridSpan", "rowSpan", "hMerge", "vMerge" ]),
            attrList(firstElement(tc, "a:tcPr"), [ "vert", "anchor" ]), firstElement(tc, "a:pPr")?.getAttribute("algn") ?? "",
            elements(tc, "a:t").map((t) => t.textContent).join("") ].join("/")).join(" | ")}`)
    if (firstElement(shape, "a:hlinkClick") !== null)
        out.push("hlinkClick")
    return out
}

/**  the XML facts of the in-place edits of 1.2 phase 4 in one shape:
     alt text, transform, fill and outline  */
const editFacts = (shape: Element): string[] => {
    const xfrm = firstElement(shape, "a:xfrm") ?? firstElement(shape, "p:xfrm")
    const spPr = firstElement(shape, "p:spPr")
    const direct = (parent: Element | null, tag: string): Element | null =>
        parent === null ? null : elements(parent, tag).find((el) => el.parentNode === parent) ?? null
    const color = (el: Element | null): string => el === null ? "-"
        : (firstElement(el, "a:srgbClr") ?? firstElement(el, "a:schemeClr"))?.getAttribute("val") ?? "?"
    const ln = direct(spPr, "a:ln")
    return [
        `cNvPr:${attrList(firstElement(shape, "p:cNvPr"), [ "descr", "title" ])}`,
        `xfrm:${attrList(xfrm, [ "rot" ])}|${attrList(firstElement(xfrm ?? shape, "a:off"), [ "x", "y" ])}|${attrList(firstElement(xfrm ?? shape, "a:ext"), [ "cx", "cy" ])}`,
        `fill:${color(direct(spPr, "a:solidFill"))} ln:${color(direct(ln, "a:solidFill"))}/${ln?.getAttribute("w") ?? "-"}`
    ]
}

/**  the facts of a chart part: grouping, labels, legend, axes, colors  */
const chartFacts = (chart: Element): string[] => {
    const vals = (tag: string): string => [ ...new Set(elements(chart, tag).map((el) => el.getAttribute("val"))) ].join(",")
    return [
        `grouping=${vals("c:grouping")} dLblPos=${vals("c:dLblPos")} showVal=${vals("c:showVal")} showPercent=${vals("c:showPercent")}`,
        `legendPos=${vals("c:legendPos")} holeSize=${vals("c:holeSize")} smooth=${vals("c:smooth")} min=${vals("c:min")} max=${vals("c:max")}`,
        `numFmt=${[ ...new Set(elements(chart, "c:numFmt").map((el) => el.getAttribute("formatCode"))) ].join(",")}`,
        `axisTitles=${[ ...elements(chart, "c:catAx"), ...elements(chart, "c:valAx") ]
            .map((ax) => elements(firstElement(ax, "c:title") ?? ax, "a:t").map((t) => t.textContent).join("")).join(",")}`,
        `scheme=${elements(chart, "a:schemeClr").map((c) => c.getAttribute("val")).join(",")}`
    ]
}

/**  what the 1.2 features write on each slide: picture crops and frame,
     the runs of every shape with their language, and on "Phase 3" slides
     the facts of every shape, chart and the background  */
const featurePrint = async (deck: string): Promise<unknown[]> => {
    const state = await readDeckState(await DeckArchive.open(deck))
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const out: unknown[] = []
    for (const slide of state.slides) {
        const doc = parseXml(await (zip.file(slide.part) as JSZip.JSZipObject).async("string"))
        const pictures = elements(doc, "p:pic").map((pic) => {
            const src = firstElement(pic, "a:srcRect")
            const off = firstElement(pic, "a:off")
            const ext = firstElement(pic, "a:ext")
            return [ maskName(firstElement(pic, "p:cNvPr")?.getAttribute("name") ?? ""),
                src === null ? "no-crop" : [ "l", "t", "r", "b" ].map((a) => `${a}=${src.getAttribute(a) ?? ""}`).join(","),
                [ off?.getAttribute("x"), off?.getAttribute("y"), ext?.getAttribute("cx"), ext?.getAttribute("cy") ].join(",") ].join("|")
        })
        const runs = [ ...elements(doc, "p:sp"), ...elements(doc, "p:graphicFrame") ].flatMap((shape) => {
            const name = maskName(firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? "")
            return elements(shape, "a:r").map((r) =>
                `${name}: ${firstElement(r, "a:t")?.textContent ?? ""}=${firstElement(r, "a:rPr")?.getAttribute("lang") ?? "-"}`)
        })
        if (slide.title?.startsWith("Phase 5") === true) {
            const rels = relTargets(slide.part, parseXml(await (zip.file(relsPartOf(slide.part)) as JSZip.JSZipObject).async("string")))
            const titleOf = new Map(state.slides.map((s) => [ s.part, s.title ]))
            const links = elements(doc, "a:hlinkClick").map((h) => {
                let shape = h.parentNode as Element | null
                while (shape !== null && ![ "p:sp", "p:pic", "p:graphicFrame" ].includes(shape.nodeName))
                    shape = shape.parentNode as Element | null
                const rel = rels.find((r) => r.id === h.getAttributeNS(NS_R, "id"))
                return [ maskName(shape === null ? "" : firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? ""),
                    h.getAttribute("action") ?? "-", rel?.kind ?? "?",
                    rel?.target === null ? rel.el.getAttribute("Target") : titleOf.get(rel?.target ?? "") ?? rel?.target ].join("|")
            })
            out.push({ pictures, runs, links })
            continue
        }
        if (slide.title !== "Phase 3" && slide.title !== "Phase 4") {
            out.push({ pictures, runs })
            continue
        }
        const rels = relTargets(slide.part, parseXml(await (zip.file(relsPartOf(slide.part)) as JSZip.JSZipObject).async("string")))
        const facts: Record<string, string[]> = { background: elements(firstElement(doc, "p:bg") ?? doc.createElement("none"), "a:schemeClr")
            .map((c) => c.getAttribute("val") ?? "") }
        for (const shape of [ ...elements(doc, "p:sp"), ...elements(doc, "p:graphicFrame"), ...elements(doc, "p:pic") ]) {
            const name = maskName(firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? "")
            const rid = firstElement(shape, "c:chart")?.getAttributeNS(NS_R, "id")
            const target = rels.find((r) => r.id === rid)?.target ?? null
            facts[name] = target === null ? shapeFacts(shape)
                : chartFacts(parseXml(await (zip.file(target) as JSZip.JSZipObject).async("string")).documentElement as Element)
            if (slide.title === "Phase 4")
                facts[name] = [ ...facts[name] ?? [], ...editFacts(shape) ]
        }
        out.push({ pictures, runs, facts })
    }
    return out
}

/**  one step: ops and whether the apply must fail (with its code)  */
interface Step { ops: unknown[], fails?: string }

describe("golden 1.2 features", () => {
    it("matches on an empty deck from the fixture", async () => {
        mkdirSync(TMP, { recursive: true })
        const landscape = path.join(TMP, "landscape.png")
        const portrait = path.join(TMP, "portrait.png")
        writeFileSync(landscape, png(1600, 900))
        writeFileSync(portrait, png(900, 1600))
        const deck = path.join(TMP, "features.pptx")
        writeFileSync(deck, await buildEmptyDeck(FIXTURE))

        const steps: Step[] = [
            /*  A2: text language from pptcDeckLang, with overrides  */
            { ops: [
                { op: "meta.props", set: { custom: { pptcDeckLang: "de-DE" } } },
                { op: "slide.add", layout: "CONTENT", ref: "text", notes: "Notiz", footer: "Fußzeile",
                    placeholders: { title: { text: "Sprache" },
                        body: { text: [ { text: "deutsch" }, { text: "english", lang: "en-GB" },
                            { runs: [ { text: "uno" }, { text: "deux", lang: "fr-FR" } ], lang: "it-IT" } ] } } },
                { op: "el.add", slide: "$text", elements: [
                    { type: "textbox", name: "Box", frame: { x: 0.5, y: 5.5, w: 3, h: 0.6 }, text: "Kasten" },
                    { type: "shape", name: "Form", shape: "rect", frame: { x: 4, y: 5.5, w: 2, h: 0.6 }, text: "Form", lang: "nl-NL" },
                    { type: "table", name: "Tab", frame: { x: 6.5, y: 5.5, w: 4 }, data: { headers: [ "Kopf" ], rows: [ [ "Zelle" ] ] } }
                ] }
            ] },
            /*  B3: picture placeholder fits  */
            { ops: [
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: landscape } } },
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: portrait, fit: "crop" } } },
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: portrait, fit: "contain" } } },
                { op: "slide.add", layout: "PICTURE", placeholders: { image: { image: landscape, fit: "stretch" } } }
            ] },
            /*  B1: el.set text on a table fails, on a textbox it works  */
            { ops: [ { op: "el.set", slide: "0", name: "Tab", text: "x" } ], fails: "E_SCHEMA" },
            { ops: [ { op: "el.set", slide: "0", name: "Box", text: "Neu" } ] },
            /*  phase 3: el.add options (A1, A3, U4, A4, A7, A5, A6, A8)  */
            { ops: [
                { op: "slide.add", layout: "Title Only", ref: "p3", background: { color: "lt2" },
                    placeholders: { title: { text: "Phase 3" } } },
                { op: "el.add", slide: "$p3", elements: [
                    { type: "connector", name: "F-Arrow", from: [ 4, 2 ], to: [ 0.5, 1.8 ], color: "accent2",
                        startArrow: "diamond", endArrow: "stealth", dash: "longDash" },
                    { type: "textbox", name: "F-Text", frame: { x: 0.5, y: 2, w: 3, h: 2 }, valign: "middle",
                        margin: [ 0.1, 0.05, 0.1, 0.05 ], fill: "bg2", text: [
                            { runs: [ { text: "H", color: "tx1" }, { text: "2", subscript: true }, { text: "O", strike: true, highlight: "accent4" },
                                { text: "x", superscript: true } ], bullet: "number", spaceBefore: 6, spaceAfter: 2, lineSpacing: 1.2 },
                            { text: "Punkt", bullet: true } ] },
                    { type: "textbox", name: "F-Vertical", frame: { x: 3.7, y: 2, w: 0.6, h: 2 }, text: "vertikal", textDirection: "rotate270" },
                    { type: "shape", name: "F-Shape", shape: "roundRect", frame: { x: 4.5, y: 2, w: 2, h: 1 }, fill: "accent1",
                        fontColor: "lt1", border: "dk2", valign: "bottom", text: [ { runs: [ { text: "fett", bold: true }, { text: " normal" } ] } ] },
                    { type: "table", name: "F-Table", frame: { x: 0.5, y: 4.3, w: 6 }, colWidths: [ 1, 2, 3 ],
                        colAlign: [ "left", "center", "right" ], textDirection: "horizontal",
                        data: { headers: [ { text: "Zwei", colspan: 2 }, "Drei" ], style: { headerBg: "accent1", headerFg: "bg1" },
                            rows: [ [ { text: "hoch", rowspan: 2, textDirection: "rotate90", fill: "accent6" }, "b", 1 ],
                                [ { text: [ { runs: [ { text: "reich", italic: true } ] } ], align: "left", valign: "bottom", color: "dk2" }, 2 ] ] } },
                    { type: "image", name: "F-Crop", frame: { x: 7, y: 1.7, w: 2.5, h: 1.2 }, path: landscape, fit: "crop",
                        hyperlink: "https://example.com" },
                    { type: "image", name: "F-Contain", frame: { x: 7, y: 3, w: 2.5, h: 1.2 }, path: portrait, fit: "contain" },
                    { type: "image", name: "F-Stretch", frame: { x: 7, y: 4.3, w: 2.5, h: 1.2 }, path: landscape, fit: "stretch" }
                ] }
            ] },
            { ops: [
                { op: "slide.add", layout: "Title Only", ref: "p3c", placeholders: { title: { text: "Phase 3" } } },
                { op: "el.add", slide: "$p3c", elements: [
                    { type: "chart", name: "F-Percent", frame: { x: 0.5, y: 1.7, w: 4.5, h: 2.5 }, data: { type: "column",
                        categories: [ "Q1", "Q2" ], series: [ { name: "a", values: [ 1, 2 ] }, { name: "b", values: [ 3, 1 ] } ],
                        stacked: "percent", labels: { position: "center", format: "0.0" }, legend: "right", colors: [ "accent1", "dk2" ],
                        axes: { x: { title: "Quartal" }, y: { title: "Anteil" } } } },
                    { type: "chart", name: "F-Doughnut", frame: { x: 5.2, y: 1.7, w: 4.3, h: 2.5 }, data: { type: "doughnut",
                        categories: [ "a", "b" ], series: [ { name: "s", values: [ 1, 3 ] } ], holeSize: 70, labels: { percent: true }, legend: "left" } },
                    { type: "chart", name: "F-Scatter", frame: { x: 0.5, y: 4.4, w: 4.5, h: 2.4 }, data: { type: "scatter",
                        series: [ { name: "s", x: [ 1, 2, 3 ], y: [ 2, 1, 3 ] }, { name: "t", x: [ 1, 2, 3 ], y: [ 1, 3, 2 ] } ], labels: { position: "above" },
                        axes: { x: { min: 0, max: 4, format: "0.0" }, y: { min: 0, max: 4 } } } },
                    { type: "chart", name: "F-Bars", frame: { x: 5.2, y: 4.4, w: 4.3, h: 2.4 }, data: { type: "bar",
                        categories: [ "a", "b" ], series: [ { name: "s", values: [ 1, 3 ] } ], labels: { position: "outsideEnd" }, legend: false } }
                ] }
            ] },
            /*  phase 3: text layout through the DOM path (slide.fill, el.set)  */
            { ops: [
                { op: "slide.add", layout: "CONTENT", ref: "p3t", placeholders: { title: { text: "Phase 3" },
                    body: { text: [ { runs: [ { text: "x" }, { text: "2", superscript: true }, { text: "alt", strike: true, highlight: "FFFF00" } ],
                        bullet: "number", spaceBefore: 6, lineSpacing: 1.5 } ] } } }
            ] },
            { ops: [ { op: "el.set", slide: "7", name: "PptcPh-1", text: [ { text: "neu", subscript: true, color: "accent2", spaceAfter: 12 },
                { text: "zwei", bullet: "number" } ] } ] },
            /*  phase 4: in-place edits (B2, B6, B4, B5, B7, C1)  */
            { ops: [
                { op: "slide.add", layout: "Title Only", ref: "p4", slideNumber: false, date: false, placeholders: { title: { text: "Phase 4" } } },
                { op: "el.add", slide: "$p4", elements: [
                    { type: "textbox", name: "G-Box", frame: { x: 0.5, y: 1.7, w: 3, h: 1 }, altText: "Kasten", title: "Box",
                        text: [ { runs: [ { text: "alpha beta", bold: true } ], alignment: "center" } ] },
                    { type: "image", name: "G-Pic", frame: { x: 4, y: 1.7, w: 3, h: 2 }, path: landscape, fit: "crop" },
                    { type: "table", name: "G-Tab", frame: { x: 0.5, y: 4.2, w: 6 },
                        data: { headers: [ "K", "V" ], rows: [ [ "alpha", "1" ], [ "beta", { text: "2", fill: "accent6" } ] ] } },
                    { type: "shape", name: "G-Shape", shape: "rect", frame: { x: 7.5, y: 1.7, w: 2, h: 1 }, text: "alpha" } ] }
            ] },
            { ops: [
                { op: "el.set", slide: "title:Phase 4", name: "G-Box", frame: { x: 1, rotation: 15 }, fill: "accent3", border: "dk2",
                    borderPt: 1.5, fontColor: "C00000", text: "gamma alpha" },
                { op: "el.set", slide: "title:Phase 4", name: "G-Pic", image: portrait, altText: "Hochformat" },
                { op: "el.set", slide: "title:Phase 4", name: "G-Tab", cells: [ { row: 1, col: 1, text: "10" } ],
                    addRows: { rows: [ [ "gamma", 3 ] ] }, removeRows: [ 2 ] },
                { op: "text.replace", find: "alpha", replace: "omega", slide: "title:Phase 4" },
                { op: "slide.fill", slide: "title:Phase 4", slideNumber: true, date: true }
            ] },
            /*  phase 5: sections and slide links (C2, C3) -- the agenda links
                forward to a slide added later in the document  */
            { ops: [
                { op: "slide.add", layout: "CONTENT", ref: "p5", section: "Phase 5", placeholders: { title: { text: "Phase 5" },
                    body: { text: [ { runs: [ { text: "zu Phase 4", hyperlink: "slide:title:Phase 4" } ] },
                        { runs: [ { text: "zum Ende", hyperlink: "slide:$p5e" } ] },
                        { runs: [ { text: "Web", hyperlink: "https://example.com/" } ] } ] } } },
                { op: "el.add", slide: "$p5", elements: [
                    { type: "shape", name: "H-Back", shape: "roundRect", frame: { x: 8, y: 6, w: 1.5, h: 0.5 }, text: "Start",
                        hyperlink: "slide:index:0" },
                    { type: "image", name: "H-Pic", frame: { x: 6, y: 6, w: 1, h: 0.6 }, path: landscape, hyperlink: "slide:$p5e" },
                    { type: "table", name: "H-Tab", frame: { x: 0.5, y: 6, w: 4 },
                        data: { rows: [ [ { text: [ { runs: [ { text: "Phase 4", hyperlink: "slide:title:Phase 4" } ] } ] }, "x" ] ] } } ] },
                { op: "slide.add", layout: "Title Only", ref: "p5e", section: "Ende", placeholders: { title: { text: "Phase 5 Ende" } } },
                { op: "el.add", slide: "$p5e", elements: [ { type: "textbox", name: "H-Box", frame: { x: 8, y: 6, w: 1.5, h: 0.5 },
                    text: "zurück", hyperlink: "slide:$p5" } ] }
            ] },
            /*  split, rename and remove sections; a slide link set in place  */
            { ops: [
                { op: "slide.fill", slide: "title:Phase 4", section: "Phase 4" },
                { op: "section.rename", from: "Default Section", to: "Phasen 1-3" },
                { op: "section.rm", name: "Ende" },
                { op: "el.set", slide: "title:Phase 5 Ende", name: "PptcPh-0", hyperlink: "slide:title:Sprache" }
            ] },
            { ops: [ { op: "slide.fill", slide: "title:Phase 5 Ende", section: "Phase 4" } ], fails: "E_SCHEMA" },
            { ops: [ { op: "el.set", slide: "title:Phase 5", name: "H-Back", hyperlink: "slide:title:Nirgends" } ], fails: "E_ADDR_NOTFOUND" }
        ]

        const prints: unknown[] = []
        for (const step of steps) {
            let error: string | null = null
            try {
                await executeOps(deck, { ops: step.ops } as never, opts)
            }
            catch (err) {
                error = err instanceof PptcError ? err.code : String(err)
            }
            expect(error).toBe(step.fails ?? null)
            await expectIntact(deck)
            prints.push({ ops: JSON.parse(JSON.stringify(step.ops).split(TMP).join("<tmp>")) as unknown, error,
                slides: await fingerprint(deck), features: await featurePrint(deck) })
        }
        expect(prints).toMatchSnapshot()
    }, 120000)
})
