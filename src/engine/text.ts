/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/text: build DrawingML rich text (`a:p`/`a:r`/`a:t`) from the
**  validated RichText payload, directly into a shape's txBody via DOM.
**  This is the write-side counterpart of reader's drawingText().
*/

import type { Document, Element } from "@xmldom/xmldom"
import { NS_A, colorElement, elements, firstElement } from "./xml.js"
import type { Paragraph, RichText, Run } from "../schema/payloads.js"

/**  marker attribute the engine replaces by a hyperlink relationship  */
export const HLINK_ATTR = "pptc-hlink"

/**  normalize plain-string content into the paragraph form  */
const toParagraphs = (text: RichText): Paragraph[] =>
    typeof text === "string"
        ? text.split("\n").map((line) => ({ text: line }))
        : text

/**  text baseline shift of superscript and subscript runs, in 1/1000 %
     (PowerPoint's own values)  */
export const BASELINE = { superscript: 30000, subscript: -25000 }

/**  build the `a:rPr` run properties element (always: it carries the
     run's language); children in schema order: fill, highlight, latin,
     hlinkClick  */
const buildRunProps = (doc: Document, run: Partial<Run>, lang: string): Element => {
    const rPr = doc.createElementNS(NS_A, "a:rPr")
    rPr.setAttribute("lang", run.lang ?? lang)
    if (run.bold !== undefined)
        rPr.setAttribute("b", run.bold ? "1" : "0")
    if (run.italic !== undefined)
        rPr.setAttribute("i", run.italic ? "1" : "0")
    if (run.underline === true)
        rPr.setAttribute("u", "sng")
    if (run.strike !== undefined)
        rPr.setAttribute("strike", run.strike ? "sngStrike" : "noStrike")
    if (run.superscript === true)
        rPr.setAttribute("baseline", String(BASELINE.superscript))
    else if (run.subscript === true)
        rPr.setAttribute("baseline", String(BASELINE.subscript))
    if (run.size !== undefined)
        rPr.setAttribute("sz", String(Math.round(run.size * 100)))
    if (run.color !== undefined) {
        const fill = doc.createElementNS(NS_A, "a:solidFill")
        fill.appendChild(colorElement(doc, run.color))
        rPr.appendChild(fill)
    }
    if (run.highlight !== undefined) {
        const highlight = doc.createElementNS(NS_A, "a:highlight")
        highlight.appendChild(colorElement(doc, run.highlight))
        rPr.appendChild(highlight)
    }
    if (run.font !== undefined) {
        const latin = doc.createElementNS(NS_A, "a:latin")
        latin.setAttribute("typeface", run.font)
        rPr.appendChild(latin)
    }
    if (run.hyperlink !== undefined) {
        /*  the relationship id is assigned by the engine; mark with the URL  */
        const hlink = doc.createElementNS(NS_A, "a:hlinkClick")
        hlink.setAttribute(HLINK_ATTR, run.hyperlink)
        rPr.appendChild(hlink)
    }
    return rPr
}

/**  build one `a:r` text run  */
const buildRun = (doc: Document, run: Run, lang: string): Element => {
    const r = doc.createElementNS(NS_A, "a:r")
    r.appendChild(buildRunProps(doc, run, lang))
    const t = doc.createElementNS(NS_A, "a:t")
    t.appendChild(doc.createTextNode(run.text))
    r.appendChild(t)
    return r
}

/**  a spacing element (`a:lnSpc`, `a:spcBef`, `a:spcAft`) with its value  */
const spacing = (doc: Document, tag: string, unit: "a:spcPct" | "a:spcPts", val: number): Element => {
    const el = doc.createElementNS(NS_A, tag)
    const v = doc.createElementNS(NS_A, unit)
    v.setAttribute("val", String(Math.round(val)))
    el.appendChild(v)
    return el
}

/**  the run of a paragraph without runs: its text with the paragraph's
     character formatting  */
export const paragraphRun = (p: Paragraph): Run => ({
    text: p.text ?? "",
    ...(p.bold !== undefined && { bold: p.bold }),
    ...(p.italic !== undefined && { italic: p.italic }),
    ...(p.underline !== undefined && { underline: p.underline }),
    ...(p.strike !== undefined && { strike: p.strike }),
    ...(p.superscript !== undefined && { superscript: p.superscript }),
    ...(p.subscript !== undefined && { subscript: p.subscript }),
    ...(p.size !== undefined && { size: p.size }),
    ...(p.font !== undefined && { font: p.font }),
    ...(p.color !== undefined && { color: p.color }),
    ...(p.highlight !== undefined && { highlight: p.highlight })
})

/**  build one `a:p` paragraph including paragraph properties (children in
     schema order: line spacing, space before/after, bullet); its runs
     take the paragraph's language, else the given one. `bullet: true`
     keeps the placeholder's own bullet style.  */
const buildParagraph = (doc: Document, p: Paragraph, lang: string): Element => {
    const el = doc.createElementNS(NS_A, "a:p")
    const needsPPr = p.level !== undefined || p.alignment !== undefined || p.bullet === false
        || p.bullet === "number" || p.lineSpacing !== undefined || p.spaceBefore !== undefined
        || p.spaceAfter !== undefined
    if (needsPPr) {
        const pPr = doc.createElementNS(NS_A, "a:pPr")
        if (p.level !== undefined)
            pPr.setAttribute("lvl", String(p.level))
        if (p.alignment !== undefined)
            pPr.setAttribute("algn", { left: "l", center: "ctr", right: "r", justify: "just" }[p.alignment])
        if (p.lineSpacing !== undefined)
            pPr.appendChild(spacing(doc, "a:lnSpc", "a:spcPct", p.lineSpacing * 100000))
        if (p.spaceBefore !== undefined)
            pPr.appendChild(spacing(doc, "a:spcBef", "a:spcPts", p.spaceBefore * 100))
        if (p.spaceAfter !== undefined)
            pPr.appendChild(spacing(doc, "a:spcAft", "a:spcPts", p.spaceAfter * 100))
        if (p.bullet === false)
            pPr.appendChild(doc.createElementNS(NS_A, "a:buNone"))
        else if (p.bullet === "number") {
            const num = doc.createElementNS(NS_A, "a:buAutoNum")
            num.setAttribute("type", "arabicPeriod")
            pPr.appendChild(num)
        }
        el.appendChild(pPr)
    }
    const runs: Run[] = p.runs ?? [ paragraphRun(p) ]
    for (const run of runs)
        el.appendChild(buildRun(doc, run, p.lang ?? lang))
    return el
}

/**
 *  Replace (or append to) the text content of a shape's txBody.
 *
 *  @param shape - the `p:sp` element to write into
 *  @param text - validated rich text payload
 *  @param append - keep existing paragraphs and append instead of replacing
 *  @param lang - the language of runs that state none (the deck language)
 */
export const setShapeText = (shape: Element, text: RichText, append: boolean, lang: string): void => {
    const doc = shape.ownerDocument as Document
    const txBody = firstElement(shape, "p:txBody")
    if (txBody === null)
        return
    if (!append)
        for (const p of elements(txBody, "a:p"))
            txBody.removeChild(p)
    for (const paragraph of toParagraphs(text))
        txBody.appendChild(buildParagraph(doc, paragraph, lang))
}
