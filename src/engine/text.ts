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
import { NS_A, elements, firstElement } from "./xml.js"
import type { Paragraph, RichText, Run } from "../schema/payloads.js"

/**  marker attribute the engine replaces by a hyperlink relationship  */
export const HLINK_ATTR = "pptc-hlink"

/**  normalize plain-string content into the paragraph form  */
const toParagraphs = (text: RichText): Paragraph[] =>
    typeof text === "string"
        ? text.split("\n").map((line) => ({ text: line }))
        : text

/**  build the `a:rPr` run properties element when any formatting is set  */
const buildRunProps = (doc: Document, run: Partial<Run>): Element | null => {
    const hasProps = run.bold !== undefined || run.italic !== undefined
        || run.underline !== undefined || run.size !== undefined
        || run.font !== undefined || run.color !== undefined
        || run.hyperlink !== undefined
    if (!hasProps)
        return null
    const rPr = doc.createElementNS(NS_A, "a:rPr")
    if (run.bold !== undefined)
        rPr.setAttribute("b", run.bold ? "1" : "0")
    if (run.italic !== undefined)
        rPr.setAttribute("i", run.italic ? "1" : "0")
    if (run.underline === true)
        rPr.setAttribute("u", "sng")
    if (run.size !== undefined)
        rPr.setAttribute("sz", String(Math.round(run.size * 100)))
    if (run.color !== undefined) {
        const fill = doc.createElementNS(NS_A, "a:solidFill")
        const clr = doc.createElementNS(NS_A, "a:srgbClr")
        clr.setAttribute("val", run.color)
        fill.appendChild(clr)
        rPr.appendChild(fill)
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
const buildRun = (doc: Document, run: Run): Element => {
    const r = doc.createElementNS(NS_A, "a:r")
    const rPr = buildRunProps(doc, run)
    if (rPr !== null)
        r.appendChild(rPr)
    const t = doc.createElementNS(NS_A, "a:t")
    t.appendChild(doc.createTextNode(run.text))
    r.appendChild(t)
    return r
}

/**  build one `a:p` paragraph including paragraph properties  */
const buildParagraph = (doc: Document, p: Paragraph): Element => {
    const el = doc.createElementNS(NS_A, "a:p")
    const needsPPr = p.level !== undefined || p.alignment !== undefined || p.bullet === false
    if (needsPPr) {
        const pPr = doc.createElementNS(NS_A, "a:pPr")
        if (p.level !== undefined)
            pPr.setAttribute("lvl", String(p.level))
        if (p.alignment !== undefined)
            pPr.setAttribute("algn", { left: "l", center: "ctr", right: "r", justify: "just" }[p.alignment])
        if (p.bullet === false)
            pPr.appendChild(doc.createElementNS(NS_A, "a:buNone"))
        el.appendChild(pPr)
    }
    const runs: Run[] = p.runs ?? [{
        text: p.text ?? "",
        ...(p.bold !== undefined && { bold: p.bold }),
        ...(p.italic !== undefined && { italic: p.italic }),
        ...(p.underline !== undefined && { underline: p.underline }),
        ...(p.size !== undefined && { size: p.size }),
        ...(p.font !== undefined && { font: p.font }),
        ...(p.color !== undefined && { color: p.color })
    }]
    for (const run of runs)
        el.appendChild(buildRun(doc, run))
    return el
}

/**
 *  Replace (or append to) the text content of a shape's txBody.
 *
 *  @param shape - the `p:sp` element to write into
 *  @param text - validated rich text payload
 *  @param append - keep existing paragraphs and append instead of replacing
 */
export const setShapeText = (shape: Element, text: RichText, append = false): void => {
    const doc = shape.ownerDocument as Document
    const txBody = firstElement(shape, "p:txBody")
    if (txBody === null)
        return
    if (!append)
        for (const p of elements(txBody, "a:p"))
            txBody.removeChild(p)
    for (const paragraph of toParagraphs(text))
        txBody.appendChild(buildParagraph(doc, paragraph))
}
