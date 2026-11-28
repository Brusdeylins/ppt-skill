/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/text: build DrawingML rich text (`a:p`/`a:r`/`a:t`) from the
**  validated RichText payload, directly into a shape's txBody via DOM.
**  This is the write-side counterpart of reader's drawingText(). An
**  existing body may keep its first paragraph's and run's format for the
**  new text that sets none (el.set); `replaceInBody` is text.replace on a
**  body's runs.
*/

import type { Document, Element } from "@xmldom/xmldom"
import { NS_A, colorElement, firstElement } from "./xml.js"
import type { Paragraph, RichText, Run } from "../schema/payloads.js"
import { replaceRuns, type ReplaceSpec } from "../core/text-replace.js"

/**  the format a body keeps for new text that sets none: its first
     paragraph's `a:pPr` and its first run's `a:rPr` (else the
     paragraph-end properties)  */
interface KeptFormat {
    pPr: Element | null
    rPr: Element | null
}

/**  the direct children of an element with one of the given names  */
const children = (parent: Element, names: string[]): Element[] => {
    const out: Element[] = []
    for (let n = parent.firstChild; n !== null; n = n.nextSibling)
        if (n.nodeType === 1 && names.includes(n.nodeName))
            out.push(n as Element)
    return out
}

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

/**  whether a run sets any character format of its own  */
const runFormatted = (run: Partial<Run>): boolean =>
    [ run.bold, run.italic, run.underline, run.strike, run.superscript, run.subscript, run.size, run.font,
        run.color, run.highlight, run.hyperlink ].some((v) => v !== undefined)

/**  the kept run properties for a run without format of its own: a copy
     without links (they belong to the old text) and spelling marks, in the
     run's language, else the kept one, else the deck's  */
const keptRunProps = (doc: Document, kept: Element, lang: string | undefined, deckLang: string): Element => {
    const rPr = doc.createElementNS(NS_A, "a:rPr")
    for (let i = 0; i < kept.attributes.length; i++) {
        const attr = kept.attributes.item(i)
        if (attr !== null && attr.name !== "err")
            rPr.setAttribute(attr.name, attr.value)
    }
    for (const child of children(kept, [ "a:ln", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill",
        "a:grpFill", "a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill", "a:latin",
        "a:ea", "a:cs", "a:sym", "a:rtl" ]))
        rPr.appendChild(doc.importNode(child, true))
    rPr.setAttribute("lang", lang ?? kept.getAttribute("lang") ?? deckLang)
    return rPr
}

/**  build one `a:r` text run  */
const buildRun = (doc: Document, run: Run, lang: string, kept: KeptFormat | null = null, own?: string): Element => {
    const r = doc.createElementNS(NS_A, "a:r")
    r.appendChild(kept?.rPr !== null && kept?.rPr !== undefined && !runFormatted(run)
        ? keptRunProps(doc, kept.rPr, run.lang ?? own, lang)
        : buildRunProps(doc, run, run.lang ?? own ?? lang))
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
const buildParagraph = (doc: Document, p: Paragraph, lang: string, kept: KeptFormat | null = null): Element => {
    const el = doc.createElementNS(NS_A, "a:p")
    const needsPPr = p.level !== undefined || p.alignment !== undefined || p.bullet === false
        || p.bullet === "number" || p.lineSpacing !== undefined || p.spaceBefore !== undefined
        || p.spaceAfter !== undefined
    if (!needsPPr && p.bullet === undefined && kept?.pPr !== null && kept?.pPr !== undefined)
        el.appendChild(doc.importNode(kept.pPr, true))
    else if (needsPPr) {
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
        el.appendChild(kept === null ? buildRun(doc, run, p.lang ?? lang) : buildRun(doc, run, lang, kept, p.lang))
    return el
}

/**
 *  Replace (or append to) the text of a text body (`p:txBody` of a
 *  shape, `a:txBody` of a table cell).
 *
 *  @param txBody - the text body
 *  @param text - validated rich text payload
 *  @param append - keep existing paragraphs and append instead of replacing
 *  @param lang - the language of runs that state none (the deck language)
 *  @param keep - whether paragraphs and runs that set no format keep the
 *                body's first paragraph and run format (el.set)
 */
export const setBodyText = (txBody: Element, text: RichText, append: boolean, lang: string, keep = false): void => {
    const doc = txBody.ownerDocument as Document
    const first = children(txBody, [ "a:p" ])[0]
    const kept: KeptFormat | null = !keep || first === undefined ? null : {
        pPr: children(first, [ "a:pPr" ])[0] ?? null,
        rPr: children(txBody, [ "a:p" ]).flatMap((p) => children(p, [ "a:r" ]))
            .map((r) => children(r, [ "a:rPr" ])[0] ?? null)[0] ?? children(first, [ "a:endParaRPr" ])[0] ?? null
    }
    if (!append)
        for (const p of children(txBody, [ "a:p" ]))
            txBody.removeChild(p)
    for (const paragraph of toParagraphs(text))
        txBody.appendChild(buildParagraph(doc, paragraph, lang, kept))
}

/**
 *  Replace (or append to) the text content of a shape's txBody.
 *
 *  @param shape - the `p:sp` element to write into
 *  @param text - validated rich text payload
 *  @param append - keep existing paragraphs and append instead of replacing
 *  @param lang - the language of runs that state none (the deck language)
 *  @param keep - whether new text that sets no format keeps the first
 *                paragraph and run format (el.set)
 */
export const setShapeText = (shape: Element, text: RichText, append: boolean, lang: string, keep = false): void => {
    const txBody = firstElement(shape, "p:txBody")
    if (txBody !== null)
        setBodyText(txBody, text, append, lang, keep)
}

/**
 *  Find and replace plain text in a text body, paragraph by paragraph: a
 *  match may span runs (the first touched run takes the replacement and
 *  keeps its format, emptied runs go) but never a field (`a:fld`, e.g. the
 *  slide number) or a line break, whose text stays untouched.
 *
 *  @param txBody - the text body (changed in place)
 *  @param spec - what to replace
 *  @returns the number of replacements and the runs that went (their
 *           links may leave the part)
 */
export const replaceInBody = (txBody: Element, spec: ReplaceSpec): { count: number, dropped: Element[] } => {
    let count = 0
    const dropped: Element[] = []
    for (const p of children(txBody, [ "a:p" ])) {
        const segments: Element[][] = [ [] ]
        for (let n = p.firstChild; n !== null; n = n.nextSibling)
            if (n.nodeType === 1 && n.nodeName === "a:r")
                (segments[segments.length - 1] as Element[]).push(n as Element)
            else if (n.nodeType === 1)
                segments.push([])
        for (const runs of segments.filter((seg) => seg.length > 0)) {
            const ts = runs.map((r) => children(r, [ "a:t" ])[0] ?? null)
            const done = replaceRuns(ts.map((t) => t?.textContent ?? ""), spec)
            if (done.count === 0)
                continue
            count += done.count
            runs.forEach((r, i) => {
                const text = done.texts[i]
                const t = ts[i] ?? null
                if (text === null || text === undefined) {
                    p.removeChild(r)
                    dropped.push(r)
                }
                else if (t !== null && text !== t.textContent) {
                    while (t.firstChild !== null)
                        t.removeChild(t.firstChild)
                    t.appendChild((r.ownerDocument as Document).createTextNode(text))
                }
            })
        }
    }
    return { count, dropped }
}
