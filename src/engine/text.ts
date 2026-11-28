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

/**  marker attribute the engine replaces by a hyperlink relationship: a
     URL, or a slide link (`slide:SEL`) the plan resolved  */
export const HLINK_ATTR = "pptc-hlink"

/**
 *  Set or clear the click target of a whole shape or picture (`a:hlinkClick`
 *  in its `p:cNvPr`); the relationship is wired with the slide's other
 *  links.
 *
 *  @param shape - the shape or picture
 *  @param target - a URL or slide link, null or "" to remove the link
 */
export const setClickLink = (shape: Element, target: string | null): void => {
    const cNvPr = firstElement(shape, "p:cNvPr")
    if (cNvPr === null)
        return
    for (const old of children(cNvPr, [ "a:hlinkClick" ]))
        cNvPr.removeChild(old)
    if (target === null || target === "")
        return
    const hlink = (shape.ownerDocument as Document).createElementNS(NS_A, "a:hlinkClick")
    hlink.setAttribute(HLINK_ATTR, target)
    cNvPr.insertBefore(hlink, cNvPr.firstChild)
}

/**  normalize plain-string content into the paragraph form  */
const toParagraphs = (text: RichText): Paragraph[] =>
    typeof text === "string"
        ? text.split("\n").map((line) => ({ text: line }))
        : text

/**  text baseline shift of superscript and subscript runs, in 1/1000 %
     (PowerPoint's own values)  */
export const BASELINE = { superscript: 30000, subscript: -25000 }

/**  the schema order of the `a:rPr` children  */
const RPR_ORDER = [ "a:ln", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill",
    "a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill", "a:latin", "a:ea",
    "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst" ]

/**  the schema order of the `a:pPr` children  */
const PPR_ORDER = [ "a:lnSpc", "a:spcBef", "a:spcAft", "a:buClrTx", "a:buClr", "a:buSzTx", "a:buSzPct", "a:buSzPts",
    "a:buFontTx", "a:buFont", "a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip", "a:tabLst", "a:defRPr", "a:extLst" ]

/**  the fills of a run (one at most)  */
const RUN_FILLS = [ "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill" ]

/**  the bullet kinds of a paragraph (one at most)  */
const BULLETS = [ "a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip" ]

/**  the bullet indent of a box without list-style bullets in EMU: 27 pt,
     as PowerPoint and PptxGenJS write it, one more per level  */
const BULLET_INDENT = 342900

/**  put a child into its place by the schema order of its parent, after
     removing the children of the given names  */
const place = (parent: Element, child: Element, order: string[], replaces: string[]): void => {
    for (const old of children(parent, replaces))
        parent.removeChild(old)
    const rank = order.indexOf(child.nodeName)
    parent.insertBefore(child, children(parent, order.slice(rank + 1))[0] ?? null)
}

/**  apply a run's own character format to run properties: on a fresh
     `a:rPr` a `false` underline or baseline states nothing, on kept
     properties (`merge`) it clears the kept one  */
const applyRunFormat = (doc: Document, rPr: Element, run: Partial<Run>, merge: boolean): Element => {
    if (run.bold !== undefined)
        rPr.setAttribute("b", run.bold ? "1" : "0")
    if (run.italic !== undefined)
        rPr.setAttribute("i", run.italic ? "1" : "0")
    if (run.underline === true)
        rPr.setAttribute("u", "sng")
    else if (run.underline === false && merge)
        rPr.setAttribute("u", "none")
    if (run.strike !== undefined)
        rPr.setAttribute("strike", run.strike ? "sngStrike" : "noStrike")
    const baseline = Number(rPr.getAttribute("baseline") ?? "0")
    if (run.superscript === true)
        rPr.setAttribute("baseline", String(BASELINE.superscript))
    else if (run.subscript === true)
        rPr.setAttribute("baseline", String(BASELINE.subscript))
    else if ((run.superscript === false && baseline > 0) || (run.subscript === false && baseline < 0))
        rPr.removeAttribute("baseline")
    if (run.size !== undefined)
        rPr.setAttribute("sz", String(Math.round(run.size * 100)))
    if (run.color !== undefined) {
        const fill = doc.createElementNS(NS_A, "a:solidFill")
        fill.appendChild(colorElement(doc, run.color))
        place(rPr, fill, RPR_ORDER, RUN_FILLS)
    }
    if (run.highlight !== undefined) {
        const highlight = doc.createElementNS(NS_A, "a:highlight")
        highlight.appendChild(colorElement(doc, run.highlight))
        place(rPr, highlight, RPR_ORDER, [ "a:highlight" ])
    }
    if (run.font !== undefined) {
        const latin = doc.createElementNS(NS_A, "a:latin")
        latin.setAttribute("typeface", run.font)
        place(rPr, latin, RPR_ORDER, [ "a:latin" ])
    }
    if (run.hyperlink !== undefined) {
        /*  the relationship id is assigned by the engine; mark with the URL  */
        const hlink = doc.createElementNS(NS_A, "a:hlinkClick")
        hlink.setAttribute(HLINK_ATTR, run.hyperlink)
        place(rPr, hlink, RPR_ORDER, [ "a:hlinkClick" ])
    }
    return rPr
}

/**  build the `a:rPr` run properties element (always: it carries the
     run's language); children in schema order: fill, highlight, latin,
     hlinkClick  */
const buildRunProps = (doc: Document, run: Partial<Run>, lang: string): Element => {
    const rPr = doc.createElementNS(NS_A, "a:rPr")
    rPr.setAttribute("lang", run.lang ?? lang)
    return applyRunFormat(doc, rPr, run, false)
}

/**  whether a run sets any character format of its own  */
const runFormatted = (run: Partial<Run>): boolean =>
    [ run.bold, run.italic, run.underline, run.strike, run.superscript, run.subscript, run.size, run.font,
        run.color, run.highlight, run.hyperlink ].some((v) => v !== undefined)

/**  the kept run properties: a copy without links (they belong to the old
     text) and spelling marks, in the run's language, else the kept one,
     else the deck's  */
const keptRunProps = (doc: Document, kept: Element, lang: string | undefined, deckLang: string): Element => {
    const rPr = doc.createElementNS(NS_A, "a:rPr")
    for (let i = 0; i < kept.attributes.length; i++) {
        const attr = kept.attributes.item(i)
        if (attr !== null && attr.name !== "err")
            rPr.setAttribute(attr.name, attr.value)
    }
    for (const child of children(kept, RPR_ORDER.filter((name) => ![ "a:hlinkClick", "a:hlinkMouseOver", "a:extLst" ].includes(name))))
        rPr.appendChild(doc.importNode(child, true))
    rPr.setAttribute("lang", lang ?? kept.getAttribute("lang") ?? deckLang)
    return rPr
}

/**  the properties of a line break (`a:br`) inside a run: the run's own,
     without its click target  */
const breakProps = (rPr: Element): Element => {
    const props = rPr.cloneNode(true) as Element
    for (const link of children(props, [ "a:hlinkClick", "a:hlinkMouseOver" ]))
        props.removeChild(link)
    return props
}

/**  build the `a:r` text runs of one run; a vertical tab in its text is a
     line break inside the paragraph (`a:br` with the run's properties)  */
const buildRun = (doc: Document, run: Run, lang: string, kept: KeptFormat | null = null, own?: string): Element[] => {
    const rPr = kept?.rPr === null || kept?.rPr === undefined
        ? buildRunProps(doc, run, run.lang ?? own ?? lang)
        : runFormatted(run)
            ? applyRunFormat(doc, keptRunProps(doc, kept.rPr, run.lang ?? own, lang), run, true)
            : keptRunProps(doc, kept.rPr, run.lang ?? own, lang)
    const pieces = run.text.split("\v")
    const out: Element[] = []
    pieces.forEach((piece, i) => {
        if (i > 0) {
            const br = doc.createElementNS(NS_A, "a:br")
            br.appendChild(breakProps(rPr))
            out.push(br)
        }
        if (piece === "" && pieces.length > 1)
            return
        const r = doc.createElementNS(NS_A, "a:r")
        r.appendChild(rPr.cloneNode(true))
        const t = doc.createElementNS(NS_A, "a:t")
        t.appendChild(doc.createTextNode(piece))
        r.appendChild(t)
        out.push(r)
    })
    return out
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

/**  what the list style a placeholder inherits gives one paragraph
     level: a bullet (character, picture or number) and a hanging indent  */
export interface StyleLevel {
    bullet: boolean
    hanging: boolean
}

/**  the inherited list style of a placeholder by paragraph level (0-8)  */
export type ListStyle = (lvl: number) => StyleLevel

/**  the level properties (`a:lvlNpPr`) of a list style, null when absent  */
const levelProps = (style: Element | null, lvl: number): Element | null =>
    style === null ? null : children(style, [ `a:lvl${lvl + 1}pPr` ])[0] ?? null

/**  the list style of a shape's text body, null when it has none  */
const ownListStyle = (shape: Element | null): Element | null => {
    const txBody = shape === null ? null : children(shape, [ "p:txBody" ])[0] ?? null
    return txBody === null ? null : children(txBody, [ "a:lstStyle" ])[0] ?? null
}

/**  the placeholder shapes of a layout or master  */
const placeholdersOf = (doc: Document | null): { sp: Element, ph: Element }[] => {
    const tree = doc === null ? null : firstElement(doc, "p:spTree")
    return (tree === null ? [] : children(tree, [ "p:sp" ])).flatMap((sp) => {
        const ph = firstElement(sp, "p:ph")
        return ph === null ? [] : [ { sp, ph } ]
    })
}

/**
 *  The list style a placeholder inherits, level by level: its own
 *  `a:lstStyle`, then the layout placeholder's (by idx, else by type),
 *  the master placeholder's (by type) and the master text style (title,
 *  body or other). The first level that states a bullet (or `a:buNone`)
 *  decides the bullet, the first that states an indent the hanging indent.
 *
 *  @param shape - the placeholder shape on the slide
 *  @param layout - the slide's layout DOM, null when unknown
 *  @param master - the layout's master DOM, null when unknown
 *  @returns the inherited list style
 */
export const placeholderListStyle = (shape: Element, layout: Document | null, master: Document | null): ListStyle => {
    const ph = firstElement(shape, "p:ph")
    const type = ph?.getAttribute("type") ?? "obj"
    const idx = ph?.getAttribute("idx") ?? "0"
    const typeOf = (el: Element): string => el.getAttribute("type") ?? "obj"
    const inLayout = placeholdersOf(layout)
    const layoutSp = (inLayout.find((p) => (p.ph.getAttribute("idx") ?? "0") === idx && typeOf(p.ph) === type)
        ?? inLayout.find((p) => (p.ph.getAttribute("idx") ?? "0") === idx)
        ?? inLayout.find((p) => typeOf(p.ph) === type))?.sp ?? null
    const kind = [ "title", "ctrTitle" ].includes(type) ? "title" : [ "dt", "ftr", "sldNum", "hdr" ].includes(type) ? type : "body"
    const masterSp = placeholdersOf(master).find((p) => typeOf(p.ph) === kind)?.sp ?? null
    const txStyles = master === null ? null : firstElement(master, "p:txStyles")
    const textStyle = txStyles === null ? null : children(txStyles, [
        kind === "title" ? "p:titleStyle" : kind === "body" ? "p:bodyStyle" : "p:otherStyle" ])[0] ?? null
    const chain = [ ownListStyle(shape), ownListStyle(layoutSp), ownListStyle(masterSp), textStyle ]
    return (lvl) => {
        const levels = chain.map((style) => levelProps(style, lvl)).filter((l): l is Element => l !== null)
        const bullet = levels.map((l) => children(l, BULLETS)[0]).find((b) => b !== undefined)
        const indent = levels.find((l) => l.hasAttribute("indent"))?.getAttribute("indent") ?? "0"
        return { bullet: bullet !== undefined && bullet.nodeName !== "a:buNone", hanging: Number(indent) < 0 }
    }
}

/**  apply a paragraph's bullet to its properties: `false` none, `"number"`
     a numbered list, `true` a bullet -- the list style's own where the
     placeholder's style shows one at the level, else a bullet character
     with its hanging indent (unless the style has one); a kept bullet
     character or picture stays  */
const applyBullet = (doc: Document, pPr: Element, bullet: boolean | "number", style: StyleLevel): void => {
    if (bullet === false) {
        place(pPr, doc.createElementNS(NS_A, "a:buNone"), PPR_ORDER, BULLETS)
        return
    }
    if (bullet === "number") {
        const num = doc.createElementNS(NS_A, "a:buAutoNum")
        num.setAttribute("type", "arabicPeriod")
        place(pPr, num, PPR_ORDER, BULLETS)
        return
    }
    for (const old of children(pPr, [ "a:buNone", "a:buAutoNum" ]))
        pPr.removeChild(old)
    if (style.bullet)
        return
    if (pPr.hasAttribute("indent") ? Number(pPr.getAttribute("indent")) >= 0 : !style.hanging) {
        /*  no hanging indent yet (PptxGenJS writes 0/0 for text without bullets)  */
        pPr.setAttribute("marL", String(BULLET_INDENT * (Number(pPr.getAttribute("lvl") ?? "0") + 1)))
        pPr.setAttribute("indent", String(-BULLET_INDENT))
    }
    if (children(pPr, [ "a:buChar", "a:buBlip" ]).length > 0)
        return
    if (children(pPr, [ "a:buSzTx", "a:buSzPct", "a:buSzPts" ]).length === 0) {
        const size = doc.createElementNS(NS_A, "a:buSzPct")
        size.setAttribute("val", "100000")
        place(pPr, size, PPR_ORDER, [])
    }
    const char = doc.createElementNS(NS_A, "a:buChar")
    char.setAttribute("char", "•")
    place(pPr, char, PPR_ORDER, [])
}

/**  build one `a:p` paragraph including paragraph properties; its runs
     take the paragraph's language, else the given one. The paragraph's
     own fields go onto the kept first paragraph's properties (el.set),
     else onto fresh ones.  */
const buildParagraph = (doc: Document, p: Paragraph, lang: string, kept: KeptFormat | null, styleAt: ListStyle): Element => {
    const el = doc.createElementNS(NS_A, "a:p")
    const keptLevel = Number(kept?.pPr?.getAttribute("lvl") ?? "0")
    const style = styleAt(p.level ?? keptLevel)
    const needsPPr = p.level !== undefined || p.alignment !== undefined || p.bullet === false
        || p.bullet === "number" || (p.bullet === true && !style.bullet) || p.lineSpacing !== undefined
        || p.spaceBefore !== undefined || p.spaceAfter !== undefined
    const pPr = kept?.pPr !== null && kept?.pPr !== undefined ? doc.importNode(kept.pPr, true) as Element
        : needsPPr ? doc.createElementNS(NS_A, "a:pPr") : null
    if (pPr !== null) {
        /*  a kept indent belongs to the kept level: another level takes
            its list style's own  */
        const relevel = p.level !== undefined && kept?.pPr !== null && kept?.pPr !== undefined && p.level !== keptLevel
        if (relevel) {
            pPr.removeAttribute("marL")
            pPr.removeAttribute("indent")
        }
        if (p.level !== undefined)
            pPr.setAttribute("lvl", String(p.level))

        /*  a kept bullet character gets the hanging indent of the new level  */
        if (relevel && p.bullet === undefined && children(pPr, [ "a:buChar", "a:buBlip" ]).length > 0)
            applyBullet(doc, pPr, true, style)
        if (p.alignment !== undefined)
            pPr.setAttribute("algn", { left: "l", center: "ctr", right: "r", justify: "just" }[p.alignment])
        if (p.lineSpacing !== undefined)
            place(pPr, spacing(doc, "a:lnSpc", "a:spcPct", p.lineSpacing * 100000), PPR_ORDER, [ "a:lnSpc" ])
        if (p.spaceBefore !== undefined)
            place(pPr, spacing(doc, "a:spcBef", "a:spcPts", p.spaceBefore * 100), PPR_ORDER, [ "a:spcBef" ])
        if (p.spaceAfter !== undefined)
            place(pPr, spacing(doc, "a:spcAft", "a:spcPts", p.spaceAfter * 100), PPR_ORDER, [ "a:spcAft" ])
        if (p.bullet !== undefined)
            applyBullet(doc, pPr, p.bullet, style)
        el.appendChild(pPr)
    }
    const runs: Run[] = p.runs ?? [ paragraphRun(p) ]
    for (const run of runs)
        for (const node of kept === null ? buildRun(doc, run, p.lang ?? lang) : buildRun(doc, run, lang, kept, p.lang))
            el.appendChild(node)
    return el
}

/**
 *  Replace (or append to) the text of a text body (`p:txBody` of a
 *  shape, `a:txBody` of a table cell). A vertical tab in the text is a
 *  line break inside its paragraph (`a:br`). Replaced text drops the
 *  shrink-on-overflow scale PowerPoint computed for the old one
 *  (`a:normAutofit` stays, without `fontScale` and `lnSpcReduction`).
 *
 *  @param txBody - the text body
 *  @param text - validated rich text payload
 *  @param append - keep existing paragraphs and append instead of replacing
 *  @param lang - the language of runs that state none (the deck language)
 *  @param keep - whether the body's first paragraph and run format stays
 *                under the new text, its own fields merged in (el.set)
 *  @param styleBullets - whether the body's list style brings its own
 *                bullets (`true`: at every level), or the inherited list
 *                style of a placeholder by level: where it shows a bullet,
 *                `bullet: true` writes none
 */
export const setBodyText = (txBody: Element, text: RichText, append: boolean, lang: string, keep = false,
    styleBullets: boolean | ListStyle = false): void => {
    const doc = txBody.ownerDocument as Document
    const styleAt: ListStyle = typeof styleBullets === "function" ? styleBullets
        : () => ({ bullet: styleBullets, hanging: false })
    const first = children(txBody, [ "a:p" ])[0]
    const kept: KeptFormat | null = !keep || first === undefined ? null : {
        pPr: children(first, [ "a:pPr" ])[0] ?? null,
        rPr: children(txBody, [ "a:p" ]).flatMap((p) => children(p, [ "a:r" ]))
            .map((r) => children(r, [ "a:rPr" ])[0] ?? null)[0] ?? children(first, [ "a:endParaRPr" ])[0] ?? null
    }
    if (!append) {
        for (const p of children(txBody, [ "a:p" ]))
            txBody.removeChild(p)
        for (const fit of children(txBody, [ "a:bodyPr" ]).flatMap((bodyPr) => children(bodyPr, [ "a:normAutofit" ]))) {
            fit.removeAttribute("fontScale")
            fit.removeAttribute("lnSpcReduction")
        }
    }
    for (const paragraph of toParagraphs(text))
        txBody.appendChild(buildParagraph(doc, paragraph, lang, kept, styleAt))
}

/**
 *  Replace (or append to) the text content of a shape's txBody. A
 *  placeholder keeps the bullets of its list style.
 *
 *  @param shape - the `p:sp` element to write into
 *  @param text - validated rich text payload
 *  @param append - keep existing paragraphs and append instead of replacing
 *  @param lang - the language of runs that state none (the deck language)
 *  @param keep - whether new text keeps the first paragraph and run
 *                format, its own fields merged in (el.set)
 *  @param listStyle - the list style a placeholder inherits
 *                (`placeholderListStyle`); unknown: bullets at every level
 */
export const setShapeText = (shape: Element, text: RichText, append: boolean, lang: string, keep = false,
    listStyle: ListStyle | null = null): void => {
    const txBody = firstElement(shape, "p:txBody")
    if (txBody !== null)
        setBodyText(txBody, text, append, lang, keep, firstElement(shape, "p:ph") === null ? false : listStyle ?? true)
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
