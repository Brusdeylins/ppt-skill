/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/xml: thin OOXML/XML utilities shared by the engine modules --
**  parsing, serializing, namespace constants and element iteration.
*/

import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { Attr, Document, Element } from "@xmldom/xmldom"
import { isSchemeColor, unicodeName, xmlIllegalAt } from "../schema/payloads.js"

/**  DrawingML main namespace (`a:`)  */
export const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
/**  PresentationML namespace (`p:`)  */
export const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
/**  OfficeDocument relationships namespace (`r:`)  */
export const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
/**  Package relationships namespace (rels parts)  */
export const NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

/**  whether a code point is a character XML 1.0 allows  */
const xmlChar = (cp: number): boolean =>
    cp === 0x9 || cp === 0xA || cp === 0xD || (cp >= 0x20 && cp <= 0xD7FF) || (cp >= 0xE000 && cp <= 0xFFFD)
        || (cp >= 0x10000 && cp <= 0x10FFFF)

/**
 *  The first character an XML 1.0 document must not contain, raw or as a
 *  character reference (`&#11;`) -- xmldom accepts both, PowerPoint
 *  rejects the part.
 *
 *  @param xml - XML text
 *  @returns a description of the character and its offset, null when none
 */
export const illegalXmlChar = (xml: string): string | null => {
    const at = xmlIllegalAt(xml, false)
    if (at >= 0)
        return `character ${unicodeName(xml.charCodeAt(at))} not allowed in XML (offset ${at})`
    for (const ref of xml.matchAll(/&#(x[0-9A-Fa-f]+|[0-9]+);/g)) {
        const body = ref[1] as string
        const cp = body.startsWith("x") ? parseInt(body.slice(1), 16) : parseInt(body, 10)
        if (!xmlChar(cp))
            return `character reference ${ref[0]} (${unicodeName(cp)}) not allowed in XML (offset ${ref.index})`
    }
    return null
}

/**
 *  A text with the characters XML cannot store left out -- a vertical tab
 *  (a line break inside a paragraph) as a space -- for text pptc copies
 *  from the deck into another part (the slide titles of the app
 *  properties).
 *
 *  @param s - the text
 *  @returns the storable text
 */
export const xmlStorable = (s: string): string => {
    let out = ""
    let from = 0
    for (let at = xmlIllegalAt(s, false); at >= 0; at = xmlIllegalAt(s.slice(from), false)) {
        out += s.slice(from, from + at) + (s.charCodeAt(from + at) === 0x0B ? " " : "")
        from += at + 1
    }
    return out + s.slice(from)
}

/**
 *  Parse an XML string into a DOM document. A leading byte-order mark
 *  (written by some .NET tools and 6 of PowerPoint's own themes) is
 *  dropped -- xmldom would reject the declaration behind it.
 *
 *  @param xml - XML source text
 *  @param strict - whether recoverable errors and characters XML does not
 *                 allow fail as well (verify)
 *  @returns parsed document
 */
export const parseXml = (xml: string, strict = false): Document => {
    const illegal = strict ? illegalXmlChar(xml) : null
    if (illegal !== null)
        throw new Error(illegal)
    return new DOMParser({
        onError: (level, msg): void => {
            /*  lenient: recoverable errors are tolerated silently (no stderr
                noise), only fatal ones fail; strict: everything fails but
                the warning about a U+FFFD character, which is valid text  */
            if (level === "fatalError" || (strict && !/replacement character/i.test(msg)))
                throw new Error(msg)
        }
    }).parseFromString(xml.replace(/^\uFEFF/, ""), "application/xml")
}

/**
 *  Decode the bytes of an XML part: UTF-8 (with or without byte-order
 *  mark) or UTF-16 (little or big endian, as some tools write custom XML).
 *
 *  @param bytes - the raw part content
 *  @returns the XML text without byte-order mark
 */
export const decodeXml = (bytes: Uint8Array): string => {
    const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    if (buf[0] === 0xFF && buf[1] === 0xFE)
        return buf.subarray(2).toString("utf16le")
    if (buf[0] === 0xFE && buf[1] === 0xFF)
        return Buffer.from(buf.subarray(2)).swap16().toString("utf16le")
    if (buf[0] === 0x3C && buf[1] === 0x00)
        return buf.toString("utf16le")
    return buf.toString("utf8").replace(/^\uFEFF/, "")
}

/**
 *  An XML text about to be written as UTF-8, with a UTF-16 encoding
 *  declaration of its decoded source corrected.
 *
 *  @param xml - XML text
 *  @returns the text declaring UTF-8 where it declared UTF-16
 */
export const utf8Declared = (xml: string): string =>
    xml.replace(/^(<\?xml[^>]*encoding=["'])UTF-16(?:LE|BE)?(["'])/i, "$1UTF-8$2")

/**
 *  Serialize a DOM document back to its XML string. The text is written as
 *  UTF-8, so a UTF-16 declaration of a decoded part is corrected.
 *
 *  @param doc - DOM document
 *  @returns XML text
 */
export const serializeXml = (doc: Document): string =>
    utf8Declared(new XMLSerializer().serializeToString(doc))

/**
 *  Escape text for an XML text node, or for a double-quoted attribute value.
 *
 *  @param s - plain text
 *  @param attr - whether the text goes into an attribute value
 *  @returns escaped text
 */
export const xmlEscape = (s: string, attr = false): string => {
    const text = s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    return attr ? text.replace(/"/g, "&quot;") : text
}

/**
 *  Every relationship reference (`r:id`, `r:embed`, `r:link`, ... -- any
 *  attribute in the relationships namespace) of an element subtree.
 *
 *  @param scope - the subtree root
 *  @returns the referencing elements with their attribute, in document order
 */
export const relRefs = (scope: Element): { el: Element, attr: Attr }[] => {
    const refs: { el: Element, attr: Attr }[] = []
    const walk = (el: Element): void => {
        for (let i = 0; i < el.attributes.length; i++) {
            const attr = el.attributes.item(i)
            if (attr !== null && attr.namespaceURI === NS_R)
                refs.push({ el, attr })
        }
        for (let n = el.firstChild; n !== null; n = n.nextSibling)
            if (n.nodeType === 1)
                walk(n as Element)
    }
    walk(scope)
    return refs
}

/**
 *  Whether an element lies inside an `mc:Fallback` -- the alternative
 *  PowerPoint writes for older readers, which repeats the shape ids of the
 *  `mc:Choice` content by design.
 *
 *  @param el - the element
 *  @returns true inside a fallback
 */
export const inFallback = (el: Element): boolean => {
    for (let up = el.parentNode as Element | null; up !== null && up.nodeType === 1; up = up.parentNode as Element | null)
        if (up.nodeName === "mc:Fallback")
            return true
    return false
}

/**
 *  Serialize a single element subtree to its XML string.
 *
 *  @param el - DOM element
 *  @returns XML text of the subtree
 */
export const serializeElement = (el: Element): string =>
    new XMLSerializer().serializeToString(el)

/**
 *  Collect all descendant elements with a given qualified tag name.
 *
 *  @param scope - document or element to search
 *  @param qname - qualified tag name like "p:sp"
 *  @returns elements in document order
 */
export const elements = (scope: Document | Element, qname: string): Element[] => {
    const list = scope.getElementsByTagName(qname)
    const out: Element[] = []
    for (let i = 0; i < list.length; i++)
        out.push(list.item(i) as Element)
    return out
}

/**
 *  First descendant element with a given qualified tag name, or null.
 *
 *  @param scope - document or element to search
 *  @param qname - qualified tag name like "p:txBody"
 *  @returns the first matching element or null
 */
export const firstElement = (scope: Document | Element, qname: string): Element | null =>
    scope.getElementsByTagName(qname).item(0) as Element | null

/**
 *  The text of one paragraph: its `a:t` texts in order, a line break
 *  (`a:br`) as a vertical tab -- the form the text writers take it in.
 *
 *  @param p - the `a:p` element
 *  @returns the paragraph text
 */
export const paragraphText = (p: Element): string => {
    let text = ""
    const walk = (el: Element): void => {
        for (let n = el.firstChild; n !== null; n = n.nextSibling)
            if (n.nodeType !== 1)
                continue
            else if (n.nodeName === "a:t")
                text += n.textContent ?? ""
            else if (n.nodeName === "a:br")
                text += "\v"
            else
                walk(n as Element)
    }
    walk(p)
    return text
}

/**
 *  Concatenated text content of all `a:t` runs below a scope element.
 *  Paragraphs (`a:p`) are joined with newlines, a line break inside a
 *  paragraph (`a:br`) reads as a vertical tab.
 *
 *  @param scope - element containing DrawingML text (e.g. a txBody)
 *  @returns plain text, empty string when there is none
 */
export const drawingText = (scope: Element): string =>
    elements(scope, "a:p").map(paragraphText).join("\n")

/**
 *  Build the DrawingML color element of a validated color: `a:srgbClr`
 *  for a hex value, `a:schemeClr` for a theme color slot.
 *
 *  @param doc - the owning document
 *  @param color - RRGGBB or a theme color slot name
 *  @returns the color element
 */
export const colorElement = (doc: Document, color: string): Element => {
    const el = doc.createElementNS(NS_A, isSchemeColor(color) ? "a:schemeClr" : "a:srgbClr")
    el.setAttribute("val", color)
    return el
}
