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

/**  DrawingML main namespace (`a:`)  */
export const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
/**  PresentationML namespace (`p:`)  */
export const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
/**  OfficeDocument relationships namespace (`r:`)  */
export const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
/**  Package relationships namespace (rels parts)  */
export const NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

/**
 *  Parse an XML string into a DOM document. A leading byte-order mark
 *  (written by some .NET tools and 6 of PowerPoint's own themes) is
 *  dropped -- xmldom would reject the declaration behind it.
 *
 *  @param xml - XML source text
 *  @param strict - whether recoverable errors fail as well (verify)
 *  @returns parsed document
 */
export const parseXml = (xml: string, strict = false): Document =>
    new DOMParser({
        onError: (level, msg): void => {
            /*  lenient: recoverable errors are tolerated silently (no stderr
                noise), only fatal ones fail; strict: errors fail too --
                warnings never (e.g. a U+FFFD character in valid text)  */
            if (level === "fatalError" || (strict && level === "error"))
                throw new Error(msg)
        }
    }).parseFromString(xml.replace(/^\uFEFF/, ""), "application/xml")

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
 *  Concatenated text content of all `a:t` runs below a scope element.
 *  Paragraphs (`a:p`) are joined with newlines.
 *
 *  @param scope - element containing DrawingML text (e.g. a txBody)
 *  @returns plain text, empty string when there is none
 */
export const drawingText = (scope: Element): string => {
    const parts: string[] = []
    for (const p of elements(scope, "a:p")) {
        const runs = elements(p, "a:t").map((t) => t.textContent ?? "")
        parts.push(runs.join(""))
    }
    return parts.join("\n")
}
