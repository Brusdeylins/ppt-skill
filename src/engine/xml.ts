/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/xml: thin OOXML/XML utilities shared by the engine modules --
**  parsing, serializing, namespace constants and element iteration.
*/

import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { Document, Element } from "@xmldom/xmldom"

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
 *  @returns parsed document
 */
export const parseXml = (xml: string): Document =>
    new DOMParser().parseFromString(xml.replace(/^\uFEFF/, ""), "application/xml")

/**
 *  Serialize a DOM document back to its XML string.
 *
 *  @param doc - DOM document
 *  @returns XML text
 */
export const serializeXml = (doc: Document): string =>
    new XMLSerializer().serializeToString(doc)

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
