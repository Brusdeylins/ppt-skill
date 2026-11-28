/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: meta.props `custom` patches docProps/custom.xml. Only the
**  named properties change (string values); every other property -- a
**  number, a date, a boolean, a linked one -- stays exactly as it was.
*/

import { describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import JSZip from "jszip"
import { elements, parseXml, serializeElement } from "../../src/engine/xml.js"
import { apply, freshDeck } from "../util/elements.js"
import { propsOf } from "../util/fingerprint.js"

const FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

/**  a custom.xml with properties of other types than a string  */
const CUSTOM = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\""
    + " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
    + `<property fmtid="${FMTID}" pid="2" name="Count"><vt:i4>42</vt:i4></property>`
    + `<property fmtid="${FMTID}" pid="3" name="Approved"><vt:bool>true</vt:bool></property>`
    + `<property fmtid="${FMTID}" pid="5" name="Due"><vt:filetime>2026-11-28T00:00:00Z</vt:filetime></property>`
    + `<property fmtid="${FMTID}" pid="6" name="Linked" linkTarget="_Toc1"><vt:lpwstr>Kapitel</vt:lpwstr></property>`
    + `<property fmtid="${FMTID}" pid="7" name="Owner"><vt:lpwstr>Alt</vt:lpwstr></property>`
    + `<property fmtid="${FMTID}" pid="8" name="Score"><vt:r8>1.5</vt:r8></property>`
    + "</Properties>"

/**  the property nodes of a deck's custom.xml by name, as XML text (the
     namespace declarations a lone element serializes with left out)  */
const propertyNodes = async (deck: string): Promise<Map<string, string>> => {
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const doc = parseXml(await (zip.file("docProps/custom.xml") as JSZip.JSZipObject).async("string"))
    return new Map(elements(doc, "property").map((p) => [ p.getAttribute("name") ?? "",
        serializeElement(p).replace(/ xmlns(:\w+)?="[^"]*"/g, "") ]))
}

describe("meta.props custom", () => {
    it("patches only the named properties and keeps every other one as it is", async () => {
        const deck = await freshDeck("custom-props-patch.pptx")
        await apply(deck, [ { op: "meta.props", set: { custom: { Owner: "Start" } } } ])
        const zip = await JSZip.loadAsync(readFileSync(deck))
        zip.file("docProps/custom.xml", CUSTOM)
        writeFileSync(deck, await zip.generateAsync({ type: "nodebuffer" }))
        const before = await propertyNodes(deck)

        await apply(deck, [ { op: "meta.props", set: { custom: { Owner: "Neu", Score: "hoch", Extra: "" } } } ])
        const after = await propertyNodes(deck)
        for (const name of [ "Count", "Approved", "Due", "Linked" ])
            expect(after.get(name), name).toBe(before.get(name))
        const raw = await (await JSZip.loadAsync(readFileSync(deck))).file("docProps/custom.xml")?.async("string") ?? ""
        const kept = CUSTOM.match(/<property [^>]*name="(Count|Approved|Due|Linked)"[^>]*>.*?<\/property>/g) ?? []
        expect(kept.length).toBe(4)
        for (const node of kept)
            expect(raw).toContain(node)
        expect(after.get("Owner")).toMatch(/pid="7" name="Owner"><vt:lpwstr>Neu<\/vt:lpwstr>/)
        expect(after.get("Score")).toMatch(/pid="8" name="Score"><vt:lpwstr>hoch<\/vt:lpwstr>/)
        expect(after.get("Extra")).toMatch(/pid="9" name="Extra"><vt:lpwstr\/?>(<\/vt:lpwstr>)?/)
        expect([ ...after.keys() ]).toEqual([ "Count", "Approved", "Due", "Linked", "Owner", "Score", "Extra" ])
        expect((await propsOf(deck)).custom).toMatchObject({ Owner: "Neu", Score: "hoch", Extra: "", Linked: "Kapitel" })
    })

    it("creates the part with string properties when the deck has none", async () => {
        const deck = await freshDeck("custom-props-new.pptx")
        const zip = await JSZip.loadAsync(readFileSync(deck))
        expect(zip.file("docProps/custom.xml")).toBeNull()
        await apply(deck, [ { op: "meta.props", set: { custom: { a: "1", b: "<2>" } } } ])
        const nodes = await propertyNodes(deck)
        expect([ ...nodes.values() ]).toEqual([
            `<property fmtid="${FMTID}" pid="2" name="a"><vt:lpwstr>1</vt:lpwstr></property>`,
            `<property fmtid="${FMTID}" pid="3" name="b"><vt:lpwstr>&lt;2&gt;</vt:lpwstr></property>`
        ])
    })
})
