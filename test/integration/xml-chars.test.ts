/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: characters XML 1.0 cannot store. A vertical tab in text is
**  the line break inside a paragraph (`a:br`, as PowerPoint and
**  python-pptx write it) in every text writer and reads back as one; every
**  other such character fails validation (E_SCHEMA naming the field), and
**  verify reports a part that carries one -- inherited, it stays as it was
**  (W_INPUT_DEFECT).
*/

import { describe, expect, it } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import JSZip from "jszip"
import { executeOps } from "../../src/commands/apply.js"
import { PptcError } from "../../src/infra/errors.js"
import { verifyFile } from "../../src/engine/verify.js"
import { elements, firstElement } from "../../src/engine/xml.js"
import { OPTS, apply, craftSlide, freshDeck, shapeNamed, slideDoc, stateOf } from "../util/elements.js"

/**  the E_SCHEMA issues of a failing apply: path and message  */
const schemaIssues = async (deck: string, ops: unknown[]): Promise<{ path: string, message: string }[]> => {
    try {
        await executeOps(deck, { ops } as never, OPTS)
    }
    catch (err) {
        if (err instanceof PptcError && err.code === "E_SCHEMA") {
            const issues = (err.details as { issues: { path: (string | number)[], message: string }[] }).issues
            return issues.map((i) => ({ path: i.path.join("."), message: i.message }))
        }
        throw err
    }
    throw new Error("apply succeeded")
}

/**  every XML part of a deck as text  */
const xmlParts = async (deck: string): Promise<Map<string, string>> => {
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const out = new Map<string, string>()
    for (const name of Object.keys(zip.files).filter((n) => /\.(xml|rels)$/.test(n)))
        out.set(name, await (zip.file(name) as JSZip.JSZipObject).async("string"))
    return out
}

/**  the paragraphs of a shape as text: runs as text, line breaks as |  */
const lines = (shape: { getElementsByTagName: unknown }): string[] =>
    elements(shape as never, "a:p").map((p) => {
        let text = ""
        for (let n = p.firstChild; n !== null; n = n.nextSibling)
            if (n.nodeName === "a:r")
                text += firstElement(n as never, "a:t")?.textContent ?? ""
            else if (n.nodeName === "a:br")
                text += "|"
        return text
    })

describe("characters XML cannot store", () => {
    it("rejects them with E_SCHEMA naming the field", async () => {
        const deck = await freshDeck("xml-chars-schema.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT" } ])
        const cases: [ unknown, string, string ][] = [
            [ { op: "slide.fill", slide: "0", placeholders: { title: { text: "A\u0007B" } } }, "ops.0.placeholders.title.text", "U+0007" ],
            [ { op: "slide.fill", slide: "0", notes: "n\u000cx" }, "ops.0.notes", "U+000C" ],
            [ { op: "slide.fill", slide: "0", footer: "f\u0000" }, "ops.0.footer", "U+0000" ],
            [ { op: "meta.props", set: { title: "T\u001f" } }, "ops.0.set.title", "U+001F" ],
            [ { op: "meta.props", set: { custom: { k: "v\u0001" } } }, "ops.0.set.custom.k", "U+0001" ],
            [ { op: "el.set", slide: "0", name: "PptcPh-1", altText: "Bild\u000bzwei" }, "ops.0.altText", "vertical tab" ],
            [ { op: "text.replace", find: "a", replace: "b\u000bc" }, "ops.0.replace", "vertical tab" ],
            [ { op: "text.replace", find: "\u0008", replace: "b" }, "ops.0.find", "U+0008" ],
            [ { op: "slide.fill", slide: "0", placeholders: { body: { text: [ { runs: [ { text: "x\ud800y" } ] } ] } } },
                "ops.0.placeholders.body.text.0.runs.0.text", "U+D800" ],
            [ { op: "el.add", slide: "0", elements: [ { type: "textbox", frame: { x: 1, y: 1, w: 2, h: 1 }, text: "t",
                hyperlink: "https://example.com/\u0007" } ] }, "ops.0.elements.0.hyperlink", "U+0007" ],
            [ { op: "el.add", slide: "0", elements: [ { type: "chart", frame: { x: 1, y: 1, w: 3, h: 2 }, data: { type: "bar",
                categories: [ "a￿" ], series: [ { name: "s", values: [ 1 ] } ] } } ] }, "ops.0.elements.0.data.categories.0", "U+FFFF" ]
        ]
        for (const [ op, path, what ] of cases) {
            const issues = await schemaIssues(deck, [ op ])
            expect(issues.map((i) => i.path), path).toContain(path)
            expect(issues.find((i) => i.path === path)?.message, path).toContain(what)
        }
    })

    it("writes a vertical tab as a line break in every text writer", async () => {
        const deck = await freshDeck("xml-chars-break.pptx")
        await apply(deck, [
            { op: "slide.add", layout: "CONTENT", ref: "s", footer: "Fuß\u000bzeile", notes: "n1\u000bn2\nn3",
                placeholders: { title: { text: "Zeile 1\u000bZeile 2" },
                    body: { text: [ { runs: [ { text: "\u000bvorn" }, { text: "a\u000bb", bold: true }, { text: "hinten\u000b" } ] } ] } } },
            { op: "el.add", slide: "$s", elements: [
                { type: "textbox", name: "Box", frame: { x: 0.5, y: 5.5, w: 3, h: 0.6 }, text: "X\u000bY\nZ" },
                { type: "textbox", name: "Rich", frame: { x: 4, y: 5.5, w: 3, h: 0.6 },
                    text: [ { runs: [ { text: "\u000bp" }, { text: "q\u000br", italic: true } ], bullet: true } ] },
                { type: "shape", name: "Form", shape: "rect", frame: { x: 8, y: 5.5, w: 2, h: 0.6 }, text: "s\u000bt" },
                { type: "table", name: "Tab", frame: { x: 0.5, y: 6.3, w: 4 }, data: { headers: [ "k\u000bl" ], rows: [ [ "c\u000bd" ] ] } }
            ] }
        ])
        for (const [ part, xml ] of await xmlParts(deck))
            expect(xml.includes("\u000b") || /&#(x0*b|0*11);/i.test(xml), part).toBe(false)
        const doc = await slideDoc(deck, 0)
        expect(lines(shapeNamed(doc, "PptcPh-0"))).toEqual([ "Zeile 1|Zeile 2" ])
        expect(lines(shapeNamed(doc, "PptcPh-1"))).toEqual([ "|vorna|bhinten|" ])
        expect(lines(shapeNamed(doc, "Box"))).toEqual([ "X|Y", "Z" ])
        expect(lines(shapeNamed(doc, "Rich"))).toEqual([ "|pq|r" ])
        expect(lines(shapeNamed(doc, "Form"))).toEqual([ "s|t" ])
        expect(lines(shapeNamed(doc, "Tab"))).toEqual([ "k|l", "c|d" ])
        const footer = elements(doc, "p:sp").find((sp) => firstElement(sp, "p:ph")?.getAttribute("type") === "ftr")
        expect(lines(footer as never)).toEqual([ "Fuß|zeile" ])

        /*  the bold run's line break carries the run's format  */
        const br = elements(shapeNamed(doc, "PptcPh-1"), "a:br")[1]
        expect(firstElement(br as never, "a:rPr")?.getAttribute("b")).toBe("1")

        /*  state reads the line breaks back: title, text, runs, notes  */
        const slide = (await stateOf(deck)).slides[0]
        expect(slide?.title).toBe("Zeile 1\u000bZeile 2")
        expect(slide?.notes).toBe("n1\u000bn2\nn3")
        const body = slide?.shapes.find((sh) => sh.name === "PptcPh-1")
        expect(body?.text).toBe("\u000bvorna\u000bbhinten\u000b")
        expect(body?.paragraphs?.[0]?.runs.map((r) => r.text)).toEqual([ "\u000bvorn", "a\u000b", "b", "hinten\u000b" ])

        /*  round trip: the title selects the slide, the text writes back the same  */
        await apply(deck, [ { op: "slide.fill", slide: "title:Zeile 1\u000bZeile 2",
            placeholders: { body: { text: body?.text as string } } } ])
        expect(lines(shapeNamed(await slideDoc(deck, 0), "PptcPh-1"))).toEqual([ "|vorna|bhinten|" ])
    })

    it("verify reports a part with such a character, an apply keeps it as an input defect", async () => {
        const deck = await freshDeck("xml-chars-verify.pptx")
        await apply(deck, [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Eins" } } },
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Zwei" } } } ])
        await craftSlide(deck, 0, (xml) => xml.replace("<a:t>Eins</a:t>", "<a:t>Ei\u000bns&#7;</a:t>"))
        const found = await verifyFile(deck)
        expect(found.join(" ")).toMatch(/malformed XML.*U\+000B/)

        /*  a character reference alone is caught as well  */
        const refOnly = await freshDeck("xml-chars-ref.pptx")
        writeFileSync(refOnly, readFileSync(deck))
        await craftSlide(refOnly, 0, (xml) => xml.replace("Ei\u000bns", "Eins"))
        expect((await verifyFile(refOnly)).join(" ")).toMatch(/&#7;/)

        /*  an apply on another slide: the inherited defect stays, reported  */
        const res = await executeOps(deck, { ops: [ { op: "slide.fill", slide: "1", notes: "ok" } ] } as never, OPTS)
        expect(res.warnings.map((w) => w.code)).toContain("W_INPUT_DEFECT")
    })
})
