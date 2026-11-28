/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the content guard around text edits. A text.replace edits
**  run text only -- a picture bullet (pPr) and a picture cell fill (tcPr)
**  stay, so a picture that vanishes after it is a loss; `cells` keeps the
**  cell fill. Text that replaces a whole body may take a picture bullet
**  along. Tampered outputs show the guard at work: a picture or chart an
**  edit dropped fails the apply, a foreign image is blocked, and an image
**  adopted with layout content must be exactly the declared one; a
**  footer-band placeholder cloned anew or removed may take its image along.
**  Speaker notes may change only on slides where a replacement counted.
*/

import { beforeAll, describe, expect, it, vi } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import type { Document, Element } from "@xmldom/xmldom"
import { executeOps } from "../../src/commands/apply.js"
import { digestBytes, verifyContent, type SlideExpect } from "../../src/engine/guard.js"
import { relTargets, relsPartOf } from "../../src/engine/parts.js"
import { PptcError } from "../../src/infra/errors.js"
import { NS_A, NS_R, elements, firstElement, parseXml, serializeXml } from "../../src/engine/xml.js"
import { OPTS, TMP, apply, craftSlide, failure, freshDeck, png, shapeNamed, slideDoc, stateOf } from "../util/elements.js"

/*  sabotage: the edit functions drop the element of this name afterwards,
    and a notes replacement that found nothing changes the notes anyway,
    as an engine bug would  */
const sabotage = vi.hoisted(() => ({ name: null as string | null, notes: false }))
vi.mock("../../src/engine/inplace/edits.js", async (importOriginal) => {
    const mod = await importOriginal<typeof import("../../src/engine/inplace/edits.js")>()
    const drop = (doc: Document): void => {
        const cNvPr = sabotage.name === null ? undefined : elements(doc, "p:cNvPr")
            .find((el) => (el.getAttribute("name") ?? "").startsWith(`${sabotage.name ?? ""}-`))
        const shape = cNvPr?.parentNode?.parentNode ?? null
        shape?.parentNode?.removeChild(shape)
    }
    return {
        ...mod,
        replaceExisting: (...args: Parameters<typeof mod.replaceExisting>) => {
            const done = mod.replaceExisting(...args)
            drop(args[0])
            return done
        },
        replaceNotes: (...args: Parameters<typeof mod.replaceNotes>) => {
            const count = mod.replaceNotes(...args)
            const t = elements(args[0], "a:t")[0]
            if (sabotage.notes && count === 0 && t !== undefined)
                t.textContent = "manipuliert"
            return count
        },
        applyEdit: async (...args: Parameters<typeof mod.applyEdit>) => {
            const swap = await mod.applyEdit(...args)
            drop(args[3].ownerDocument as Document)
            return swap
        }
    }
})

const Q = [ 1, 2, 3, 4, 5 ].map((i) => path.join(TMP, `text-guard-${i}.png`))
let deck: string

/**  the image a relationship id of slide 0 points to, by content hash  */
const imageOf = async (rid: string | null): Promise<string | null> => {
    const part = (await stateOf(deck)).slides[0]?.part ?? ""
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const rels = parseXml(await (zip.file(relsPartOf(part)) as JSZip.JSZipObject).async("string"))
    const target = relTargets(part, rels).find((r) => r.id === rid)?.target ?? null
    const file = target === null ? null : zip.file(target)
    return file === null ? null : digestBytes(await file.async("nodebuffer"))
}

/**  a table cell of a table shape  */
const cellOf = (table: Element, row: number, col: number): Element =>
    elements(elements(table, "a:tr")[row] as Element, "a:tc")[col] as Element

/**  a picture bullet in a paragraph  */
const bullet = (p: Element, rid: string): void => {
    const doc = p.ownerDocument as Document
    let pPr = elements(p, "a:pPr")[0]
    if (pPr === undefined) {
        pPr = doc.createElementNS(NS_A, "a:pPr")
        p.insertBefore(pPr, p.firstChild)
    }
    for (const old of elements(pPr, "a:buNone"))
        pPr.removeChild(old)
    const buBlip = doc.createElementNS(NS_A, "a:buBlip")
    const blip = doc.createElementNS(NS_A, "a:blip")
    blip.setAttributeNS(NS_R, "r:embed", rid)
    buBlip.appendChild(blip)
    pPr.appendChild(buBlip)
}

/**  the error of an apply, null when it succeeds  */
const errorOf = async (ops: unknown[]): Promise<PptcError | null> => {
    try {
        await executeOps(deck, { ops } as never, OPTS)
        return null
    }
    catch (err) {
        return err as PptcError
    }
}

beforeAll(async () => {
    Q.forEach((file, i) => writeFileSync(file, Buffer.concat([ png(20, 20), Buffer.from(`text-guard-${i}`) ])))
    deck = await freshDeck("text-guard.pptx")
    await apply(deck, [
        { op: "slide.add", layout: "Title Only", ref: "s", placeholders: { title: { text: "Schutz" } } },
        { op: "el.add", slide: "$s", elements: [
            { type: "textbox", name: "Box", frame: { x: 0.5, y: 1.5, w: 4, h: 1 }, text: [ { runs: [ { text: "Punkt eins" } ] }, { runs: [ { text: "Punkt zwei" } ] } ] },
            { type: "table", name: "Tab", frame: { x: 5, y: 1.5, w: 4 },
                data: { rows: [ [ "Kopf", "K2" ], [ "Zelle", "Feld" ] ] } },
            { type: "chart", name: "Diag", frame: { x: 0.5, y: 4, w: 3, h: 2 },
                data: { type: "bar", categories: [ "a", "b" ], series: [ { name: "S", values: [ 1, 2 ] } ] } },
            { type: "image", name: "Bild", frame: { x: 5, y: 4, w: 1, h: 1 }, path: Q[3] as string },
            { type: "textbox", name: "Lauf", frame: { x: 0.5, y: 6.5, w: 4, h: 0.5 },
                text: [ { runs: [ { text: "Hallo " }, { text: "Welt", bold: true } ] } ] },
            ...[ 0, 1, 2, 4 ].map((i) => ({ type: "image", name: `Q${i}`, frame: { x: 7 + i / 2, y: 4, w: 0.5, h: 0.5 }, path: Q[i] })) ] }
    ])

    /*  Q0 becomes a picture bullet of the Box, Q1 a picture fill of cell
        1,0, Q2 a picture bullet of the second paragraph of cell 1,1, Q4
        the picture fill of the run "Welt"  */
    await craftSlide(deck, 0, (xml) => {
        const doc = parseXml(xml)
        const rid = (name: string): string => {
            const pic = shapeNamed(doc, name)
            const id = firstElement(pic, "a:blip")?.getAttributeNS(NS_R, "embed") ?? ""
            pic.parentNode?.removeChild(pic)
            return id
        }
        const [ q0, q1, q2, q4 ] = [ rid("Q0"), rid("Q1"), rid("Q2"), rid("Q4") ]
        bullet(elements(shapeNamed(doc, "Box"), "a:p")[1] as Element, q0)
        const tab = shapeNamed(doc, "Tab")
        const tcPr = firstElement(cellOf(tab, 1, 0), "a:tcPr") as Element
        for (const fill of elements(tcPr, "a:solidFill").filter((f) => f.parentNode === tcPr))
            tcPr.removeChild(fill)
        const blipFill = doc.createElementNS(NS_A, "a:blipFill")
        const blip = doc.createElementNS(NS_A, "a:blip")
        blip.setAttributeNS(NS_R, "r:embed", q1)
        blipFill.appendChild(blip)
        blipFill.appendChild(doc.createElementNS(NS_A, "a:stretch")).appendChild(doc.createElementNS(NS_A, "a:fillRect"))
        tcPr.insertBefore(blipFill, elements(tcPr, "a:extLst")[0] ?? null)
        const body = firstElement(cellOf(tab, 1, 1), "a:txBody") as Element
        const second = (elements(body, "a:p")[0] as Element).cloneNode(true) as Element
        body.appendChild(second)
        bullet(second, q2)
        const rPr = firstElement(elements(shapeNamed(doc, "Lauf"), "a:r")[1] as Element, "a:rPr") as Element
        const runFill = doc.createElementNS(NS_A, "a:blipFill")
        const runBlip = doc.createElementNS(NS_A, "a:blip")
        runBlip.setAttributeNS(NS_R, "r:embed", q4)
        runFill.appendChild(runBlip)
        runFill.appendChild(doc.createElementNS(NS_A, "a:stretch")).appendChild(doc.createElementNS(NS_A, "a:fillRect"))
        rPr.insertBefore(runFill, rPr.firstChild)
        return serializeXml(doc)
    })
})

describe("text edits keep pictures in paragraph and cell properties", () => {
    it("text.replace keeps a picture bullet and a picture cell fill", async () => {
        await apply(deck, [ { op: "text.replace", find: "zwei", replace: "drei" }, { op: "text.replace", find: "Zelle", replace: "Ort" } ])
        const doc = await slideDoc(deck, 0)
        const p = elements(shapeNamed(doc, "Box"), "a:p")[1] as Element
        expect(p.textContent).toBe("Punkt drei")
        expect(await imageOf(firstElement(p, "a:blip")?.getAttributeNS(NS_R, "embed") ?? null)).toBe(digestBytes(readFileSync(Q[0] as string)))
        const cell = cellOf(shapeNamed(doc, "Tab"), 1, 0)
        expect(firstElement(cell, "a:t")?.textContent).toBe("Ort")
        expect(await imageOf(firstElement(cell, "a:blip")?.getAttributeNS(NS_R, "embed") ?? null)).toBe(digestBytes(readFileSync(Q[1] as string)))
    })

    it("a text.replace may take a picture-filled run along that it empties", async () => {
        await apply(deck, [ { op: "text.replace", find: "o Welt", replace: "o Erde" } ])
        const lauf = shapeNamed(await slideDoc(deck, 0), "Lauf")
        expect(lauf.textContent).toBe("Hallo Erde")
        expect(elements(lauf, "a:blipFill")).toHaveLength(0)
    })

    it("el.set cells keeps a picture cell fill", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 1, col: 0, text: "neu" } ] } ])
        const cell = cellOf(shapeNamed(await slideDoc(deck, 0), "Tab"), 1, 0)
        expect(await imageOf(firstElement(cell, "a:blip")?.getAttributeNS(NS_R, "embed") ?? null)).toBe(digestBytes(readFileSync(Q[1] as string)))
    })

    it("a cell text replaced as a whole may take a picture bullet along", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 1, col: 1, text: "eins" } ] } ])
        expect(elements(cellOf(shapeNamed(await slideDoc(deck, 0), "Tab"), 1, 1), "a:buBlip")).toHaveLength(0)
    })
})

describe("content guard after text edits", () => {
    it("fails when a picture or chart vanishes with a text.replace", async () => {
        for (const name of [ "Bild", "Diag" ]) {
            sabotage.name = name
            const err = await errorOf([ { op: "text.replace", find: "Punkt", replace: "Satz" } ])
            sabotage.name = null
            expect(err?.code, `${name} dropped`).toBe("E_INTEGRITY")
            expect(JSON.stringify(err?.details)).toContain(`${name === "Bild" ? "image" : "chart"} would be lost`)
        }
    })

    it("fails when a chart vanishes with an el.set cells", async () => {
        sabotage.name = "Diag"
        const err = await errorOf([ { op: "el.set", slide: "0", name: "Tab", cells: [ { row: 0, col: 1, text: "x" } ] } ])
        sabotage.name = null
        expect(err?.code).toBe("E_INTEGRITY")
        expect(JSON.stringify(err?.details)).toContain("chart would be lost")
    })

    it("blocks a foreign image after a text.replace", async () => {
        const before = await JSZip.loadAsync(readFileSync(deck))
        const out = path.join(TMP, "text-guard-out.pptx")
        await executeOps(deck, { ops: [ { op: "text.replace", find: "Punkt", replace: "Satz" } ] } as never, { ...OPTS, outFile: out })
        const after = await JSZip.loadAsync(readFileSync(out))
        const part = (await stateOf(deck)).slides[0]?.part ?? ""
        const exp: SlideExpect = { source: part, label: "#0", notes: false, assets: "same", textReplaced: true }
        expect(await verifyContent(before, after, [ exp ])).toEqual({ blocked: [], lost: [] })

        /*  tamper: the picture shows another image  */
        after.file("ppt/media/text-guard-foreign.png", Buffer.concat([ png(20, 20), Buffer.from("foreign") ]))
        const rels = await (after.file(relsPartOf(part)) as JSZip.JSZipObject).async("string")
        after.file(relsPartOf(part), rels.replace("</Relationships>", "<Relationship Id=\"rIdForeign\" "
            + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
            + "Target=\"../media/text-guard-foreign.png\"/></Relationships>"))
        const slide = parseXml(await (after.file(part) as JSZip.JSZipObject).async("string"))
        firstElement(shapeNamed(slide, "Bild"), "a:blip")?.setAttributeNS(NS_R, "r:embed", "rIdForeign")
        after.file(part, serializeXml(slide))
        const tampered = await JSZip.loadAsync(await after.generateAsync({ type: "nodebuffer" }))
        expect((await verifyContent(before, tampered, [ exp ])).blocked.length).toBeGreaterThan(0)
    })
})

/**  a deck with one slide whose layout's slide-number placeholder has a
     picture fill; returns the slide part and the picture bytes  */
const bandDeck = async (own: string): Promise<{ part: string, bandImage: Buffer }> => {
    await apply(own, [ { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Band" } }, slideNumber: false } ])

    /*  the layout's slide number placeholder gets a picture fill  */
    const part = (await stateOf(own)).slides[0]?.part ?? ""
    const zip = await JSZip.loadAsync(readFileSync(own))
    const slideRels = parseXml(await (zip.file(relsPartOf(part)) as JSZip.JSZipObject).async("string"))
    const layout = relTargets(part, slideRels).find((r) => r.kind === "slideLayout")?.target ?? ""
    const bandImage = Buffer.concat([ png(30, 30), Buffer.from("band") ])
    zip.file("ppt/media/text-guard-band.png", bandImage)
    const layoutRels = await (zip.file(relsPartOf(layout)) as JSZip.JSZipObject).async("string")
    zip.file(relsPartOf(layout), layoutRels.replace("</Relationships>", "<Relationship Id=\"rIdBand\" "
        + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
        + "Target=\"../media/text-guard-band.png\"/></Relationships>"))
    const doc = parseXml(await (zip.file(layout) as JSZip.JSZipObject).async("string"))
    const sp = elements(doc, "p:sp").find((s) => firstElement(s, "p:ph")?.getAttribute("type") === "sldNum") as Element
    const spPr = firstElement(sp, "p:spPr") as Element
    const blipFill = doc.createElementNS(NS_A, "a:blipFill")
    const blip = doc.createElementNS(NS_A, "a:blip")
    blip.setAttributeNS(NS_R, "r:embed", "rIdBand")
    blipFill.appendChild(blip)
    blipFill.appendChild(doc.createElementNS(NS_A, "a:stretch")).appendChild(doc.createElementNS(NS_A, "a:fillRect"))
    spPr.insertBefore(blipFill, elements(spPr, "a:ln")[0] ?? elements(spPr, "a:extLst")[0] ?? null)
    zip.file(layout, serializeXml(doc))
    const types = await (zip.file("[Content_Types].xml") as JSZip.JSZipObject).async("string")
    if (!types.includes("Extension=\"png\""))
        zip.file("[Content_Types].xml", types.replace("<Default ", "<Default Extension=\"png\" ContentType=\"image/png\"/><Default "))
    writeFileSync(own, await zip.generateAsync({ type: "nodebuffer" }))
    return { part, bandImage }
}

describe("content guard of layout content adopted with the slide number", () => {
    it("accepts exactly the adopted image", async () => {
        const own = await freshDeck("text-guard-band.pptx")
        const { part, bandImage } = await bandDeck(own)

        const before = await JSZip.loadAsync(readFileSync(own))
        const out = path.join(TMP, "text-guard-band-out.pptx")
        await executeOps(own, { ops: [ { op: "slide.fill", slide: "0", slideNumber: true } ] } as never, { ...OPTS, outFile: out })
        const after = await JSZip.loadAsync(readFileSync(out))
        const base: SlideExpect = { source: part, label: "#0", notes: false, assets: "same" }
        const right = digestBytes(bandImage)
        expect(await verifyContent(before, after, [ { ...base, replaced: [ { from: [], to: right } ] } ])).toEqual({ blocked: [], lost: [] })
        expect((await verifyContent(before, after, [ base ])).blocked.length).toBeGreaterThan(0)
        const wrong = await verifyContent(before, after, [ { ...base, replaced: [ { from: [], to: digestBytes(Buffer.from("other")) } ] } ])
        expect(wrong.blocked.join(" ")).toContain("is missing")
    })

    it("shows the slide number again and removes it with its image (no false E_INTEGRITY)", async () => {
        const own = await freshDeck("text-guard-band-again.pptx")
        const { bandImage } = await bandDeck(own)
        const images = async (): Promise<string[]> => {
            const zip = await JSZip.loadAsync(readFileSync(own))
            const slide = (await stateOf(own)).slides[0]?.part ?? ""
            const rels = parseXml(await (zip.file(relsPartOf(slide)) as JSZip.JSZipObject).async("string"))
            const out: string[] = []
            for (const rel of relTargets(slide, rels).filter((r) => r.kind === "image"))
                out.push(digestBytes(await (zip.file(rel.target as string) as JSZip.JSZipObject).async("nodebuffer")))
            return out
        }
        const band = digestBytes(bandImage)
        await apply(own, [ { op: "slide.fill", slide: "0", slideNumber: true } ])
        expect(await images()).toEqual([ band ])

        /*  again: the placeholder is cloned anew, its image stays  */
        await apply(own, [ { op: "slide.fill", slide: "0", slideNumber: true } ])
        expect(await images()).toEqual([ band ])

        /*  removed: its image goes with it  */
        await apply(own, [ { op: "slide.fill", slide: "0", slideNumber: false } ])
        expect(await images()).toEqual([])
        expect((await stateOf(own)).slides[0]?.shapes.some((sh) => sh.name.startsWith("Slide Number"))).toBe(false)
    })

    it("still blocks a lost image the band op did not remove", async () => {
        const own = await freshDeck("text-guard-band-strict.pptx")
        const { part, bandImage } = await bandDeck(own)
        await apply(own, [ { op: "slide.fill", slide: "0", slideNumber: true } ])
        const before = await JSZip.loadAsync(readFileSync(own))
        const out = path.join(TMP, "text-guard-band-strict-out.pptx")
        await executeOps(own, { ops: [ { op: "slide.fill", slide: "0", slideNumber: false } ] } as never, { ...OPTS, outFile: out })
        const after = await JSZip.loadAsync(readFileSync(out))
        const base: SlideExpect = { source: part, label: "#0", notes: false, assets: "same" }
        expect((await verifyContent(before, after, [ base ])).lost.map((l) => l.what)).toEqual([ "image" ])
        expect(await verifyContent(before, after, [ { ...base, replaced: [ { from: [ digestBytes(bandImage) ], to: null } ] } ]))
            .toEqual({ blocked: [], lost: [] })
    })
})

describe("content guard of speaker notes", () => {
    it("exempts the notes only of slides where a text.replace replaced in them", async () => {
        const notes = await freshDeck("text-guard-notes.pptx")
        await apply(notes, [
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Eins" } }, notes: "Sprechertext" },
            { op: "slide.add", layout: "CONTENT", placeholders: { title: { text: "Zwei" } }, notes: "Andere Notiz" } ])
        /*  the second slide is touched, its notes hold no match  */
        const ops = [ { op: "slide.fill", slide: "1", placeholders: { title: { text: "Zwei neu" } } },
            { op: "text.replace", find: "Sprecher", replace: "Redner", notes: true } ]
        sabotage.notes = true
        try {
            expect(await failure(notes, ops)).toBe("E_INTEGRITY")
        }
        finally {
            sabotage.notes = false
        }
        await apply(notes, ops)
        expect((await stateOf(notes)).slides.map((s) => s.notes)).toEqual([ "Rednertext", "Andere Notiz" ])
    })
})
