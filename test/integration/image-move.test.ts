/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Integration: the content guard counts el.set image swaps as a
**  multiset -- the new images must come, the old ones may go. An image
**  moved from one picture to another passes in either op order, two
**  pictures exchanging their images pass, and a swap the engine did not
**  carry out fails the apply although the slide's other text changed
**  (which lets an image go otherwise).
*/

import { beforeEach, describe, expect, it, vi } from "vitest"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import JSZip from "jszip"
import type { Element } from "@xmldom/xmldom"
import { digestBytes } from "../../src/engine/guard.js"
import { NS_R, firstElement } from "../../src/engine/xml.js"
import { TMP, apply, failure, freshDeck, png, stateOf } from "../util/elements.js"

/*  sabotage: the swap of the picture of this name is undone after the
    edit, which still reports it -- as an engine bug would  */
const sabotage = vi.hoisted(() => ({ name: null as string | null }))
vi.mock("../../src/engine/inplace/edits.js", async (importOriginal) => {
    const mod = await importOriginal<typeof import("../../src/engine/inplace/edits.js")>()
    return {
        ...mod,
        applyEdit: async (...args: Parameters<typeof mod.applyEdit>) => {
            const shape = args[3] as Element
            const blip = firstElement(shape, "a:blip")
            const before = blip?.getAttributeNS(NS_R, "embed") ?? null
            const swap = await mod.applyEdit(...args)
            const name = firstElement(shape, "p:cNvPr")?.getAttribute("name") ?? ""
            if (sabotage.name !== null && name.startsWith(`${sabotage.name}-`) && blip !== null && before !== null)
                blip.setAttributeNS(NS_R, "r:embed", before)
            return swap
        }
    }
})

const [ ONE, TWO, THREE ] = [ 1, 2, 3 ].map((i) => path.join(TMP, `image-move-${i}.png`)) as [ string, string, string ]
let deck: string

/**  the images slide 0 shows per picture, by content hash  */
const shown = async (): Promise<Record<string, string>> => {
    const zip = await JSZip.loadAsync(readFileSync(deck))
    const out: Record<string, string> = {}
    for (const sh of (await stateOf(deck)).slides[0]?.shapes ?? [])
        if (sh.type === "picture")
            out[sh.name.replace(/-[0-9a-f-]{36}$/, "")] = digestBytes(await (zip.file(`ppt/media/${sh.image ?? ""}`) as JSZip.JSZipObject).async("nodebuffer"))
    return out
}

const hash = (file: string): string => digestBytes(readFileSync(file))

beforeEach(async () => {
    sabotage.name = null
    for (const [ i, file ] of [ ONE, TWO, THREE ].entries())
        writeFileSync(file, Buffer.concat([ png(40, 40), Buffer.from(`image-move-${i}`) ]))
    deck = await freshDeck("image-move.pptx")
    await apply(deck, [ { op: "slide.add", layout: "Title Only", placeholders: { title: { text: "Bilder" } } },
        { op: "el.add", slide: "0", elements: [
            { type: "image", name: "Links", frame: { x: 1, y: 2, w: 2, h: 2 }, path: ONE },
            { type: "image", name: "Rechts", frame: { x: 5, y: 2, w: 2, h: 2 }, path: TWO },
            { type: "textbox", name: "Text", frame: { x: 1, y: 5, w: 4, h: 1 }, text: "Bildtext" } ] } ])
})

describe("image swaps as a multiset", () => {
    it("moves an image to the other picture in either op order", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Rechts", image: ONE },
            { op: "el.set", slide: "0", name: "Links", image: THREE } ])
        expect(await shown()).toEqual({ Links: hash(THREE), Rechts: hash(ONE) })
        await apply(deck, [ { op: "el.set", slide: "0", name: "Rechts", image: TWO },
            { op: "el.set", slide: "0", name: "Links", image: THREE } ])
        await apply(deck, [ { op: "el.set", slide: "0", name: "Links", image: ONE },
            { op: "el.set", slide: "0", name: "Rechts", image: THREE } ])
        expect(await shown()).toEqual({ Links: hash(ONE), Rechts: hash(THREE) })
    })

    it("exchanges the images of two pictures", async () => {
        await apply(deck, [ { op: "el.set", slide: "0", name: "Links", image: TWO },
            { op: "el.set", slide: "0", name: "Rechts", image: ONE } ])
        expect(await shown()).toEqual({ Links: hash(TWO), Rechts: hash(ONE) })
    })

    it("fails an exchange the engine carried out only half", async () => {
        sabotage.name = "Links"
        expect(await failure(deck, [ { op: "el.set", slide: "0", name: "Text", text: "neu" },
            { op: "el.set", slide: "0", name: "Links", image: TWO },
            { op: "el.set", slide: "0", name: "Rechts", image: ONE } ])).toBe("E_INTEGRITY")
        expect(await shown()).toEqual({ Links: hash(ONE), Rechts: hash(TWO) })
    })
})
