/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the byte guard of an apply.
*/

import { describe, expect, it } from "vitest"
import JSZip from "jszip"
import { byteGuard } from "../../src/engine/guard.js"

/**  a zip reloaded from its bytes  */
const reload = async (zip: JSZip): Promise<JSZip> =>
    await JSZip.loadAsync(await zip.generateAsync({ type: "nodebuffer" }))

describe("byte guard", () => {
    it("fires when an untouched part changes or is rewritten", async () => {
        const src = new JSZip()
        src.file("ppt/slides/slide1.xml", "<a/>")
        const pkg = await reload(src)
        const untouched = new Map([ [ "ppt/slides/slide1.xml", pkg.file("ppt/slides/slide1.xml") ] ])
        expect(await byteGuard(pkg, untouched, await reload(pkg))).toEqual([])

        const other = new JSZip()
        other.file("ppt/slides/slide1.xml", "<b/>")
        expect(await byteGuard(pkg, untouched, await reload(other))).toHaveLength(1)

        pkg.file("ppt/slides/slide1.xml", "<a/>")
        expect(await byteGuard(pkg, untouched, await reload(pkg))).toHaveLength(1)
    })
})
