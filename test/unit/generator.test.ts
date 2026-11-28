/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Unit: the PptxGenJS import resolves to a usable generator class
**  (engine/inplace/transplant).
*/

import { describe, expect, it } from "vitest"
import PptxGenJsPkg from "pptxgenjs"
import { PptxGenJS, unwrapDefault } from "../../src/engine/inplace/transplant.js"

describe("PptxGenJS import", () => {
    it("is a constructor whose instances can add slides and write", () => {
        expect(typeof PptxGenJS).toBe("function")
        const gen = new PptxGenJS() as unknown as Record<string, unknown>
        expect(typeof gen["addSlide"]).toBe("function")
        expect(typeof gen["write"]).toBe("function")
    })

    it("loads the ESM build, whose default export is the class itself", () => {
        expect(typeof PptxGenJsPkg).toBe("function")
        expect(PptxGenJS).toBe(PptxGenJsPkg)
    })

    it("unwraps CommonJS interop default levels", () => {
        const Gen = function Gen (): void { /*  a stand-in constructor  */ }
        expect(unwrapDefault(Gen)).toBe(Gen)
        expect(unwrapDefault({ default: Gen })).toBe(Gen)
        expect(unwrapDefault({ default: { default: Gen } })).toBe(Gen)
    })
})
