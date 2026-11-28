/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
*/

import { describe, expect, it } from "vitest"
import { estimateCapacity, estimateUsedLines } from "../../src/core/describe/capacity.js"
import { coverageFraction, describePosition, horizontalBand, nearestAspect, regionWithin, verticalBand } from "../../src/core/describe/position.js"
import { renderMinimap } from "../../src/core/describe/minimap.js"
import { suitabilityHint } from "../../src/core/describe/narrate.js"
import { lintPlaceholderText } from "../../src/core/lint.js"
import type { Layout, Placeholder } from "../../src/core/model.js"

const SLIDE_W = 13.33
const SLIDE_H = 7.5

describe("capacity", () => {
    it("estimates lines and chars from geometry and font size", () => {
        const cap = estimateCapacity({ x: 0, y: 0, w: 12, h: 4.7 }, 18)
        expect(cap.lines).toBeGreaterThanOrEqual(12)
        expect(cap.lines).toBeLessThanOrEqual(18)
        expect(cap.charsPerLine).toBeGreaterThan(60)
    })
    it("accounts for wrapping of long paragraphs", () => {
        const cap = estimateCapacity({ x: 0, y: 0, w: 5, h: 2 }, 18)
        expect(estimateUsedLines(["x".repeat(cap.charsPerLine * 3)], cap)).toBe(3)
        expect(estimateUsedLines(["a", "b", "c"], cap)).toBe(3)
    })
})

describe("position", () => {
    it("classifies horizontal and vertical bands", () => {
        expect(horizontalBand({ x: 0.5, y: 0, w: 12.3, h: 1 }, SLIDE_W)).toBe("full width")
        expect(horizontalBand({ x: 0, y: 0, w: 6, h: 1 }, SLIDE_W)).toBe("left half")
        expect(horizontalBand({ x: 9.5, y: 0, w: 3.5, h: 1 }, SLIDE_W)).toBe("right column")
        expect(verticalBand({ x: 0, y: 0.2, w: 1, h: 1 }, SLIDE_H)).toBe("upper area")
        expect(verticalBand({ x: 0, y: 0, w: 1, h: 7 }, SLIDE_H)).toBe("full height")
    })
    it("includes the area share", () => {
        expect(describePosition({ x: 0, y: 0, w: 6.67, h: 7.5 }, SLIDE_W, SLIDE_H))
            .toMatch(/left half.*50% of slide/)
    })
    it("snaps to the nearest common aspect ratio", () => {
        expect(nearestAspect({ x: 0, y: 0, w: 16, h: 9 })).toBe("16:9")
        expect(nearestAspect({ x: 0, y: 0, w: 4, h: 3.1 })).toBe("4:3")
        expect(nearestAspect({ x: 0, y: 0, w: 3, h: 3 })).toBe("1:1")
    })
})

const ph = (idx: number, kind: Placeholder["kind"], frame: Placeholder["frame"]): Placeholder => ({
    idx, kind, name: `ph${idx}`, frame,
    capacity: frame !== null && kind !== "picture" ? estimateCapacity(frame, 18) : null
})

describe("minimap and suitability", () => {
    const layout: Layout = {
        index: 9, name: "Bild links", reserved: [],
        placeholders: [
            ph(0, "title", { x: 4.9, y: 0.6, w: 5, h: 0.8 }),
            ph(14, "picture", { x: 0, y: 0, w: 4.7, h: 7.5 }),
            ph(13, "body", { x: 5.4, y: 1.9, w: 7.5, h: 4.7 })
        ]
    }
    it("renders labeled boxes", () => {
        const map = renderMinimap(layout, SLIDE_W, SLIDE_H)
        expect(map).toContain("IMG:14")
        expect(map).toContain("TXT:13")
        expect(map.split("\n")).toHaveLength(13)
    })
    it("derives a suitability hint from structure", () => {
        expect(suitabilityHint(layout)).toMatch(/picture/)
        expect(suitabilityHint({ index: 0, name: "x", reserved: [], placeholders: [ph(0, "title", null)] }))
            .toMatch(/key message/)
    })
})

describe("lint", () => {
    it("warns when planned text exceeds the capacity", () => {
        const target = ph(12, "body", { x: 0, y: 0, w: 5, h: 0.4 })
        const warning = lintPlaceholderText(
            "Dieser Text ist deutlich zu lang für einen winzigen einzeiligen Platzhalter und muss daher eine Overflow-Warnung erzeugen",
            target, { id: 1, index: 0, title: null })
        expect(warning?.code).toBe("W_TEXT_OVERFLOW")
        expect(warning?.estimatedLines).toBeGreaterThan(warning?.maxLines as number)
    })
    it("counts a vertical tab (line break) as a line of its own", () => {
        const target = ph(13, "body", { x: 0, y: 0, w: 12, h: 1 })
        const lines = (target.capacity as { lines: number }).lines
        const text = Array.from({ length: lines + 1 }, (_, i) => `Zeile ${i}`)
        expect(lintPlaceholderText(text.slice(0, lines).join("\v"), target, { id: 1, index: 0, title: null })).toBeNull()
        expect(lintPlaceholderText(text.join("\v"), target, { id: 1, index: 0, title: null })?.code).toBe("W_TEXT_OVERFLOW")
        expect(lintPlaceholderText([ { runs: [ { text: text.join("\v") } ] } ], target, { id: 1, index: 0, title: null })?.code)
            .toBe("W_TEXT_OVERFLOW")
    })
    it("stays silent when the text fits", () => {
        const target = ph(13, "body", { x: 0, y: 0, w: 12, h: 5 })
        expect(lintPlaceholderText("kurz", target, { id: 1, index: 0, title: null })).toBeNull()
    })
})

describe("regionWithin (picture overlay regions)", () => {
    const pic = { x: 0, y: 0, w: 13.33, h: 7.5 }
    it("classifies bands relative to the picture", () => {
        expect(regionWithin(pic, { x: 0.7, y: 5.6, w: 6, h: 1.4 })).toBe("bottom area, left part")
        expect(regionWithin(pic, { x: 0, y: 0, w: 13.33, h: 1.2 })).toBe("top area")
        expect(regionWithin(pic, { x: 0, y: 0, w: 13.33, h: 7.5 })).toBe("entire image")
        expect(regionWithin(pic, { x: 5, y: 3, w: 3, h: 1.5 })).toBe("middle area, center")
    })
    it("returns null for disjoint boxes", () => {
        expect(regionWithin({ x: 0, y: 0, w: 5, h: 7.5 }, { x: 6, y: 1, w: 4, h: 1 })).toBeNull()
    })
})

describe("coverageFraction (background-image detection)", () => {
    const pic = { x: 0, y: 0, w: 10, h: 10 }
    it("is 0 with no overlaps and 1 when fully covered", () => {
        expect(coverageFraction(pic, [])).toBe(0)
        expect(coverageFraction(pic, [{ x: 0, y: 0, w: 10, h: 10 }])).toBeCloseTo(1)
    })
    it("sums a partial overlap", () => {
        expect(coverageFraction(pic, [{ x: 0, y: 5, w: 10, h: 5 }])).toBeCloseTo(0.5)
    })
    it("does not double-count overlapping inners (union, not sum)", () => {
        const a = { x: 0, y: 0, w: 6, h: 10 }
        const b = { x: 4, y: 0, w: 6, h: 10 }
        expect(coverageFraction(pic, [a, b])).toBeCloseTo(1)
    })
    it("clips inners to the picture box", () => {
        expect(coverageFraction(pic, [{ x: -5, y: 0, w: 10, h: 10 }])).toBeCloseTo(0.5)
    })
    it("treats a title+footer band as partial, a full body box as background", () => {
        const slide = { x: 0, y: 0, w: 13.33, h: 7.5 }
        const titleFooter = coverageFraction(slide, [
            { x: 0, y: 0, w: 13.33, h: 1.2 },
            { x: 0, y: 7, w: 13.33, h: 0.5 }
        ])
        expect(titleFooter).toBeLessThan(0.65)
        const fullBody = coverageFraction(slide, [{ x: 0, y: 1, w: 13.33, h: 6.5 }])
        expect(fullBody).toBeGreaterThanOrEqual(0.65)
    })
})
