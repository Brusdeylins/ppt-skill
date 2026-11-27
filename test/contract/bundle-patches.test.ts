/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Contract: the pptx-automizer patches (engine/automizer-patches.ts) are
**  effective in the BUILT bundle, not only in the sources -- the reported
**  regression (images added in separate applies overwriting each other)
**  replayed through dst/pptc.mjs. `npm test` builds first.
*/

import { describe, expect, it } from "vitest"
import { spawnSync } from "node:child_process"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"

const here = path.dirname(fileURLToPath(import.meta.url))
const BIN = path.join(here, "..", "..", "dst", "pptc.mjs")
const TEMPLATE = path.join(here, "..", "fixtures", "neutral-template.pptx")
const TMP = path.join(here, "..", "tmp")

/**  run the CLI bundle and parse its envelope  */
const run = (...argv: string[]): Record<string, unknown> =>
    JSON.parse(spawnSync("node", [ BIN, ...argv ], { encoding: "utf8" }).stdout) as Record<string, unknown>

describe("automizer patches in the built bundle", () => {
    it("keeps three SVG images added in three applies apart", async () => {
        mkdirSync(TMP, { recursive: true })
        const deck = path.join(TMP, "bundle-patches.pptx")
        expect(run("new", deck, "--template", TEMPLATE, "--force").ok).toBe(true)
        const colors = [ "red", "green", "blue" ]
        for (const c of colors) {
            const svg = path.join(TMP, `bundle-patches-${c}.svg`)
            writeFileSync(svg, `<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="${c}"/></svg>`)
            const op = { op: "el.add", slide: "$s", elements: [ { type: "image", name: "B", frame: { x: 1, y: 1, w: 2, h: 2 }, path: svg } ] }
            const res = spawnSync("node", [ BIN, "apply", deck, "--ops", "-" ], { encoding: "utf8",
                input: JSON.stringify({ ops: [ { op: "slide.add", ref: "s", layout: "CONTENT" }, op ] }) })
            expect((JSON.parse(res.stdout) as { ok: boolean }).ok).toBe(true)
        }
        const zip = await JSZip.loadAsync(readFileSync(deck))
        const fills = await Promise.all(Object.keys(zip.files).filter((f) => /^ppt\/media\/.+\.svg$/.test(f))
            .map(async (f) => /fill="(\w+)"/.exec(await (zip.file(f) as JSZip.JSZipObject).async("string"))?.[1]))
        expect(fills.sort()).toEqual([ ...colors ].sort())
    })
})
