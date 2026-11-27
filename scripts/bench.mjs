/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  bench.mjs: wall-clock time of a trivial apply (one notes fill) through the
**  built CLI (dst/pptc.mjs), median of several runs, on the neutral fixture
**  and -- when present locally -- on the demo deck. Run `npm run build` first.
**  Usage: node scripts/bench.mjs [runs]
*/

import { copyFileSync, existsSync, mkdtempSync, rmSync } from "node:fs"
import { spawnSync } from "node:child_process"
import { tmpdir } from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..")
const cli = path.join(root, "dst/pptc.mjs")
const runs = Number(process.argv[2] ?? "5")
const decks = [
    [ "neutral fixture", path.join(root, "test/fixtures/neutral-template.pptx") ],
    [ "demo deck", path.join(root, "demo/kuehlschrank-funktionsweise.pptx") ]
]
const op = JSON.stringify({ op: "slide.fill", slide: "index:0", notes: "bench" })

const dir = mkdtempSync(path.join(tmpdir(), "pptc-bench-"))
try {
    for (const [ label, deck ] of decks) {
        if (!existsSync(deck)) {
            console.log(`${label}: skipped (not present)`)
            continue
        }
        const times = []
        for (let i = 0; i < runs; i++) {
            const work = path.join(dir, `deck-${i}.pptx`)
            copyFileSync(deck, work)
            const start = process.hrtime.bigint()
            const res = spawnSync(process.execPath, [ cli, "apply", work, "-e", op ], { encoding: "utf8" })
            const ms = Number(process.hrtime.bigint() - start) / 1e6
            if (res.status !== 0)
                throw new Error(`${label}: apply failed: ${res.stdout}${res.stderr}`)
            times.push(ms)
        }
        times.sort((a, b) => a - b)
        const median = times[Math.floor(times.length / 2)]
        console.log(`${label}: median ${(median / 1000).toFixed(2)} s over ${runs} runs `
            + `(min ${(times[0] / 1000).toFixed(2)} s, max ${(times[times.length - 1] / 1000).toFixed(2)} s)`)
    }
}
finally {
    rmSync(dir, { recursive: true, force: true })
}
