/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Contract: pptc and its skills carry no update check, version banner or
**  self-update -- external calls and process spawns from a skill trip the
**  Claude app security scan. Guards the CLI surface, the skill text and the
**  shipped bundle against reintroducing them.
*/

import { describe, expect, it } from "vitest"
import { spawnSync } from "node:child_process"
import { readFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..")
const bundle = path.join(root, "dst", "pptc.mjs")
const pkg = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8")) as { version: string }

const run = (...args: string[]) => {
    const res = spawnSync("node", [ bundle, ...args ], { encoding: "utf8" })
    return { status: res.status, out: JSON.parse(res.stdout) as Record<string, unknown> }
}

describe("no update machinery", () => {
    it("`--version` returns the plain envelope without an update field", () => {
        const { status, out } = run("--version")
        expect(status).toBe(0)
        expect(out).toEqual({ ok: true, cmd: "--version", result: { version: pkg.version } })
    })

    it("`update` is an unknown command", () => {
        const { status, out } = run("update")
        expect(status).toBe(2)
        expect(out["ok"]).toBe(false)
        expect(JSON.stringify(out)).toContain("unknown command 'update'")
    })

    it("the skills run no inline node call and no network call", () => {
        for (const skill of [ "ppt", "ppt-prepare" ]) {
            const md = readFileSync(path.join(root, "plugin", "skills", skill, "SKILL.md"), "utf8")
            for (const needle of [ "node -e", "fetch(", "api.github.com", "curl ", "wget " ])
                expect(md, `${skill}/SKILL.md contains '${needle}'`).not.toContain(needle)
        }
    })

    it("the bundle contains no process spawn or registry call of its own", () => {
        const js = readFileSync(path.join(root, "plugin", "skills", "ppt", "scripts", "pptc.mjs"), "utf8")
        for (const needle of [ "child_process", "registry.npmjs.org", "spawnSync", "npm install" ])
            expect(js, `bundle contains '${needle}'`).not.toContain(needle)
    })
})
