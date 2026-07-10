/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  Contract: documentation stays in sync with the code. Every dispatched
**  command must be covered by the detailed help registry, the usage text
**  and the README; every op and every lint warning code must be
**  documented. Doc drift is a failing test, not a hope.
*/

import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { helpFor } from "../../src/cli/help.js"
import { OP_NAMES } from "../../src/schema/ops.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const root = path.join(here, "..", "..")
const mainSrc = readFileSync(path.join(root, "src", "cli", "main.ts"), "utf8")
const lintSrc = readFileSync(path.join(root, "src", "core", "lint.ts"), "utf8")
const opsSrc  = readFileSync(path.join(root, "src", "schema", "ops.ts"), "utf8")
const readme = readFileSync(path.join(root, "README.md"), "utf8")

/**  command names parsed from the dispatch tables in main.ts  */
const commands = (): string[] => {
    const top = [...mainSrc.matchAll(/^ {4}"([a-z-]+)": /gm)]
        .map((m) => m[1] as string)
        .filter((c) => c !== "--version")
    const tplBlock = /TPL_COMMANDS[^}]*}/s.exec(mainSrc)?.[0] ?? ""
    const tpl = [...tplBlock.matchAll(/"([a-z]+)":/g)].map((m) => `tpl ${m[1]}`)
    return [...top.filter((c) => !["list", "describe", "inspect", "validate"].includes(c)), ...tpl]
}

/**  warning codes parsed from the LintWarning union in lint.ts  */
const warningCodes = (): string[] =>
    [...lintSrc.matchAll(/"(W_[A-Z_]+)"/g)].map((m) => m[1] as string)

/**  field names parsed from the shared fill payload in schema/ops.ts  */
const fillFields = (): string[] => {
    const block = /const fillProps = \{(.*?)\n\}/s.exec(opsSrc)?.[1] ?? ""
    return [...block.matchAll(/^ {4}(\w+):/gm)].map((m) => m[1] as string)
}

describe("documentation sync contract", () => {
    it("every command has a detailed help entry", () => {
        for (const cmd of commands())
            expect(helpFor(cmd.split(" ")), cmd).not.toBeNull()
    })

    it("the usage text mentions every command", () => {
        const usage = /const USAGE = `([^`]*)`/s.exec(mainSrc)?.[1] ?? ""
        for (const cmd of commands())
            expect(usage, cmd).toContain(cmd.split(" ").pop() as string)
    })

    it("the README documents every command", () => {
        for (const cmd of commands())
            expect(readme, cmd).toContain(`pptc ${cmd}`)
    })

    it("the README and 'help ops' document every op", () => {
        const opsHelp = helpFor(["ops"]) ?? ""
        for (const op of OP_NAMES) {
            expect(readme, op).toContain(op)
            expect(opsHelp, op).toContain(op)
        }
    })

    it("the README and 'help ops' document every fill payload field", () => {
        const opsHelp = helpFor(["ops"]) ?? ""
        const fields = fillFields()
        expect(fields.length).toBeGreaterThanOrEqual(5)
        for (const field of fields) {
            expect(readme, field).toContain(field)
            expect(opsHelp, field).toContain(field)
        }
    })

    it("the README documents the skill plugin (install + SKILL.md link)", () => {
        expect(readme).toContain("/plugin marketplace add Brusdeylins/ppt-skill")
        expect(readme).toContain("/plugin install ppt@ppt-skill")
        expect(readme).toContain("plugin/skills/ppt/SKILL.md")
        const pluginReadme = readFileSync(path.join(root, "plugin", "README.md"), "utf8")
        expect(pluginReadme).toContain("/plugin install ppt@ppt-skill")
    })

    it("both skills import the shared plugin-level control.md", () => {
        const shared = readFileSync(path.join(root, "plugin", "meta", "control.md"), "utf8")
        expect(shared).toContain("Control Tags")
        for (const skill of ["ppt", "ppt-prepare"]) {
            const src = readFileSync(path.join(root, "plugin", "skills", skill, "SKILL.md"), "utf8")
            expect(src, skill).toContain("@${CLAUDE_SKILL_DIR}/../../meta/control.md")
        }
    })

    it("the README and the apply help document every lint warning code", () => {
        const codes = warningCodes()
        expect(codes.length).toBeGreaterThanOrEqual(2)
        const applyHelp = helpFor(["apply"]) ?? ""
        for (const code of codes) {
            expect(readme, code).toContain(code)
            expect(applyHelp, code).toContain(code)
        }
    })
})
