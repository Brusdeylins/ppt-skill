/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  skill-zip.mjs: package the skills for a marketplace owner into ONE ZIP in
**  the plugin layout (skills/ppt, skills/ppt-prepare, meta/control.md), so the
**  skills' relative `../../meta/control.md` import keeps working unchanged.
**  Output: deploy/ppt-skills-<version>.zip
**
**  Also one stand-alone ZIP per skill (deploy/<skill>.zip, skill folder at
**  the ZIP root) for upload into the Claude app: there is no ../../meta, so
**  the `@...control.md` import line in SKILL.md is replaced by the inlined
**  content of meta/control.md.
*/

import { readdirSync, statSync, readFileSync, mkdirSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..")
const pluginDir = path.join(root, "plugin")
const version = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8")).version
const outDir = path.join(root, "deploy")

/*  collect every file below dir (skipping macOS cruft)  */
const walk = (dir) => {
    const out = []
    for (const entry of readdirSync(dir)) {
        const p = path.join(dir, entry)
        if (statSync(p).isDirectory())
            out.push(...walk(p))
        else if (entry !== ".DS_Store")
            out.push(p)
    }
    return out
}

const zip = new JSZip()
for (const file of [ ...walk(path.join(pluginDir, "skills")), ...walk(path.join(pluginDir, "meta")) ])
    zip.file(path.relative(pluginDir, file).split(path.sep).join("/"), readFileSync(file))

mkdirSync(outDir, { recursive: true })

/*  stand-alone per-skill ZIPs with control.md inlined into SKILL.md  */
const importLine = /^@\$\{CLAUDE_SKILL_DIR\}\/\.\.\/\.\.\/meta\/control\.md[ \t]*$/m
const control = readFileSync(path.join(pluginDir, "meta", "control.md"), "utf8").trim()
for (const name of readdirSync(path.join(pluginDir, "skills"))) {
    const skillPath = path.join(pluginDir, "skills", name)
    if (!statSync(skillPath).isDirectory())
        continue
    const skillZip = new JSZip()
    for (const file of walk(skillPath)) {
        const rel = path.relative(skillPath, file).split(path.sep).join("/")
        let data = readFileSync(file)
        if (rel === "SKILL.md") {
            const text = data.toString("utf8")
            if (!importLine.test(text))
                throw new Error(`${name}/SKILL.md: control.md import line not found`)
            data = Buffer.from(text.replace(importLine, () => control))
        }
        skillZip.file(`${name}/${rel}`, data)
    }
    const skillFile = path.join(outDir, `${name}.zip`)
    const skillBuf = await skillZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
    writeFileSync(skillFile, skillBuf)
    console.log(`${path.relative(root, skillFile)}  (${(skillBuf.length / 1024).toFixed(0)} KB)`)
}

const outFile = path.join(outDir, `ppt-skills-${version}.zip`)
const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
writeFileSync(outFile, buf)
console.log(`${path.relative(root, outFile)}  (${(buf.length / 1024).toFixed(0)} KB)`)
