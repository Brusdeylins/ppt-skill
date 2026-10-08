/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  skill-zip.mjs: package the skills into ZIPs under deploy/ (gitignored).
**
**    ppt-skills-<version>.zip   plugin layout (skills/ppt, skills/ppt-prepare,
**                               meta/control.md) for the marketplace owner, so
**                               the relative `../../meta/control.md` import
**                               keeps working unchanged
**    ppt.zip, ppt-prepare.zip   stand-alone skill folders for upload into the
**                               Claude app; there is no ../../meta, so the
**                               `@...control.md` import line in SKILL.md is
**                               replaced by the inlined meta/control.md
**
**  If ./private-templates exists (or --from <dir> is given), the msg variants
**  are built as well: ppt-msg.zip, ppt-prepare-msg.zip and
**  ppt-skills-msg-<version>.zip. The msg
**  templates (.potx/.pptx + optional <name>.md sidecars) REPLACE the neutral
**  default in the ppt skill's assets/ (neutral-template.pptx and its sidecar are dropped).
**  They are never part of the git tree or a GitHub release, which always
**  carries the neutral template only.
*/

import { readdirSync, statSync, readFileSync, mkdirSync, writeFileSync, existsSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..")
const pluginDir = path.join(root, "plugin")
const skillsDir = path.join(pluginDir, "skills")
const version = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8")).version
const outDir = path.join(root, "deploy")

const args = process.argv.slice(2)
const fromIdx = args.indexOf("--from")
const fromDir = fromIdx !== -1 ? args[fromIdx + 1] : null
const templatesSrc = fromDir ? path.resolve(fromDir) : path.join(root, "private-templates")

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

const importLine = /^@\$\{CLAUDE_SKILL_DIR\}\/\.\.\/\.\.\/meta\/control\.md[ \t]*$/m
const control = readFileSync(path.join(pluginDir, "meta", "control.md"), "utf8").trim()

/*  files of one skill as [ relative path, data ]; msg templates replace the neutral default  */
const skillFiles = (name, msgFiles, inlineControl) => {
    const skillPath = path.join(skillsDir, name)
    const hasAssets = existsSync(path.join(skillPath, "assets"))
    const out = []
    for (const file of walk(skillPath)) {
        const rel = path.relative(skillPath, file).split(path.sep).join("/")
        if (msgFiles.length > 0 && hasAssets && /^neutral-template\.(pptx|md)$/.test(path.basename(file)))
            continue
        let data = readFileSync(file)
        if (inlineControl && rel === "SKILL.md") {
            const text = data.toString("utf8")
            if (!importLine.test(text))
                throw new Error(`${name}/SKILL.md: control.md import line not found`)
            data = Buffer.from(text.replace(importLine, () => control))
        }
        out.push([ rel, data ])
    }
    if (hasAssets)
        for (const file of msgFiles)
            out.push([ `assets/${path.basename(file)}`, readFileSync(file) ])
    return out
}

const write = async (zip, file, note = "") => {
    const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
    writeFileSync(path.join(outDir, file), buf)
    console.log(`deploy/${file}  (${(buf.length / 1024).toFixed(0)} KB${note})`)
}

mkdirSync(outDir, { recursive: true })
const names = readdirSync(skillsDir).filter((n) => statSync(path.join(skillsDir, n)).isDirectory())

/*  variant "" = neutral template (public), "msg" = msg templates  */
const variants = [ { tag: "", msgFiles: [] } ]
if (existsSync(templatesSrc)) {
    const msgFiles = walk(templatesSrc).filter((f) => /\.(potx|pptx|md)$/i.test(f))
    if (!msgFiles.some((f) => /\.(potx|pptx)$/i.test(f)))
        throw new Error(`no .potx/.pptx templates in ${templatesSrc}`)
    variants.push({ tag: "msg", msgFiles })
}
else if (fromDir)
    throw new Error(`template source not found: ${templatesSrc}`)

for (const { tag, msgFiles } of variants) {
    const note = msgFiles.length > 0 ? `, ${msgFiles.length} msg template file(s)` : ""

    /*  marketplace owner: plugin layout  */
    const plugin = new JSZip()
    for (const name of names)
        for (const [ rel, data ] of skillFiles(name, msgFiles, false))
            plugin.file(`skills/${name}/${rel}`, data)
    plugin.file("meta/control.md", readFileSync(path.join(pluginDir, "meta", "control.md")))
    await write(plugin, `ppt-skills${tag ? `-${tag}` : ""}-${version}.zip`, note)

    /*  Claude app: stand-alone skill folders (skills without templates are identical in both variants)  */
    for (const name of names) {
        const zip = new JSZip()
        for (const [ rel, data ] of skillFiles(name, msgFiles, true))
            zip.file(`${name}/${rel}`, data)
        await write(zip, `${name}${tag ? `-${tag}` : ""}.zip`, note)
    }
}
