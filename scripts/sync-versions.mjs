/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  sync-versions.mjs: propagate the single source-of-truth version from
**  package.json into the skill VERSION sidecars so a release never leaves
**  one of them stale. Run after `version-bump.mjs`, as part of `plugin:sync`.
**  The plugin version (plugin.json) is owned by the marketplace and untouched.
**
**  Derived sources (2 files):
**    plugin/skills/ppt/VERSION                skill self-report
**    plugin/skills/ppt-prepare/VERSION        skill self-report
*/

import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..")
const version = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8")).version

/*  overwrite a plain-text VERSION sidecar  */
const writeVersionFile = rel =>
    writeFileSync(path.join(root, rel), `${version}\n`)

writeVersionFile("plugin/skills/ppt/VERSION")
writeVersionFile("plugin/skills/ppt-prepare/VERSION")

console.log(`synced version ${version} -> 2x VERSION`)
