/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  build.mjs: bundle the CLI into a single executable ESM file (dst/pptc.mjs).
**  Runtime dependencies are bundled, so the published artifact needs no
**  node_modules at execution time.
*/

import * as esbuild from "esbuild"
import { readFileSync } from "node:fs"

const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8"))

await esbuild.build({
    entryPoints: ["src/cli/main.ts"],
    outfile: "dst/pptc.mjs",
    bundle: true,
    platform: "node",
    format: "esm",
    target: "node20",
    banner: {
        /*  bundled CommonJS dependencies need require/__dirname in the ESM bundle  */
        js: "#!/usr/bin/env node\n"
            + "import { createRequire as __pptcCreateRequire } from 'node:module';\n"
            + "import { fileURLToPath as __pptcFileURLToPath } from 'node:url';\n"
            + "import { dirname as __pptcDirname } from 'node:path';\n"
            + "const require = __pptcCreateRequire(import.meta.url);\n"
            + "const __filename = __pptcFileURLToPath(import.meta.url);\n"
            + "const __dirname = __pptcDirname(__filename);"
    },
    define: {
        "PPTC_VERSION": JSON.stringify(pkg.version)
    },
    logLevel: "info"
})
