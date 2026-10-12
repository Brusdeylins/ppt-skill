/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  infra/version: the CLI version fact.
*/

/**  current CLI version (injected at build time, dev fallback otherwise)  */
export const VERSION: string = typeof PPTC_VERSION !== "undefined" ? PPTC_VERSION : "0.0.0-dev"
