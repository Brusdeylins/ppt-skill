<p align="center">
  <img src="https://raw.githubusercontent.com/Brusdeylins/ppt-skill/main/logo.png" alt="pptc logo" width="200"/>
</p>

# pptc

**Deterministic PowerPoint (PPTX) CLI for LLM agents.**

`pptc` creates and edits PowerPoint presentations from the command line --
template-aware, schema-validated and atomic. It is designed to be driven by an
LLM agent (e.g. a Claude Code skill): every command emits exactly one JSON
envelope on stdout, every failure has a stable error code and exit code, and
every write either fully applies or leaves the file byte-identical.

```
$ pptc state deck.pptx --level summary
{
  "ok": true,
  "cmd": "state",
  "file": "/work/deck.pptx",
  "rev": "27d7f5a4ea5a",
  "result": { "slideCount": 4, "slides": [ ... ] }
}
```

## Why

Building PPTX files through ad-hoc scripting is fragile for agents: shell
quoting mangles JSON, slide indices shift when users edit between turns,
half-applied edits corrupt decks, and errors arrive as prose. `pptc` fixes
these structurally:

- **One write path.** All mutations are *ops* in a JSON document, validated as
  a whole (Zod schemas) before the first byte is touched, then applied in one
  atomic write. `--dry-run` is the same pipeline without the write.
- **Stable addressing.** Slides are addressed by their OOXML `sldId`
  (`id:257`), exact title (`title:Agenda`), or a document-local `$ref` --
  not by fragile positions (though `index:N` exists as an escape hatch).
- **Optimistic locking.** `state` returns a `rev` token; `apply --rev` (or
  `expectRev` in the document) fails with exit 6 when the deck changed
  underneath.
- **Templates as data.** `tpl describe` turns any `.potx`/`.pptx` into an
  LLM-readable description: per layout an ASCII minimap, semantic positions
  ("left column, full height"), text capacities and image aspect ratios --
  derived generically from the OOXML geometry, no template-specific code.
- **Typed failures.** Exit codes per error class, machine-readable
  `error.code`, and Zod issue paths in `error.details` for self-correction.

## Installation

```
npm install -g @brusdeylins/pptc
pptc --version
```

Or without installing: `npx @brusdeylins/pptc ...`. Requires Node.js >= 20.

## The PowerPoint Skill (Claude Code plugin)

This repository also ships **`ppt`**, a Claude Code skill built on
pptc. It turns Claude Code into a presentation engineer: template-aware
deck building with an outline-first workflow, content guardrails (action
titles, capacity-checked text, speaker notes, footer rules) and
color-faithful image prompts written onto picture placeholders --
overlay-aware, so covered regions become negative space.

A companion skill, **`ppt-prepare`**, plans the *content* first (story,
MECE arguments, headline titles, a per-slide plan) and hands the approved
plan to `ppt`.

```
/plugin marketplace add Brusdeylins/ppt-skill
/plugin install ppt@ppt-skill
```

No further setup: the plugin bundles its own pptc build
(`skills/ppt/scripts/pptc.mjs`) and a neutral fallback template;
the plugin-root `bin/` puts a `pptc` command on the Bash tool's PATH
while the plugin is enabled. The public plugin ships only a neutral
default template; point the skill at any external `.potx`/`.pptx`
(optionally with a `<name>.md` sidecar for template-specific knowledge),
or build an in-house package that bundles your company templates (see
[Packaging & Distribution](#packaging--distribution)). Deck-level settings
(language, image style, ...) persist as **custom document properties inside the
`.pptx`**, so the deck is self-describing -- hand it to someone else and the
skill reads its setup straight back, no side file.

The skill definitions live in
[`plugin/skills/ppt/SKILL.md`](plugin/skills/ppt/SKILL.md) and
[`plugin/skills/ppt-prepare/SKILL.md`](plugin/skills/ppt-prepare/SKILL.md);
each imports its own `meta/control.md`.

### Use it on claude.ai (web -- no build, no Node setup)

Not a developer? Each skill is published as a ready-to-upload ZIP on the
[Releases page](https://github.com/Brusdeylins/ppt-skill/releases). Steps:

1. **Enable code execution** (skills require it, otherwise they appear
   greyed out): **Settings → Capabilities → "Code execution and file
   creation"**. On Team/Enterprise plans the org owner enables it under
   **Organization settings → Skills**.
2. Download `ppt.zip` and `ppt-prepare.zip` from the latest release.
3. Open the **sidebar → Customize → Skills → Create skill → Upload skill**
   and upload each ZIP -- one skill per ZIP, so upload **both** separately.

That is the whole install -- no clone, no `npm`. The bundled pptc is
self-contained and runs on the claude.ai code-execution runtime.

**Handing the plan from `ppt-prepare` to `ppt` on claude.ai.** Unlike the
local CLI, each claude.ai skill runs in its own sandbox that is not shared,
so `ppt` cannot see a file `ppt-prepare` wrote to disk. `ppt-prepare` now
produces a **single** file -- `<deck>-plan.md`, with the deck setup in its
header (no separate sidecar) -- so when you switch to `ppt`, just **download
that one plan file and attach/upload it** (ideally in the same conversation).
`ppt` reads its setup straight from the plan.

## Quick Start

```bash
# 1. understand the template (the LLM picks layouts from this)
pptc tpl describe corporate.potx

# 2. create a deck and build it in one run
pptc new deck.pptx --template corporate.potx --ops @ops.json

# 3. read the truth back
pptc state deck.pptx --level full
```

`ops.json`:

```json
{
  "ops": [
    { "op": "slide.add", "ref": "intro", "layout": 0,
      "placeholders": {
        "title":    { "text": "My Talk" },
        "subtitle": { "text": "Subtitle" }
      },
      "notes": "Welcome the audience." },
    { "op": "slide.add", "layout": 4,
      "placeholders": { "title": { "text": "Figures" } } },
    { "op": "el.add", "slide": "title:Figures", "elements": [
      { "type": "chart", "frame": { "x": 0.7, "y": 1.9, "w": 12, "h": 4.7 },
        "data": { "type": "column", "categories": ["Q1", "Q2"],
                  "series": [{ "name": "Revenue", "values": [10, 14] }] } }
    ] },
    { "op": "slide.move", "slide": "$intro", "to": 0 }
  ]
}
```

Small fixes need no JSON file:

```bash
pptc text deck.pptx --slide title:Figures --ph title "Figures 2026"
pptc note deck.pptx --slide id:257 "New speaker notes"
pptc rm   deck.pptx --slide index:3
pptc apply deck.pptx -e '{"op":"slide.move","slide":"id:257","to":1}'
```

## Command Reference

Every command prints exactly one JSON envelope on stdout (see
[Envelope & Exit Codes](#envelope--exit-codes)).

**Slide selectors** (everywhere a `SEL` appears): `id:N` (canonical OOXML
`sldId`, survives reordering) · `title:...` (exact title, must be unique) ·
`index:N` or bare digits (positional escape hatch) · `$ref` (a slide created
earlier in the same ops document).

**Payloads:** `--ops @file.json` reads a file, `--ops -` reads stdin, `-e`
takes one inline op. Agents should write the file with their editor tool and
pass `@file` -- that removes shell quoting from the threat model.

**Human console:** `--plain` on the read commands (`help`, `state`,
`tpl list`, `tpl describe`, `tpl validate`) prints readable text instead of
the JSON envelope (whose strings escape newlines as `\n`). Agents should
stick to the envelope; write commands are envelope-only by design.

```bash
pptc --help                          # overview, plain text
pptc apply --help                    # any command in detail (options explained)
pptc help ops --plain                # the write API: ops document anatomy
pptc help selectors --plain          # slide addressing
pptc state deck.pptx --plain         # readable slide list
pptc tpl describe corp.potx --plain  # readable template description
```

### Read templates

#### `pptc tpl list <dir>`

Inventory of all `.potx`/`.pptx` files in a directory, sorted, each with a
`sidecar` flag telling whether a human-curated `<name>.md` companion file
exists next to it.

#### `pptc tpl describe <tpl> [--layout SEL] [--format text|json] [--plain]`

The LLM-facing template description. For every layout: an ASCII minimap of
the placeholder geometry, semantic positions ("left column, full height"),
text capacities (`~N lines of ~M chars`), image aspect ratios, overlay
warnings for picture placeholders and a suitability hint -- all derived
generically from OOXML geometry. A sidecar
`<tpl>.md` (template-specific notes: layout roles, footer pattern, design
constraints) is included verbatim in the header.

- `--layout SEL` -- restrict to one layout (zero-based index or exact name)
- `--format json` -- return the raw `TemplateInfo` data instead of Markdown
- `--plain` -- print the Markdown directly WITHOUT the JSON envelope
  (human console use; agents should stick to the envelope)

#### `pptc tpl inspect <tpl> [--layout SEL]`

The precise machine model: slide size, theme fonts, the full theme color map
(`dk1`, `lt1`, `accent1`..`accent6`, ... as `RRGGBB`), and per layout every
placeholder with OOXML `idx`, kind, shape name, frame (inches) and text
capacity. Picture placeholders carry `overlays`: which text shapes sit on
top of the image and where (region relative to the image, e.g. "bottom
area, left part") -- so image prompts can keep those regions as negative
space. Use this when you need exact values; use `describe` when an LLM
should pick layouts.

#### `pptc tpl validate <tpl>`

Checks a template against pptc's expectations (layouts present, notes master
for speaker notes, ...). Reports `issues` with severities; exit 7 when a
`fail`-grade issue exists.

### Read decks

#### `pptc state <deck> [--slide SEL] [--level summary|text|full]`

The deck read model and the **`rev` token** for optimistic locking.

- `--level summary` -- ids, indices, titles, layout indices, and `hidden:
  true` on every hidden slide ("Hide Slide"); preserved across `apply`
- `--level text` (default) -- plus all placeholder texts and notes
- `--level full` -- plus every shape with geometry and styling: type,
  name, `frame`, text, table cells + column widths, autoshape
  preset/fill/border/font, a picture's media file name (`image`), and --
  for a body with explicit run formatting -- its `paragraphs`/`runs`
  (per-run font, size, bold, italic, color, round-tripping into
  `slide.fill`) -- enough to recreate or edit a styled table, diagram or
  code block without touching the raw XML
- `--slide SEL` -- restrict to one slide

Read-before-write protocol: take `rev` from here and pass it to
`apply --rev` (or `expectRev` in the ops document).

### Verify decks

#### `pptc verify <deck> [--strict]`

Checks a finished deck against every known PowerPoint **"repair" trigger**
(dangling relationships, stale `docProps/app.xml`, shared notesSlides, broken
content types, schema-order and reference errors). `apply` already self-verifies
its own output before writing, so a freshly built deck is clean by construction;
run `verify` as an explicit gate after a build, or on any `.pptx` whose origin
you do not control.

- `result.ok` -- `true` when there are no findings
- `result.findings` -- the list of repair triggers (empty when clean)
- `--strict` -- any finding is a hard failure (exit 8) instead of a report

### Write decks

#### `pptc new <deck> --template <tpl> [--force] [--ops @file] [--strict]`

Creates a valid, zero-slide deck carrying the template's masters, layouts and
theme. Refuses to overwrite an existing file unless `--force` is given. With
`--ops` the deck is built in the same run (same semantics as `apply`).

#### `pptc apply <deck> (--ops @file|- | -e '<op>') [--template <tpl>] [--dry-run] [--strict] [--rev R] [--out F]`

The single write path. Validates and plans the **whole** ops document first
(schema, every selector, capacity lint), then applies everything in one
atomic write -- or nothing (`failedAt` in the error tells which op failed).

- `--ops @file` / `--ops -` / `-e '<op-json>'` -- exactly one of these
- `--template <tpl>` -- required whenever the document contains `slide.add`
  (new slides are instantiated from the template's layouts)
- `--dry-run` -- full validation and planning, no write; warnings included
- `--strict` -- lint warnings become exit 7: `W_TEXT_OVERFLOW` (planned
  text exceeds the placeholder capacity -- shorten or split),
  `W_ELEMENT_OVERLAP` (an `el.add` element covers a text-bearing shape,
  incl. the footer/slide-number area -- reposition it; prompt boxes are
  exempt) and `W_FONT_TOO_SMALL` (an explicit run/element font is below the
  readable minimum -- enlarge it)
- `--min-font-pt N` -- readability floor for explicit font sizes (default
  `11`); runs or `el.add` elements below `N` raise `W_FONT_TOO_SMALL`. `0`
  disables the check. Footer/slide-number/date placeholders and prompt
  boxes are always exempt
- `--rev R` -- optimistic lock: fail with exit 6 unless the deck still has
  revision `R` (alternative: `expectRev` inside the document)
- `--out F` -- write the result to a new file, leave the input untouched

### Micro edits (no JSON document needed)

Each compiles to exactly one `slide.fill`/`slide.rm`/`slide.move` op and runs
through the same validated, atomic apply path. All of them accept
`--rev R`, `--strict` and `--dry-run`.

#### `pptc text <deck> --slide SEL [--ph KEY] [--append] "text"`

Set placeholder text. `--ph` takes a placeholder key (default `title`);
`--append` appends instead of replacing.

#### `pptc note <deck> --slide SEL "speaker notes"`

Set the slide's speaker notes.

#### `pptc footer <deck> [--slide SEL] "footer text"`

Set the footer of one slide -- or, without `--slide`, of **every** slide.
Works by cloning the layout's footer placeholder; layouts without one
(typically title/closing layouts) are skipped silently.

#### `pptc rm <deck> --slide SEL`

Remove a slide.

#### `pptc move <deck> --slide SEL --to N`

Move a slide to zero-based position `N`.

### Self-description

#### `pptc schema [op|document]`

JSON Schema of one op (e.g. `pptc schema slide.fill`) or of the whole ops
document -- generated from the validating Zod schemas, so it is always
authoritative. Without argument: the list of op names.

#### `pptc update`

Self-update via npm (see [Updating](#updating)).

#### `pptc help` / `pptc --version`

Usage summary / version envelope.

**Placeholder keys** in `slide.fill`/`text`: the OOXML `idx` (`"13"`), or
semantic keys resolved against the layout: `"title"`, `"subtitle"`,
`"body"`, `"image"`, `"image:14"`, `"text:13"`.

## The Ops Document (the write API)

Reading is done with commands; **writing is done with data**. Every
mutation -- adding slides, filling placeholders, moving, free elements --
is expressed as one JSON document, the *ops document*, and handed to
`apply` (or `new --ops`). pptc validates the WHOLE document first, then
applies everything in one atomic write, or nothing.

```json
{
  "expectRev": "27d7f5a4ea5a",
  "ops": [
    { "op": "slide.add", "layout": "CONTENT", "ref": "intro",
      "placeholders": { "title": { "text": "My Talk" },
                        "body":  { "text": "Point A\nPoint B" } },
      "notes": "Speaker notes.",
      "footer": "My Talk | 2026" },
    { "op": "slide.move", "slide": "$intro", "to": 0 }
  ]
}
```

Anatomy:

- **`expectRev`** (optional): the `rev` token from `pptc state`. The
  apply fails with exit 6 when the deck changed in between
  (read-before-write protocol). `--rev R` on the command line is the
  same thing.
- **`ops`**: executed in order. Each op carries an `"op"` discriminator
  and addresses slides via selectors (`id:`, `title:`, `$ref`,
  `index:` -- see below).
- **Fill payloads** (`slide.add`/`slide.fill`): `placeholders` maps a
  key (semantic like `"title"`/`"body"`/`"image"`, or the OOXML idx
  like `"13"`) to `{ "text": ... }`, `{ "image": "path.png" }` or
  `{ "text": ..., "append": true }`. Plus optional `notes`, `footer`,
  `background` and `hidden` (`true` hides the slide -- "Hide Slide" --,
  `false` shows it; omitted keeps the current state). Text accepts plain
  strings (`\n` = new paragraph) or rich text (runs with
  bold/italic/color/size, bullet levels).
- **Authoritative field reference:** `pptc schema <op>` prints the
  exact JSON Schema of any op, `pptc schema document` the whole
  document -- generated from the validating code, never stale.
  `pptc help ops` summarizes all of this on the console.

Recommended agent loop:
`state` → write `ops.json` → `apply --dry-run --strict` → fix findings
→ `apply --rev <rev>`.

## Ops Reference

| Op | Purpose |
|---|---|
| `slide.add` | add a slide from a template layout (`layout`: index or name; optional `ref`, `at`, inline fill) |
| `slide.fill` | fill placeholders (`text`/`image`), `notes`, `footer`, `background`, `hidden` (hide/show) of a slide |
| `slide.rm` / `slide.move` / `slide.copy` | structure edits |
| `el.add` | add free elements: `textbox`, `table`, `chart`, `shape`, `image`, `connector` -- all share `frame: {x,y,w,h}` in inches |
| `el.set` / `el.rm` | retext / remove an element by shape name; matches exactly or as prefix of the engine's UUID-suffixed names (`Kasten` matches `Kasten-1d22c8b0-...`), and also targets elements generated earlier in the same ops document |
| `img.prompts` | overlay picture placeholders with visible image-prompt boxes (removable via `el.rm`) |
| `meta.props` | document properties: core fields (title, author, subject, keywords, category, comments) and/or arbitrary `custom` name/value pairs stored in the `.pptx` (read back by `state` -- the deck's self-describing memory) |

Run `pptc schema <op>` for the authoritative JSON Schema with all fields, or
`pptc schema document` for the whole ops document.

### Ops document semantics

- The **whole document is validated and planned first** (schema, every
  selector resolved, capacity lint); mutations run only when everything
  resolved. On failure the envelope carries `failedAt` and the deck is
  byte-identical to before.
- `expectRev` (or `--rev`) enforces the read-before-write protocol: exit 6 on
  mismatch.
- Text capacity is linted against the template geometry; warnings appear in
  the envelope, `--strict` turns them into exit 7.

## Envelope & Exit Codes

Exactly one JSON document on stdout; logs (if any) on stderr.

```json
{ "ok": true,  "cmd": "apply", "file": "...", "rev": { "before": "...", "after": "..." },
  "result": { "applied": 6, "slides": { "intro": { "id": 263, "index": 0 } } },
  "warnings": [ { "code": "W_TEXT_OVERFLOW", "placeholder": 13, "estimatedLines": 14, "maxLines": 12 } ] }
```

```json
{ "ok": false, "cmd": "apply",
  "error": { "code": "E_SCHEMA", "message": "...", "details": { "issues": [ ... ] } } }
```

| Exit | Class | Codes |
|---|---|---|
| 0 | success (possibly with warnings) | |
| 2 | input | `E_USAGE`, `E_SCHEMA`, `E_JSON` |
| 3 | addressing | `E_ADDR_NOTFOUND`, `E_ADDR_AMBIGUOUS` |
| 4 | file/template I/O | `E_FILE`, `E_TEMPLATE` |
| 5 | engine | `E_ENGINE` |
| 6 | revision conflict | `E_REV_CONFLICT` |
| 7 | lint under `--strict` | `E_LINT` |
| 8 | integrity (self-verify / `verify --strict`) | `E_INTEGRITY` |

## Architecture in a Nutshell

The codebase is layered strictly downward (`cli → commands → core → engine →
infra`, see [ARCHITECTURE.md](ARCHITECTURE.md)). Inside the engine, the actual
PPTX work is split across **three workhorses** -- two external libraries plus
one deliberate own layer:

1.  **[pptx-automizer](https://github.com/singerla/pptx-automizer)** -- the
    heavy lifting: imports slides between decks (kept slides from the deck
    itself, new slides from a per-template *seed deck*), carrying masters,
    layouts and styles along.
2.  **[PptxGenJS](https://github.com/gitbrent/PptxGenJS)** -- element
    generation: tables, charts, textboxes, shapes and images created via
    `el.add` are built with PptxGenJS and merged in through automizer's
    interop.
3.  **The zip post-pass** (`engine/post.ts`, one file) -- everything the
    libraries cannot express or get wrong: speaker notes, footer cloning,
    backgrounds, images into placeholders, hyperlink relationships, slide
    visibility (re-applying `show="0"` on kept hidden slides), document
    properties -- plus deterministic cleanup after automizer (orphan parts,
    stale relationships, duplicate shape ids).

The principle: libraries for the 95%, a small deterministic repair layer for
the 5% they cannot do -- never a self-built PPTX engine.

## Design Decisions

- **Kernel CLI.** Reads are commands, writes are data. An agent is better at
  producing one schema-validated JSON document than at sequencing a dozen
  correct shell invocations -- one validation pass surfaces all errors at
  once, and there is exactly one write schema to maintain.
- **Seed decks.** The engine ([pptx-automizer](https://github.com/singerla/pptx-automizer))
  imports slides, not layouts. For each template pptc derives a *seed deck*
  (one empty cloned-placeholder slide per layout) on demand, cached by
  template content hash -- `slide.add` works with any template. With no
  template, pptc seeds from the deck's OWN embedded layouts instead, so
  `slide.add` keeps working on an existing deck with no side file -- a deck
  is self-contained.
- **Two-phase apply.** Every op is a pure plan transformer; the engine session
  interprets the finished plan. `--dry-run` is "plan without commit", not a
  separate code path.
- **No silent first-match.** Ambiguous titles, ambiguous placeholder keys and
  unknown refs are errors with candidate lists -- determinism beats
  convenience.

See [ARCHITECTURE.md](ARCHITECTURE.md) for the layer model and data flow.

## Updating

`pptc` checks the npm registry at most once per day (offline-tolerant). When a
newer version exists, every envelope carries an `update` field:

```json
"update": { "current": "0.1.0", "latest": "0.2.0" }
```

An agent (or you) then runs:

```
pptc update
```

This is how a skill built on pptc keeps itself current without user
intervention.

## Development

```
npm install        # dependencies
npm test           # build + vitest (unit, golden, integration, contract)
npm run lint       # eslint (incl. TSDoc) + tsc --noEmit
npm run build      # esbuild bundle -> dst/pptc.mjs
npm run plugin:sync # rebuild + copy the bundle into the plugin (+ both VERSION files)
npm run skill:zip  # package each skill as deploy/<skill>.zip (claude.ai upload)
```

## Packaging & Distribution

How the `ppt` skill is bundled, packaged into upload-ready ZIPs (including
in-house builds with your own company templates) and released.

### Skill ZIPs (with your own templates)

Each skill is packaged as a self-contained ZIP with the skill folder at
the ZIP root (e.g. `ppt/SKILL.md`), written to `deploy/` (a git-ignored
build artifact). The ZIP is uploaded in claude.ai via the sidebar →
Customize → Skills → Create skill → Upload skill (code execution must be
enabled under Settings → Capabilities). Every variant runs `plugin:sync`
first, so the bundled pptc is always fresh.

```
npm run skill:zip               # PUBLIC: bundles only the neutral default
                                #   template  ->  deploy/ppt.zip
npm run skill:zip:internal      # IN-HOUSE: REPLACES the neutral default with
                                #   every template in ./private-templates  ->  deploy/ppt-internal.zip
npm run skill:zip -- --from DIR # IN-HOUSE: same, but from DIR instead
                                #   (no need to pre-stage files)  ->  deploy/ppt-internal.zip
```

- **Public build** ships only `assets/neutral-template.pptx` -- this is
  the ZIP attached to GitHub releases for non-developers.
- **Internal build** overlays your own `.potx`/`.pptx` (plus optional
  `<name>.md` sidecars) into the skill's `assets/` and **drops the neutral
  default** (`neutral-template.pptx`) -- the in-house package ships ONLY your
  templates, so no one picks the generic Office look by accident. When no
  template is named the skill uses your template (or, with several, offers a
  choice). Sources come from `private-templates/` (git-ignored) or any
  `--from <dir>`.
- **Company templates never touch git.** They live only inside the
  internal ZIP; `private-templates/` and `deploy/` are both git-ignored,
  so corporate material is never committed or released publicly.

Templates are richer with a sidecar: put `<name>.md` next to `<name>.potx`
(layout-role map, footer pattern, design constraints) and it is bundled
alongside, so `tpl describe`/the skill pick up the template-specific
knowledge automatically.

### Versioning & release rule

This repo carries two independently versioned artifacts: the **pptc CLI**
(`package.json`, published to npm) and the **Claude Code plugin**
(`plugin/.claude-plugin/plugin.json` + marketplace entry, distributed via
git). Each skill carries a `VERSION` file at its root
(`plugin/skills/ppt/VERSION`, `plugin/skills/ppt-prepare/VERSION`) recording
which build it ships -- both written by `plugin:sync`.

- **Skill-only change** (SKILL.md, references, meta): bump the plugin and
  marketplace version, add a CHANGELOG entry, commit and push `main`,
  tag `plugin-v<version>` and create a GitHub release from the CHANGELOG
  notes (visibility of the version history). No npm; marketplace users
  update via `/plugin marketplace update`.
- **CLI/engine change**: bump `package.json`, run `npm run plugin:sync`
  and bump the plugin version too (a new bundle is a new plugin release).
  CHANGELOG entry, commit, tag `v<version>`, push `main` + tag, GitHub
  release -- and `npm publish --access public` once the package is live
  (`prepublishOnly` enforces lint + tests + build).

Either way, attach the web-upload ZIP(s) to the GitHub release so
non-developers can install on claude.ai without building:
`npm run skill:zip` then `gh release upload <tag> deploy/ppt.zip`
(`deploy/` is a git-ignored build artifact -- the release asset is the
distribution point, not the repo tree).

The release asset is always the PUBLIC ZIP (neutral default only). Builds
that bundle company templates (`skill:zip:internal` / `--from`) produce
in-house ZIPs that are distributed within the company and never attached
to a git release -- see *Packaging skill ZIPs* above.

CLI change ⇒ plugin release, but skill change ⇏ CLI release. Lint and the
test suite must be green before any push.

## License

MIT -- Copyright (c) 2026 Matthias Brusdeylins
