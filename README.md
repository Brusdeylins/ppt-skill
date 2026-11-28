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
(optionally with a `<name>.md` sidecar for template-specific knowledge) --
in-house distributions with company templates live in a downstream
company repo, not here. Deck-level settings
(language, image style, ...) persist as **custom document properties inside the
`.pptx`**, so the deck is self-describing -- hand it to someone else and the
skill reads its setup straight back, no side file.

The skill definitions live in
[`plugin/skills/ppt/SKILL.md`](plugin/skills/ppt/SKILL.md) and
[`plugin/skills/ppt-prepare/SKILL.md`](plugin/skills/ppt-prepare/SKILL.md);
both import the shared [`plugin/meta/control.md`](plugin/meta/control.md)
(control tags, step banners, progress task list, the Asking-the-User
dialog, stage gate) -- one file for all skills of the plugin.

### Use it in the Claude app (desktop / claude.ai)

Not a developer? The Claude app installs the same plugin from this repo:
**Settings → Plugins → Add → "Add marketplace"**, enter
`Brusdeylins/ppt-skill`, then install the **ppt** plugin. No clone, no
`npm` -- the bundled pptc is self-contained. (Per-skill ZIP uploads were
retired in 1.0.6; the marketplace is the single distribution channel.)

**Handing the plan from `ppt-prepare` to `ppt` when files are not
shared.** On hosts where each skill run has its own sandbox (no shared
filesystem), `ppt` cannot see a file `ppt-prepare` wrote to disk.
`ppt-prepare` produces a **single** file -- `<deck>-plan.md`, with the deck
setup in its header (no separate sidecar) -- so when you switch to `ppt`,
just **download that one plan file and attach/upload it** (ideally in the
same conversation). `ppt` reads its setup straight from the plan.

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
- `--level text` (default) -- plus all placeholder texts and notes, and
  each slide's `section` (with the deck's `sections` in order) when the
  deck has sections
- `--level full` -- plus every shape, listed back to front (the stacking
  order), with geometry and styling: type,
  name, `frame`, text, table cells + column widths (`colWidths`), merged
  cells (`merges`: row, col, rowspan, colspan; covered cells read as `""`),
  vertical cell text (`cellTextDirections`) and a header row
  (`headerRow: true`), autoshape
  preset/fill (with its `transparency`)/border/font, the `textDirection`, `valign` and `margin` of a
  text body, the `rotation`, `altText` and `title` of a shape, its click
  target (`hyperlink`: URL or `slide:id:N`), connector
  `startArrow`/`endArrow`, the text language (`lang`), a picture's media
  file name (`image`), and -- for a body with explicit run formatting or
  links -- its `paragraphs`/`runs` (per-run font, size, bold, italic,
  strike, superscript/subscript, color, highlight, hyperlink; per
  paragraph bullet and spacing -- round-tripping into
  `slide.fill`) -- enough to recreate or edit a styled table, diagram or
  code block without touching the raw XML. Colors read as RRGGBB, theme
  colors resolved against the deck's theme
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
- `--template <tpl>` -- not needed on an existing deck: `slide.add` uses the
  deck's own embedded layouts. A template whose layouts differ from the
  deck's (another design or version) fails with `E_TEMPLATE`.
- `--dry-run` -- full validation and planning, no write; warnings included
- `--strict` -- lint warnings become exit 7: `W_TEXT_OVERFLOW` (planned
  text exceeds the placeholder capacity -- shorten or split),
  `W_ELEMENT_OVERLAP` (an `el.add` element covers a text-bearing shape,
  incl. the footer/slide-number area -- reposition it; prompt boxes and
  elements placed or moved to the back are exempt) and `W_FONT_TOO_SMALL` (an explicit run/element font is below the
  readable minimum -- enlarge it). Not escalated, since it is found during
  the write: `W_REPAIRED` (the input deck carried PowerPoint repair
  triggers on a slide or a layout -- duplicate shape ids, dead or duplicate
  relationships, references to missing relationships, a shared notes part
  -- and pptc repaired it, naming what it removed; the deck is written,
  tell the user) and `W_INPUT_DEFECT` (the input deck carries repair
  triggers pptc cannot repair, e.g. a slide listed twice; they stay as
  they were, only triggers pptc itself would add block the write), and
  `W_REPLACE_SKIPPED` (a `text.replace` found its text in a chart or
  SmartArt of a targeted slide, which it leaves as it is -- the message
  names the slide and part; tell the user). Content
  of a slide that would vanish, move or change although no op asked for it
  fails with `E_INTEGRITY`
- `--min-font-pt N` -- readability floor for explicit font sizes (default
  `8`); runs or `el.add`/`el.set` elements below `N` raise
  `W_FONT_TOO_SMALL`. `0` disables the check. Footer/slide-number/date
  placeholders and prompt boxes are always exempt. The floor is a pure
  legibility backstop: font sizes themselves come from the template
  (`tpl inspect` reports each placeholder's size), never from pptc. Also
  accepted by `new` and the quick-edit commands
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

Set the slide's speaker notes. An existing notes page keeps everything but
the text of its notes body; a new one is built from the notes master
(slide image, notes body, slide number), as PowerPoint does.

#### `pptc footer <deck> [--slide SEL] ["footer text"] [--number] [--date]`

Set the footer of one slide -- or, without `--slide`, of **every** slide.
Works by cloning the layout's footer placeholder; layouts without one
(typically title/closing layouts) are skipped silently. `--number` and
`--date` show the slide number and the date as well (the layout's
slide-number and date placeholders with their fields); with one of them
the text may be omitted.

#### `pptc rm <deck> --slide SEL`

Remove a slide.

#### `pptc move <deck> --slide SEL --to N`

Move a slide to zero-based position `N`.

### Self-description

#### `pptc schema [op|document]`

JSON Schema of one op (e.g. `pptc schema slide.fill`) or of the whole ops
document -- generated from the validating Zod schemas, so it is always
authoritative. Without argument: the list of op names.

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
  `{ "text": ..., "append": true }`. An image takes the placeholder
  frame by its `fit`: `"crop"` (default) covers the frame and cuts the
  overhang off centered, like PowerPoint; `"contain"` shows the whole image
  in a centered, shrunk frame; `"stretch"` distorts it to the frame (pixel
  sizes are read from PNG, JPEG, GIF, WebP, BMP and SVG; an unknown size
  stretches). Plus optional `notes`, `footer`,
  `background` and `hidden` (`true` hides the slide -- "Hide Slide" --,
  `false` shows it; omitted keeps the current state), `slideNumber` and
  `date` (`true` clones the layout's slide-number / date placeholder with
  its field and layout content, a picture included -- nothing when the
  layout has none --, `false` removes it with that content; omitted keeps
  it as it is), `section` (starts a PowerPoint section at
  the slide -- see [Sections](#sections)). Text accepts plain
  strings (`\n` = new paragraph, `\v` -- a vertical tab -- a line break
  inside the paragraph, as PowerPoint writes it) or rich text (runs with
  bold/italic/color/size, bullet levels).
- **Rich text**: a paragraph is `{ "text": ... }` or `{ "runs": [...] }`
  with `level`, `alignment`, `bullet` (`true`, `"number"` for 1. 2. 3.,
  `false`; in placeholders `true` keeps the placeholder's own bullet
  style where its level shows one, else -- a subtitle, a caption, a
  textbox, shape or table cell -- it writes a bullet `•` with a hanging
  indent), `spaceBefore`/`spaceAfter` (pt) and `lineSpacing`
  (a multiple, `1.5`); a run takes `bold`, `italic`, `underline`, `strike`,
  `superscript`, `subscript`, `size`, `font`, `color`, `highlight`,
  `hyperlink` (see *Links*) and `lang` (paragraph-level character fields
  apply to a paragraph without runs). `font` takes a typeface name or a
  theme font: `+mj-lt` (the theme's heading font), `+mn-lt` (its body
  font) -- the text then follows the template's fonts. The same vocabulary works in
  `slide.fill`, `el.set` and `el.add`. A `\v` in a text is a line break
  inside its paragraph (`a:br`); `state` reads a line break back as `\v`.
  Characters a PowerPoint file cannot store (control characters but tab,
  line feed and carriage return; `\v` outside text) fail with `E_SCHEMA`
  naming the field. Text without a `color` takes the template's color
  (`el.add` included) -- on an `el.add` fill (textbox, shape, table cell,
  header or alternating row) the theme's `dk1` or `lt1`, whichever
  contrasts more with the fill (a transparent fill as it shows over the
  background `bg1`). An `el.set` `fill` keeps the text's color: give
  `fontColor` with a new fill the text must contrast with. There is no shrink-to-fit: text that overflows is
  shortened or split (`W_TEXT_OVERFLOW`); replaced text drops the shrink
  scale PowerPoint had computed for the old one.
- **Links**: `hyperlink` -- on a run, and on a `textbox`, `shape` or
  `image` as the click target of the whole element (`el.add`, `el.set`)
  -- takes a URL or a slide link `slide:SEL`: `slide:$ref`,
  `slide:id:256`, `slide:title:Agenda`, `slide:index:2`. A slide link is a
  jump to that slide (a slide relationship with
  `ppaction://hlinksldjump`); its target is resolved after the whole
  document, so slides added later in it count, moved slides stay the
  target, and a target that does not exist then fails with
  `E_ADDR_NOTFOUND`. When a later apply removes the target, the link goes
  (the text stays); a copied slide keeps its links. `state --level full`
  reads links back as the URL or `slide:id:N`.
- **Colors**: every color field takes a hex value (`"1F4E79"`, `"#1F4E79"`)
  or a theme slot -- `accent1`-`accent6`, `tx1`, `tx2`, `bg1`, `bg2`,
  `dk1`, `lt1`, `dk2`, `lt2`, `hlink`, `folHlink` -- written as a theme reference
  (`a:schemeClr`, no tint/shade), so the element follows the template's
  color scheme. Text takes `tx1`/`tx2`, fills and backgrounds `bg1`/`bg2`
  (they follow the master's color map, so they stay readable on dark
  templates); `dk1`/`lt1`/`dk2`/`lt2` are absolute -- for contrast on
  images. A `fill` takes a `transparency` in percent (`0` opaque - `100`
  invisible; `el.add` textbox, shape and table cell, `el.set`). `state`
  reports colors as RRGGBB (theme colors resolved).
- **Text language**: every run pptc writes carries a `lang` (PowerPoint's
  proofing language). It is the deck language -- the custom property
  `pptcDeckLang` (`meta.props`), else the deck's own default text language,
  else `en-US` -- unless a `textbox`/`shape`/`table` element, a paragraph
  or a run states its own `lang` (e.g. `"de-DE"`). `state --level full`
  reports each shape's `lang`.
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
| `slide.fill` | fill placeholders (`text`/`image`), `notes`, `footer`, `slideNumber`, `date`, `background`, `hidden` (hide/show), `section` of a slide |
| `slide.rm` / `slide.move` / `slide.copy` | structure edits |
| `el.add` | add free elements: `textbox`, `table`, `chart`, `shape`, `image`, `connector` -- all share `frame: {x,y,w,h}` in inches (see below) |
| `el.set` / `el.rm` | change / remove an element by shape name in place; matches exactly or as prefix of the engine's UUID-suffixed names (`Kasten` matches `Kasten-1d22c8b0-...`), and also targets elements generated earlier in the same ops document and the placeholders of slides added in it (`PptcPh-<idx>`); `el.set` takes text, frame, style, alt text, image, table cells/rows and the stacking order -- see below |
| `text.replace` | find and replace plain text in the whole deck or on one `slide` -- see below |
| `section.rename` / `section.rm` | rename a section (`from`, `to`) / remove one (`name`; its slides join the previous section) -- see below |
| `img.prompts` | overlay picture placeholders with visible image-prompt boxes (removable via `el.rm`) |
| `meta.props` | document properties: core fields (title, author, subject, keywords, category, comments) and/or arbitrary `custom` name/value pairs stored in the `.pptx` (read back by `state` -- the deck's self-describing memory); only the named custom properties change (as text), every other one stays as it is, whatever its type |

Run `pptc schema <op>` for the authoritative JSON Schema with all fields, or
`pptc schema document` for the whole ops document.

### `el.add` elements

Every element has a `type`, an optional `name` (addresses it in later
`el.set`/`el.rm`), optional `altText` and `title` (accessibility), an
optional `zOrder` (`front` -- the default: on top of the slide's shapes
--, `back`: behind all of them, `forward`/`backward`: placed in front,
then one shape up/down) and -- except connectors -- a
`frame: {x, y, w, h, rotation}` in inches.

- **`textbox`**: `text` (rich text), `fill`, `transparency`, `border`, `borderPt`, `lang`,
  and the text box layout shared with `shape`: `valign`
  (`top`/`middle`/`bottom`), `margin` (inner margin in inches: one value
  or `[left, top, right, bottom]`), `textDirection` (`horizontal`,
  `rotate90`, `rotate270`, `stacked`); `hyperlink` (URL or slide link:
  the click target of the whole box).
- **`shape`**: `shape` -- a preset geometry: basic `rect`, `roundRect`,
  `round1Rect`, `round2SameRect`, `snip1Rect`, `ellipse`, `diamond`,
  `triangle`, `parallelogram`, `trapezoid`, `pentagon`, `hexagon`,
  `octagon`, `plus`, `donut`, `pie`, `blockArc`, `can`, `frame`, `star5`,
  `cloud`, `heart`, `line`; arrows and steps `rightArrow`, `leftArrow`,
  `upArrow`, `downArrow`, `leftRightArrow`, `upDownArrow`,
  `notchedRightArrow`, `curvedRightArrow`, `circularArrow`, `homePlate`,
  `chevron`; callouts and brackets `wedgeRectCallout`,
  `wedgeRoundRectCallout`, `wedgeEllipseCallout`, `leftBracket`,
  `rightBracket`, `leftBrace`, `rightBrace`; flowchart
  `flowChartProcess`, `flowChartDecision`, `flowChartTerminator`,
  `flowChartDocument`, `flowChartPredefinedProcess`,
  `flowChartConnector` (`state` reads it back as `shape`) --, `text`
  (rich text), `fill`, `transparency`, `border`,
  `borderPt`, `fontColor`, `fontSize`, `lang`, `valign`, `margin`,
  `textDirection`, `hyperlink` (e.g. a back button to `slide:title:Agenda`).
- **`table`**: `data: {headers, rows, style}` (`style`: `headerBg`,
  `headerFg`, `altRowBg`, `fontSize`, `border`, `headerHeight`,
  `rowHeight`); `colWidths` (inches per column, as `state` reports them;
  scaled to `frame.w` when that is given), `colAlign`
  (`left`/`center`/`right` per column), `textDirection` (all cells),
  `lang`. A cell (also a header cell) is a plain value or `{text, bold,
  fill, transparency, color, align, valign, colspan, rowspan, textDirection}` with rich
  `text`; a cell covered by a `colspan`/`rowspan` is left out of its row,
  as in HTML. The header row is bold, white on dark gray (`333333`) unless
  `headerFg`/`headerBg` say otherwise (a `headerBg` without `headerFg`:
  `dk1` or `lt1` by contrast), and is the table's header row (PowerPoint's
  *Header Row*, `state` `headerRow`); body text without a `color` takes
  the template's, on a cell fill or `altRowBg` `dk1` or `lt1` by contrast.
  Rows that do not fill the grid, and `colWidths`/`colAlign`
  of the wrong length, fail with `E_SCHEMA`.
- **`chart`**: `data: {type, categories, series, title, colors, legend,
  fontSize, labels, axes, stacked, smooth, holeSize}`. `type` is `bar`,
  `barStacked`, `column`, `columnStacked`, `line`, `lineMarkers`, `pie`,
  `doughnut`, `area`, `areaStacked`, `radar`, `scatter` or `bubble`
  (scatter/bubble series give `x` and `y` -- the same `x` for every series
  -- and bubble `size`). `legend`: `true`/`false` or `top`/`bottom`/
  `left`/`right`. `labels: {value, percent, format, position}` shows data
  labels (values unless only `percent` is asked; percentages on
  pie/doughnut only); `position` is one of `center`, `insideEnd`,
  `insideBase`, `outsideEnd` (bar/column; stacked ones not `outsideEnd`),
  `left`, `right`, `above`, `below` (line/scatter/bubble), `center`,
  `insideEnd`, `outsideEnd`, `bestFit` (pie) -- doughnut, area and radar
  take none. `axes: {x, y}` with `title`, `min`, `max`, `format` (number
  format like `"0%"`); `x` is the category axis (titles only), the x value
  axis on scatter/bubble; pie/doughnut have no axes. `stacked: "percent"`
  stacks bar/column/area series (`area` and `areaStacked`) to 100%, `smooth` smooths line charts,
  `holeSize` (10-90) sizes a doughnut's hole. Unsupported combinations fail
  with `E_SCHEMA`.
- **`image`**: `path`, `fit` (`stretch` -- the default --, `crop`:
  centered cover, `contain`: the whole image in a centered, shrunk frame),
  `hyperlink` (URL or slide link).
- **`connector`**: `from`/`to` `[x, y]`, `color`, `widthPt`, `dash`
  (`solid`, `dash`, `dot`, `dashDot`, `longDash`), `startArrow`/`endArrow`
  (`none`, `arrow`, `triangle`, `stealth`, `diamond`, `oval`) at the
  `from`/`to` end.

The overlap lint (`W_ELEMENT_OVERLAP`) takes the box a rotated element
covers, and leaves out an element sent to the back (`zOrder: "back"`,
by `el.add` or `el.set`), which lies behind every shape: it is not
checked, and the elements after it may cover it (text on a background
panel) -- until a later move takes it out of the back; the capacity of a placeholder with vertical text (rotated by 90 or
270 degrees in the layout) swaps its width and height, and a `\v` line
break counts as a line.

### `el.set` fields

`el.set` changes a top-level element in place and needs at least one
field; each is checked against the element's type (`E_SCHEMA` otherwise):

| Field | Elements | Effect |
|---|---|---|
| `text` | placeholder, textbox, shape | new rich text in the first paragraph's and run's format (alignment, bullets, font, size, color, language): a paragraph or run that sets fields of its own changes only those |
| `frame` | every element (rotation: not table/chart) | `{x, y, w, h, rotation}` -- the given keys move, resize or rotate it, the others stay; a placeholder that inherits its frame gets the layout's frame written into the slide; a group keeps its children's coordinates (they scale with it); a table's columns and rows scale to the new size |
| `fill` | placeholder, textbox, shape | solid fill; keeps the element's transparency and its text color -- give `fontColor` with a fill the text must contrast with |
| `transparency` | placeholder, textbox, shape | transparency of the solid fill in percent (`0` opaque - `100` invisible): of the new `fill`, else of the element's own (one without a solid fill of its own needs `fill`: `E_SCHEMA`) |
| `border`, `borderPt` | placeholder, textbox, shape, picture, connector | outline color and width in pt |
| `fontColor` | placeholder, textbox, shape | color of every run |
| `altText`, `title` | every element | alt text (`cNvPr@descr`) and title; `""` removes them |
| `hyperlink` | placeholder, textbox, shape, picture | click target of the whole element: a URL or a slide link (`slide:SEL`); `""` removes it |
| `image` | picture, filled picture placeholder | another image in place: the frame and animations stay, the image is cropped centered to fit the frame (like a picture placeholder fill; an image of unknown size stretches); an SVG shows through PowerPoint's SVG extension |
| `cells` | table | `[{row, col, text}]` -- 0-based grid position, header row 0; a cell covered by a merge fails |
| `addRows` | table | `{at?, rows}` -- rows inserted at `at` (default: appended) in the format of the neighbouring body row (the row above; at 1, below the header row 0, the first body row; the first row at 0 -- not above a header row: `E_SCHEMA`); one value per column |
| `removeRows` | table | `[indices]` -- rows taken out; a row of a vertical merge fails; removing row 0 of a table with a header row leaves it without one |
| `zOrder` | every element | stacking order: `front`/`back` to the top/bottom of the slide's shapes, `forward`/`backward` one shape up/down; the moves follow the op order, after the elements the document adds are placed |

Table fields apply in the order `removeRows`, `addRows`, `cells`, each seeing
the result of the one before; the table style and the cell formatting
stay, and the frame's height follows the rows. An element generated or a
picture placeholder filled earlier in the same document is changed in its
plan instead (the `el.add` vocabulary: a planned connector takes `border`/
`borderPt` as its `color`/`widthPt` but no `hyperlink`, a planned image no
`border`/`borderPt`, a planned textbox `fontColor` on its runs). Every `el.add` element also takes `altText` and `title`; a picture
is never described by its file path. Replacing an image keeps the old alt
text: give `altText` together with `image`. Replacing an image destroys the old
one in the deck -- the skill does it only on the user's explicit request.

### `text.replace`

`{ "op": "text.replace", "find": "2025", "replace": "2026", "slide": SEL,
"matchCase": false, "wholeWord": false, "notes": false }` replaces plain
text in every shape, placeholder and table cell of the deck -- or of
`slide` only --, in the speaker notes with `notes: true`, never in fields
(slide number, date) and never in charts or SmartArt: a targeted slide
whose chart (its titles, series names and categories -- not its numbers)
or SmartArt holds the text raises `W_REPLACE_SKIPPED`. A match
may span runs (text split by formatting):
the first touched run takes the replacement and keeps its format. The op
sees the deck after every earlier op of the document (text set or
elements added before it are replaced, later ones are not), and slide
`title:` selectors after it see the new titles. Zero matches is no error.
The apply result reports every `text.replace` in `result.replacements`
(not under `--dry-run`): `{ "op": 2, "count": 5, "slides": [ { "id": 257,
"index": 1, "count": 5 } ] }` -- `op` is its position in the document,
the counts are per output slide (a copy made after the op counts its own).

### Sections

PowerPoint sections group consecutive slides. The fill field `section:
"Name"` on `slide.add`/`slide.fill` starts a section at that slide
(created if missing): it runs up to the next section start, so a slide
inside a section splits it (a section left without slides goes; a slide
already in a section of that name stays). In a deck without sections the
slides before the first one form `"Default Section"` -- as PowerPoint
does; rename it with `section.rename`. `{ "op": "section.rename", "from":
"Teil 2", "to": "Ergebnis" }` renames a section, `{ "op": "section.rm",
"name": "Teil 2" }` removes one: its slides stay and join the previous
section (the next one when it was the first; removing the last section
leaves the deck without sections). Names are unique: a name in use fails
with `E_SCHEMA`, an unknown one with `E_ADDR_NOTFOUND`. Added, copied and
moved slides join the section of the slide before them (the first section
when they lead the deck). The ops compose in document order; `state`
reports each slide's `section` (levels `text` and `full`) and the section
names in deck order (`sections`).

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
  "result": { "applied": 6, "slides": { "intro": { "id": 263, "index": 0 } },
              "replacements": [ { "op": 5, "count": 2, "slides": [ { "id": 263, "index": 0, "count": 2 } ] } ] },
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
PPTX work is split across **three parts**:

1.  **The in-place engine** (`engine/inplace/`) -- edits the deck package
    directly: slides the plan does not touch are never re-imported, renamed
    or re-serialized (their parts keep every byte), new slides are created
    from the deck's own layouts, copies get their own notes, charts,
    embeddings, SmartArt, tags and comments, and every fresh part name and
    id continues above everything in use.
2.  **[PptxGenJS](https://github.com/gitbrent/PptxGenJS)** -- element
    generation only: tables, charts, textboxes, shapes and images created
    via `el.add` are rendered by PptxGenJS and transplanted into their slide
    together with their parts.
3.  **The package-wide steps** (`engine/post.ts`) -- garbage collection,
    sections, content types, document and app properties; deterministic
    repairs apply to touched parts only.

The principle: an own, small engine for the package operations, libraries
only as generators -- every apply changes exactly what the plan asks for,
self-checked before the atomic write.

## Design Decisions

- **Kernel CLI.** Reads are commands, writes are data. An agent is better at
  producing one schema-validated JSON document than at sequencing a dozen
  correct shell invocations -- one validation pass surfaces all errors at
  once, and there is exactly one write schema to maintain.
- **Slides from the deck's own layouts.** `slide.add` builds the new slide
  from the deck's embedded layout (placeholders cloned, named
  `PptcPh-<idx>`), so it keeps working on an existing deck with no side
  file -- a deck is self-contained. `--template` on `apply` only checks
  that the template's layouts equal the deck's.
- **Two-phase apply.** Every op is a pure plan transformer; the engine session
  interprets the finished plan. `--dry-run` is "plan without commit", not a
  separate code path.
- **No silent first-match.** Ambiguous titles, ambiguous placeholder keys and
  unknown refs are errors with candidate lists -- determinism beats
  convenience.

See [ARCHITECTURE.md](ARCHITECTURE.md) for the layer model and data flow, and
[DEPLOY.md](DEPLOY.md) for building and delivering releases.

## Development

```
npm install        # dependencies
npm test           # build + vitest (unit, golden, integration, contract)
npm run lint       # eslint (incl. TSDoc) + tsc --noEmit
npm run build      # esbuild bundle -> dst/pptc.mjs
npm run plugin:sync # rebuild + copy the bundle into the plugin (+ both VERSION files)
```

## Packaging & Distribution

The plugin marketplace in this repo is the single distribution channel
(Claude Code and the Claude app install from it directly); the pptc CLI is
additionally published to npm. Per-skill ZIPs were retired in 1.0.6 -- the
skills share `plugin/meta/control.md`, which a stand-alone skill folder
cannot carry. In-house distributions that bundle company templates live in
a downstream company repo that vendors this plugin and overlays its own
templates.

Templates are richer with a sidecar: put `<name>.md` next to `<name>.potx`
(layout-role map, footer pattern, design constraints) and it is picked up
alongside, so `tpl describe`/the skill get the template-specific
knowledge automatically.

### Versioning & release rule

`package.json` holds the single version (CLI and skills); each skill carries a
`VERSION` file written by `plugin:sync`. The plugin version of a company
marketplace is owned by that marketplace. `pptc` has no update check or
self-update: updates reach users through the marketplace, a skill ZIP or
`npm install -g @brusdeylins/pptc`. See [DEPLOY.md](DEPLOY.md) for the release
checklist. Lint and the test suite must be green before any push.

## License

MIT -- Copyright (c) 2026 Matthias Brusdeylins
