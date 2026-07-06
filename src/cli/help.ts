/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  cli/help: the detailed per-command help registry. `pptc help <topic>`
**  and `pptc <command> --help` render these texts; the same source feeds
**  humans (--plain) and agents (envelope).
*/

/**  detailed help text per command or topic  */
export const HELP: Record<string, string> = {

    "state": `pptc state <deck> [--slide SEL] [--level summary|text|full] [--plain]

Read model of a deck: slides, contents and the 'rev' token for
optimistic locking (read-before-write: pass it to 'apply --rev'). The
result also carries 'customProps' -- the deck's custom document
properties (docProps/custom.xml), the self-describing memory a skill
stores inside the file (e.g. the chosen image style).

Options:
  --slide SEL    restrict to one slide (see 'pptc help selectors')
  --level        summary: ids, indices, titles, layout indices, and
                   'hidden: true' on hidden slides ("Hide Slide"; reported
                   on every level, '[hidden]' under --plain)
                 text (default): plus placeholder texts and notes
                 full: plus every shape with geometry AND styling --
                   type, name, frame, text, table cells + column widths,
                   autoshape preset/fill/border/font, a picture's media file
                   name (image), and -- for a body with explicit run
                   formatting -- its paragraphs/runs (per-run font, size,
                   bold, italic, color), which pass straight back into
                   slide.fill. This is the COMPLETE structure: enough to
                   recreate or edit a styled table, diagram or code block, so
                   there is never a need to unzip the .pptx and read raw XML.
  --plain        readable text instead of the JSON envelope

Example:
  pptc state deck.pptx --level full --slide title:Agenda`,

    "verify": `pptc verify <deck> [--strict]

Check a finished deck against every known PowerPoint "repair" trigger
(dangling relationships, stale docProps/app.xml, shared notesSlides,
broken content types, schema-order and reference errors). 'apply'
already runs this on its own output before writing, so a freshly built
deck is clean by construction; use 'verify' as an explicit gate after a
build, or on any .pptx whose origin you do not control.

Options:
  --strict       any finding is a hard failure (exit 8) instead of a
                 successful report with findings in the result

Result:
  result.ok        true when the deck has no findings
  result.findings  human-readable list of repair triggers (empty = clean)

Example:
  pptc verify deck.pptx --strict`,

    "new": `pptc new <deck> --template <tpl> [--force] [--ops @file] [--strict]

Create a valid, zero-slide deck carrying the template's masters,
layouts and theme. The template's example slides, notes and sections
are stripped; the deck opens in normal view.

Options:
  --template <tpl>  the .potx/.pptx to derive the deck from (required)
  --force           overwrite an existing file
  --ops @file       build the deck in the same run (see 'pptc help ops')
  --strict          lint warnings become errors (exit 7)

Example:
  pptc new deck.pptx --template corporate.potx --ops @build.json`,

    "apply": `pptc apply <deck> (--ops @file|- | -e '<op>') [options]

THE single write path. The whole ops document is validated and planned
first (schema, every selector, capacity lint); then everything applies
in one atomic write -- or nothing ('failedAt' names the failing op and
the deck stays byte-identical).

Options:
  --ops @file|-     ops document from file or stdin ('pptc help ops')
  -e '<op-json>'    exactly one inline op instead of --ops
  --template <tpl>  layout source for 'slide.add'. OPTIONAL when adding to an
                    existing deck -- it reuses the deck's OWN embedded layouts
                    (a deck is self-contained); supply one only to introduce
                    layouts the deck does not already carry
  --dry-run         validate and plan only, no write; warnings included
  --strict          lint warnings become exit 7: W_TEXT_OVERFLOW
                    (shorten/split text), W_ELEMENT_OVERLAP (an el.add
                    element covers a text shape -- reposition it),
                    W_FONT_TOO_SMALL (a run/element font is below the
                    readable minimum -- enlarge it)
  --min-font-pt N   readability floor for explicit font sizes (default
                    11); runs/elements below N trigger W_FONT_TOO_SMALL.
                    0 disables the check. Footer/slide-number/date
                    placeholders and prompt boxes are always exempt
  --rev R           optimistic lock: fail with exit 6 unless the deck
                    still has revision R (from 'pptc state')
  --out F           write the result to a new file, keep the input

Example:
  pptc apply deck.pptx --ops @ops.json --template corporate.potx --rev 27d7f5a4ea5a --strict`,

    "text": `pptc text <deck> --slide SEL [--ph KEY] [--append] "text"

Set placeholder text on one slide (compiles to one slide.fill op).

Options:
  --slide SEL    target slide (see 'pptc help selectors')
  --ph KEY       placeholder key, default 'title'
                 (OOXML idx like "13" or semantic: title, subtitle,
                 body, image, text:13, image:14)
  --append       append to the existing content instead of replacing
  --rev R / --strict / --dry-run   as in 'apply'

Example:
  pptc text deck.pptx --slide title:Agenda --ph body "Fourth point" --append`,

    "note": `pptc note <deck> --slide SEL "speaker notes"

Set a slide's speaker notes (compiles to one slide.fill op).
Options: --slide SEL, --rev R, --strict, --dry-run (as in 'apply').

Example:
  pptc note deck.pptx --slide id:257 "Key message, figures, transition."`,

    "footer": `pptc footer <deck> [--slide SEL] "footer text"

Set the footer of one slide -- or, without --slide, of EVERY slide.
Clones the layout's footer placeholder; layouts without one (typically
title/closing layouts) are skipped silently.
Options: --slide SEL, --rev R, --strict, --dry-run (as in 'apply').

Example:
  pptc footer deck.pptx "© Acme Corp | My Talk | 2026"`,

    "rm": `pptc rm <deck> --slide SEL

Remove a slide (compiles to one slide.rm op).
Options: --slide SEL, --rev R, --strict, --dry-run (as in 'apply').

Example:
  pptc rm deck.pptx --slide index:3`,

    "move": `pptc move <deck> --slide SEL --to N

Move a slide to zero-based position N (compiles to one slide.move op).
Options: --slide SEL, --to N, --rev R, --strict, --dry-run.

Example:
  pptc move deck.pptx --slide title:Agenda --to 1`,

    "schema": `pptc schema [op|document]

JSON Schema of one op (e.g. 'pptc schema slide.fill') or of the whole
ops document ('pptc schema document') -- generated from the validating
Zod schemas, so it is always authoritative. Without argument: the list
of op names. Use this to discover every field an op accepts.

Example:
  pptc schema el.add`,

    "update": `pptc update

Self-update via npm. pptc checks the npm registry at most once per day;
when a newer version exists, every envelope carries an 'update' field.`,

    "help": `pptc help [command|ops|selectors] [--plain]

Detailed help for a command or topic. 'pptc <command> --help' is the
shortcut and always prints plain text.`,

    "tpl list": `pptc tpl list <dir> [--plain]

Inventory of all .potx/.pptx files in a directory, each with a
'sidecar' flag: a Markdown file next to the template (<name>.md)
carrying template-specific notes, included by 'tpl describe'.

Example:
  pptc tpl list ./templates --plain`,

    "tpl describe": `pptc tpl describe <tpl> [--layout SEL] [--format text|json] [--plain]

The LLM-facing template description: per layout an ASCII minimap,
semantic positions, text capacities (~N lines of ~M chars), image
aspect ratios, overlay warnings for picture placeholders ("overlaid
by ... keep these regions calm in images") and a suitability hint --
derived generically from the OOXML geometry. A sidecar <tpl>.md is
included verbatim.

Options:
  --layout SEL    restrict to one layout (zero-based index or exact name)
  --format json   raw TemplateInfo data instead of Markdown
  --plain         print the Markdown directly (human console)

Example:
  pptc tpl describe corporate.potx --plain`,

    "tpl inspect": `pptc tpl inspect <tpl> [--layout SEL]

The precise machine model: slide size, theme fonts, the full theme
color map (dk1, lt1, accent1..accent6 as RRGGBB) and per layout every
placeholder with OOXML idx, kind, name, frame (inches) and capacity.
Picture placeholders carry 'overlays': which text shapes sit on top
and where (region relative to the image) -- image prompts keep those
regions as negative space.

Example:
  pptc tpl inspect corporate.potx --layout TWO_COLUMN`,

    "tpl validate": `pptc tpl validate <tpl> [--plain]

Check a template against pptc's expectations (layouts present, notes
master, unique placeholder names, resolvable geometry). 'fail'-grade
issues exit 7; warnings are informational.

Example:
  pptc tpl validate corporate.potx --plain`,

    "ops": `THE OPS DOCUMENT -- pptc's write API

All mutations are expressed as ONE JSON document and applied atomically
by 'apply' (or 'new --ops'). Structure:

  {
    "expectRev": "27d7f5a4ea5a",     // optional optimistic lock (from 'state')
    "ops": [ <op>, <op>, ... ]       // executed in order, all-or-nothing
  }

Every <op> is an object with an "op" discriminator. The vocabulary:

  slide.add    add a slide from a template layout
               { "op": "slide.add", "layout": 0 | "LayoutName",
                 "ref": "intro",            // optional doc-local name ($intro)
                 "at": 0,                   // optional insert position
                 "placeholders": { "title": { "text": "..." },
                                   "body":  { "text": "A\\nB" },
                                   "image": { "image": "photo.png" } },
                 "notes": "...", "footer": "...",
                 "background": { "color": "1F4E79" },
                 "hidden": true }           // optional: hide/show the slide
  slide.fill   same fill payload, on an existing slide ("slide": SEL)
  slide.rm     { "op": "slide.rm",   "slide": SEL }
  slide.move   { "op": "slide.move", "slide": SEL, "to": 2 }
  slide.copy   { "op": "slide.copy", "slide": SEL, "ref": "copy1" }
  el.add       free elements: textbox | table | chart | shape | image |
               connector -- all positioned via "frame": {x,y,w,h} inches
  el.set       retext an element:  { "op": "el.set", "slide": SEL,
                 "name": "Box1", "text": "new text" }
  el.rm        remove an element by name
  img.prompts  overlay picture placeholders with visible prompt boxes
               { "op": "img.prompts", "slide": SEL, "prompts": "..." }
  meta.props   document properties: core fields (title, author, subject,
               keywords, category, comments) and/or arbitrary 'custom'
               name/value pairs -- stored in the .pptx (docProps/custom.xml),
               so they travel with the file and 'state' reads them back
               { "op": "meta.props", "set": { "title": "Q3",
                 "custom": { "pptcImageStyle": "Pencil Sketch" } } }

Text values accept plain strings (\\n = new paragraph) or rich text
(runs with bold/italic/color/size, paragraphs with bullet levels).
Field-level reference: 'pptc schema <op>' is authoritative.

Recommended agent flow: state -> build ops.json -> apply --dry-run
--strict -> fix findings -> apply --rev <rev>.`,

    "selectors": `SLIDE SELECTORS -- how ops and commands address slides

  id:257        canonical OOXML sldId; stable across reorderings
  title:Agenda  exact slide title; must be unique in the deck
  $intro        doc-local ref created by an earlier op in the SAME
                ops document (slide.add/slide.copy "ref")
  index:2       zero-based position; escape hatch only -- positions
                shift when users edit between turns
  2             bare digits = index:2

Placeholder keys in fills: the OOXML idx ("13") or semantic keys
resolved against the layout: "title", "subtitle", "body", "image",
"text:13", "image:14". Ambiguity is an error, never first-match.`
}

/**  resolve a help topic from tokens like ["tpl","describe"] or ["apply"]  */
export const helpFor = (tokens: string[]): string | null => {
    const key = tokens[0] === "tpl" && tokens.length > 1
        ? `tpl ${tokens[1]}`
        : tokens[0] ?? ""
    return HELP[key] ?? null
}
