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
                 text (default): plus placeholder texts and notes, and
                   each slide's 'section' (with the deck's 'sections' in
                   order) when the deck has sections
                 full: plus every shape, back to front (the stacking
                   order), with geometry AND styling --
                   type, name, frame, text, table cells + column widths
                   (colWidths) + merged cells (merges) + vertical cell
                   text (cellTextDirections) + header row (headerRow),
                   autoshape preset/fill (with its transparency)/
                   border/font, textDirection, valign and margin of a
                   text body, rotation, altText and title, the click
                   target (hyperlink: URL or slide:id:N), connector
                   startArrow/endArrow, the text language (lang), a
                   picture's media file name (image), and -- for a body
                   with explicit run formatting or links -- its
                   paragraphs/runs (per-run font, size, bold, italic,
                   strike, super-/subscript, color, highlight, hyperlink;
                   per paragraph bullet, spacing), which pass straight
                   back into slide.fill. Colors read as RRGGBB, theme
                   colors resolved. This is the COMPLETE structure: enough
                   to recreate or edit a styled table, diagram or code
                   block, so there is never a need to unzip the .pptx and
                   read raw XML.
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
  --min-font-pt N   readability floor for explicit font sizes (default 8,
                    0 disables) -- see 'pptc help apply'

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
  --template <tpl>  check only: 'slide.add' always uses the deck's OWN
                    embedded layouts (a deck is self-contained); a template
                    whose layouts differ from the deck's fails with
                    E_TEMPLATE
  --dry-run         validate and plan only, no write; warnings included
  --strict          lint warnings become exit 7: W_TEXT_OVERFLOW
                    (shorten/split text), W_ELEMENT_OVERLAP (an el.add
                    or moved element covers a text-bearing shape --
                    reposition it; see 'pptc help ops' for what counts),
                    W_FONT_TOO_SMALL (a run/element font is below the
                    readable minimum -- enlarge it). Not escalated, as
                    they are found during the write: W_REPAIRED (the input
                    deck carried PowerPoint repair triggers, which pptc
                    repaired -- tell the user), W_INPUT_DEFECT (the input
                    deck carries repair triggers pptc cannot repair; they
                    stay as they were -- tell the user), W_REPLACE_SKIPPED
                    (text.replace found its text in a chart or SmartArt,
                    which it leaves as it is -- tell the user)
  --min-font-pt N   readability floor for explicit font sizes (default
                    8); runs/elements below N trigger W_FONT_TOO_SMALL.
                    0 disables the check. Footer/slide-number/date
                    placeholders and prompt boxes are always exempt.
                    A pure legibility backstop -- sizes themselves come
                    from the template ('tpl inspect'), never from pptc
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

Set a slide's speaker notes (compiles to one slide.fill op). An existing
notes page keeps everything but its notes text; a new one is built from
the notes master (slide image, notes body, slide number).
Options: --slide SEL, --rev R, --strict, --dry-run (as in 'apply').

Example:
  pptc note deck.pptx --slide id:257 "Key message, figures, transition."`,

    "footer": `pptc footer <deck> [--slide SEL] ["footer text"] [--number] [--date]

Set the footer of one slide -- or, without --slide, of EVERY slide.
Clones the layout's footer placeholder; layouts without one (typically
title/closing layouts) are skipped silently -- a footer the slide has
of its own then takes the text. --number and --date show the slide
number and the date as well (the layout's placeholders with their
fields); with one of them the text may be omitted.
Options: --slide SEL, --number, --date, --rev R, --strict, --dry-run.

Example:
  pptc footer deck.pptx "© Acme Corp | My Talk | 2026" --number`,

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
                                   "image": { "image": "photo.png",
                                              "fit": "crop" } },  // crop|contain|stretch
                 "notes": "...", "footer": "...",
                 "slideNumber": true, "date": false,  // show/remove (true on
                                            // a layout without one: the slide
                                            // keeps its own, if any)
                 "background": { "color": "1F4E79" },
                 "hidden": true,            // optional: hide/show the slide
                 "section": "Teil 2" }      // start a section at the slide
  slide.fill   same fill payload, on an existing slide ("slide": SEL)
  slide.rm     { "op": "slide.rm",   "slide": SEL }
  slide.move   { "op": "slide.move", "slide": SEL, "to": 2 }
  slide.copy   { "op": "slide.copy", "slide": SEL, "ref": "copy1" }
  el.add       free elements: textbox | table | chart | shape | image |
               connector -- all positioned via "frame": {x,y,w,h} inches
               { "op": "el.add", "slide": SEL, "elements": [ ... ] }
               textbox/shape: "text", "fill", "transparency" (of the
                 fill, 0-100 %), "border", "borderPt",
                 "valign": top|middle|bottom, "margin": 0.1 | [l,t,r,b],
                 "textDirection": horizontal|rotate90|rotate270|stacked
                 (shape also "shape", "fontColor", "fontSize")
               shape presets: rect roundRect round1Rect round2SameRect
                 snip1Rect ellipse diamond triangle parallelogram
                 trapezoid pentagon hexagon octagon plus donut pie
                 blockArc can frame star5 cloud heart line rightArrow
                 leftArrow upArrow downArrow leftRightArrow upDownArrow
                 notchedRightArrow curvedRightArrow circularArrow
                 homePlate chevron wedgeRectCallout wedgeRoundRectCallout
                 wedgeEllipseCallout leftBracket rightBracket leftBrace
                 rightBrace flowChartProcess flowChartDecision
                 flowChartTerminator flowChartDocument
                 flowChartPredefinedProcess flowChartConnector
               table: "data": { "headers", "rows", "style" }, "colWidths":
                 [1,2.5], "colAlign": ["left","right"], "textDirection";
                 a cell is a value or { "text", "bold", "fill",
                 "transparency", "color", "align", "valign", "colspan",
                 "rowspan", "textDirection" }
                 -- cells covered by a span are left out (as in HTML);
                 "headers" make the first row the header row
               chart: "data": { "type", "categories", "series", "title",
                 "colors", "legend": true|top|bottom|left|right,
                 "labels": { "value", "percent", "format", "position" },
                 "axes": { "x": { "title", "min", "max", "format" }, "y" },
                 "stacked": "percent", "smooth", "holeSize" }
               image: "path", "fit": stretch|crop|contain, "hyperlink"
               textbox/shape/image: "hyperlink" -- the click target of the
                 whole element (URL or slide link, e.g. a back button)
               connector: "from": [x,y], "to": [x,y], "color", "widthPt",
                 "dash": solid|dash|dot|dashDot|longDash,
                 "startArrow"/"endArrow": none|arrow|triangle|stealth|
                 diamond|oval
               every element also takes "altText", "title" and
                 "zOrder": front (default) | back | forward | backward
                 (placed in front, then moved; "back" lies behind every
                 shape: no W_ELEMENT_OVERLAP for it or for the elements
                 placed over it later in the document)
               W_ELEMENT_OVERLAP obstacles: text placeholders, the footer
                 band, textboxes/shapes with text, tables, charts and
                 groups -- never pictures, picture placeholders, lines,
                 prompt boxes or textboxes/shapes without text (a panel,
                 also in later applies); a panel el.set gives text (not
                 at the back) warns for each element lying over it
  el.set       change an element in place (at least one field, each
               checked against the element's type):
               { "op": "el.set", "slide": SEL, "name": "Box1", ... }
                 "text" (placeholder/textbox/shape; in the first
                   paragraph's and run's format, merged with the fields
                   a paragraph or run sets),
                 "frame": { "x", "y", "w", "h", "rotation" } (given keys
                   change; not rotation on table/chart; "w"/"h" 0 only
                   on a line/connector),
                 "fill", "fontColor" (placeholder/textbox/shape),
                 "transparency" (placeholder/textbox/shape: 0-100 % of
                   the new fill, else of the element's own solid fill;
                   a new fill keeps it),
                 "border", "borderPt" (also picture/connector),
                 "altText", "title" (any element; "" removes),
                 "hyperlink" (placeholder/textbox/shape/picture: URL or
                   slide link; "" removes),
                 "image" (picture or filled picture placeholder: frame
                   and animations stay, the image is cropped centered to
                   fit the frame -- a contain-fitted picture keeps its
                   shrunk frame, give "frame" with it -- only on explicit
                   request),
                 table: "cells": [{ "row", "col", "text" }] (0-based,
                   header row 0), "addRows": { "at", "rows" } (format of
                   the neighbouring body row; never above a header row),
                   "removeRows": [ 2 ],
                 "zOrder": front|back|forward|backward (any element; in
                   op order, after the document's new elements)
               -- also on elements planned earlier in the document (a
                  planned connector takes no frame and no hyperlink, a
                  planned image no border). Ops on one element compose
                  in op order: el.rm drops its earlier edits, a later
                  slide.fill text replaces an earlier el.set text (an
                  append to it: E_SCHEMA), an el.set on a footer, slide-
                  number or date placeholder holds when slide.fill sets
                  that field in the document (false removes it)
  el.rm        remove an element by name
  img.prompts  overlay picture placeholders with visible prompt boxes
               { "op": "img.prompts", "slide": SEL, "prompts": "..." }
  meta.props   document properties: core fields (title, author, subject,
               keywords, category, comments) and/or arbitrary 'custom'
               name/value pairs -- stored in the .pptx (docProps/custom.xml),
               so they travel with the file and 'state' reads them back
               { "op": "meta.props", "set": { "title": "Q3",
                 "custom": { "pptcImageStyle": "Pencil Sketch" } } }
  text.replace find/replace plain text (a case-insensitive substring
               match by default, within one paragraph: no line break in
               "find"/"replace") in shapes, placeholders and table
               cells of the deck (or of "slide"), notes only on request,
               fields, charts and SmartArt never (W_REPLACE_SKIPPED when
               a targeted slide's chart text -- titles, series names,
               categories -- or SmartArt holds the text); a
               match may span runs (the first keeps its format); sees
               every earlier op of the document
               { "op": "text.replace", "find": "2025", "replace": "2026",
                 "slide": SEL, "matchCase": false, "wholeWord": false,
                 "notes": false }
               -> result.replacements: [{ "op", "count", "slides" }]
                  (absent under --dry-run: no preview of the counts)
  section.rename { "op": "section.rename", "from": "Teil 2", "to": "Ergebnis" }
  section.rm   { "op": "section.rm", "name": "Teil 2" } -- the slides stay
               and join the previous section (the next one for the first)

Sections: the fill field "section" starts a section at the slide (created
if missing); it runs up to the next section start, so a slide inside a
section splits it, and a section left without slides goes. In a deck
without sections the slides before it form "Default Section" (as in
PowerPoint). Names are unique. Added, copied and moved slides join the
section of the slide before them. 'state' shows "section" per slide.

Text values accept plain strings (\\n = new paragraph, \\v = line break
inside the paragraph; state reads one back as \\v) or rich text:
paragraphs { "text" | "runs", "level", "alignment", "bullet": true|
"number"|false (true: the placeholder's own bullet where its level
shows one, elsewhere a bullet character with hanging indent),
"spaceBefore"/"spaceAfter" (pt),
"lineSpacing" (1.5) }, runs { "text", "bold", "italic", "underline",
"strike", "superscript", "subscript", "size", "font", "color",
"highlight", "hyperlink", "lang" }; "font" also takes the theme fonts
"+mj-lt" (heading) and "+mn-lt" (body).
Links: "hyperlink" takes a URL or a slide link "slide:SEL" -- slide:$ref,
slide:id:256, slide:title:Agenda, slide:index:2 -- resolved after the
whole document (slides added later count, moved ones follow; a target
missing then fails with E_ADDR_NOTFOUND). A later removal of the target
drops the link; copies keep it; 'state --level full' reads slide:id:N.
Colors: hex "1F4E79" or a theme slot accent1-accent6, tx1, tx2, bg1,
bg2, dk1, lt1, dk2, lt2, hlink, folHlink (follows the template; 'state'
shows RRGGBB). Text tx1/tx2, fills and backgrounds bg1/bg2 (they follow
the master's color map); dk1/lt1/dk2/lt2 are absolute (contrast on
images). Text without "color" takes the template's -- on an el.add fill
(textbox, shape, cell, header, alternating row) dk1 or lt1, whichever
contrasts more (a transparent fill as it shows over bg1; el.add table
headers: white on the default gray). An el.set "fill" keeps the text
color: give "fontColor" with it.
Control characters but tab, line feed and carriage return (and \\v
outside text) fail with E_SCHEMA naming the field.
No shrink-to-fit: overflowing text is shortened or split (W_TEXT_OVERFLOW).
Picture fills crop to the placeholder by default ("fit": "crop");
"contain" shows the whole image, "stretch" distorts it.
Text language: every run carries a "lang" -- the deck language
(custom property pptcDeckLang, else the deck's default text language,
else en-US) unless a textbox/shape/table, paragraph or run sets "lang".
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
