# Changelog

## 1.0.4 (plugin 1.0.4)

- **Skill routing: reciprocal NOT-for clauses between `ppt` and `ppt-prepare`**
  (PR #5 by Milofax, reworked -- documentation only). Both skill descriptions
  now state their boundary explicitly: `ppt-prepare` excludes editing an
  EXISTING deck, `ppt` excludes a brand-new deck from just an idea. Unlike the
  PR, the trigger verbs "make" and "build" STAY in `ppt-prepare`'s description
  and `ppt`'s exclusion is stated by STATE, not verb ("when no approved
  content plan exists yet -- even if the user says build or make"), so
  idea-only prompts like "build me a PowerPoint about X" keep routing to
  `ppt-prepare`. The advertised method list shrinks to "Pyramid Principle and
  SCR narrative" (details stay in the reference docs); "slide messages"
  remains in the trigger-object list.
- **Skill docs: German "Fazit" replaced with "conclusion" throughout.** The
  PR renamed it only in `ppt-prepare`'s methodology.md; the `ppt`-side
  consumers of the same term (SKILL.md, prompt-formula.md, content-rules.md)
  are now renamed too, keeping the plan-to-build vocabulary consistent at the
  handoff.
- **Neutral-template label corrected at every site.** The bundled fallback
  carries a legacy Office-style palette (accent1 `4F81BD`), not Microsoft's
  current default theme; the sidecar body (as in the PR), its heading and the
  two claims in `ppt`/SKILL.md STEP 2 now all say "Office-style fallback
  theme" instead of "Microsoft's default Office design".

## 1.0.3 (plugin 1.0.3)

- **`apply` no longer un-hides hidden slides.** Every apply rebuilds the slide
  list through `pptx-automizer`, which drops the `show="0"` attribute (`p:sld`
  root -- PowerPoint's "Hide Slide") on every slide. pptc carried no notion of
  it, so a single edit silently made *all* hidden slides visible. The read model
  now carries `hidden` per slide, and the post-pass re-applies `show="0"` to
  every slide that was hidden, untouched slides included -- so visibility
  round-trips losslessly. `state` reports `hidden: true` on hidden slides (all
  levels; `[hidden]` in `--plain`).
- **`slide.fill` / `slide.add` can hide or show a slide.** A new optional
  `hidden` boolean on the fill payload deliberately sets visibility:
  `"hidden": true` hides ("Hide Slide"), `false` shows, omitted leaves the
  current state untouched. Builds on the visibility preservation above.
- **Skill docs: prose `&` replaced with the word "and".** The `&` character in
  skill markdown caused downstream ampersand-handling problems; the skill text
  now spells it out (English skills -> "and"). Documentation only; the `pptc`
  CLI is unchanged (republished at 1.0.3 to keep engine, plugin and bundle in
  lockstep).
- **Docs and drift-guards from an architecture/docs review.** ARCHITECTURE.md
  corrects its layer diagram (engine depends on core, not the other way
  round), lists the previously missing modules (`engine/verify.ts` among
  them) and documents the self-verify step and slide visibility in the
  `apply` data flow. `pptc help state` now mentions the `hidden` report. The
  docs contract test additionally pins every fill-payload field to README and
  `help ops`, and `engine/post.ts` / `engine/verify.ts` carry an explicit
  lockstep note tying their repair-trigger lists together.

## 1.0.2 (plugin 1.0.2)

- **`ppt` skill: no longer builds silently on the neutral default template.**
  When the only bundled template is the neutral Office fallback, STEP 2 now
  ASKS (via the Asking-the-User procedure) whether to build on that default or
  supply a corporate template -- a `.potx`/`.pptx` path or a directory to
  scan -- instead of using the generic design and merely informing afterwards.
  A single bundled *corporate* template (internal build) is still used
  directly. Skill-only change; the `pptc` CLI is unchanged (republished at
  1.0.2 to keep engine, plugin and bundle in lockstep).

## 1.0.1 (plugin 1.0.1)

- **`state --level full` now returns enough to recreate styled content without
  unzipping the deck.** Two read-model additions:
  - a picture shape carries its media file name (`image`, e.g. `image11.png`),
    resolved from the blip `r:embed` relationship;
  - a text body with explicit run formatting carries its `paragraphs[]` /
    `runs[]` -- per-run `font`, `size`, `bold`, `italic`, `color` -- and the
    structure round-trips straight back into `slide.fill`.

  Plain uniform text stays lean (no `paragraphs`; the flat `text` already says
  everything). This closes the last cases -- monospace code blocks and image
  references -- where an agent had to drop to raw XML, so the `state` help no
  longer over-promises "never a need to unzip".


## 1.0.0 (plugin 1.0.0)

First stable release. No code change over 0.11.3 -- this marks the surface as
settled and the two skills as production-ready. What 1.0 stands for:

- **Deterministic PPTX engine (`pptc`):** one atomic write path, schema-
  validated ops, optimistic locking, and a self-checking post-pass that refuses
  to emit a deck PowerPoint would ask to repair (`verify` is green by
  construction).
- **Self-contained decks:** `slide.add` needs no template on an existing deck
  (it seeds from the deck's own layouts), and the deck carries its own setup as
  custom document properties -- hand a `.pptx` to someone else and the skill
  keeps working, no side files.
- **Two skills, one flow:** `ppt-prepare` plans the story (Pyramid/MECE/SCR,
  headline titles) and hands a single plan file to `ppt`, which builds it
  deterministically and writes color-faithful image prompts.
- **Portable everywhere:** runs in Claude Code and on claude.ai. Gates use the
  native selection box where present and a clean Markdown dialog otherwise;
  the progress task list is host-conditional; skills announce their version and
  check GitHub for updates.
- **Consistent, teachable output:** one clean-Markdown style across every
  user-facing template, and the skills model the control-tag DSL they teach.


## 0.11.3 (plugin 0.11.3)

- **Consistent "clean Markdown" look for all skill output.** One visual style
  (icon + bold label, `·` separator, a thin `────` rule to open a section, `›`
  before a prompt line; no drawn boxes) now governs every user-facing
  template in both skills -- the version banner, the step/phase banner, the
  Asking-the-User dialog, the gate checkpoint (`◆ Checkpoint` with `✅`/`⬜`
  criteria), the `ppt` report (ruled header + 🗂/🎨/🛡/📋 lines) and the image-
  prompt mirror (`🎨 Image prompt · slide … · role …`). The style is documented
  centrally in each `meta/control.md`. Behaviour unchanged; the data plan file
  stays plain for clean parsing.
- **Consistency: `VERSION` at the skill root for both skills.** `ppt`'s
  `VERSION` moved out of `scripts/` to the skill root, matching `ppt-prepare`,
  so both read `<skill-dir>/VERSION` uniformly and `scripts/` holds only the
  bundle. `plugin:sync`, the startup banner paths and the bundle contract test
  updated accordingly.


## 0.11.2 (plugin 0.11.2)

- **Routing: a NEW deck starts with `ppt-prepare`, not `ppt`.** "Create a deck
  about X" was matching `ppt` (which built immediately, skipping the
  story-first planning). The skill descriptions are re-scoped: `ppt-prepare`
  now owns "create / make / plan a new presentation" (the entry point) and
  `ppt` owns "build from an approved plan / edit an existing deck". As a safety
  net, when `ppt` is invoked for a brand-new deck from only an idea/topic (no
  plan), it now OFFERS to plan with `ppt-prepare` first instead of building
  blindly.
- **Fix: the STEP 3 style pick no longer dangles on claude.ai.** Presenting
  the full style catalog ended on a bare colon (no answerable question) where
  no selection box exists. That path now routes through the central "Asking
  the User" procedure and must end with a concrete question ("Which image
  style? Reply with a name from the list above").
- **Control-tag audit: the skills practice what they teach.** Both SKILLs are
  meant to model the control-tag DSL, but several branch chains and case lists
  in the `<flow>` were still written as prose ("If / Else if / Else", "for
  change decks ... for analytical decks", per-placeholder "→" cases). They are
  now expressed with `<if>/<elseif>/<else>` (e.g. ppt STEP 2 template choice,
  STEP 4 routing, STEP 6 conditional ops, STEP 7 placeholder cases;
  ppt-prepare PHASE 1 entry, PHASE 3 sparkline, PHASE 7 notes, PHASE 8
  platform). All "selection box" mentions now route through the single "Asking
  the User" procedure. Behaviour is unchanged; the prompts are just
  consistent with what the deck teaches.


## 0.11.1 (plugin 0.11.1)

- **Fix: gates work on claude.ai (no selection-box tool).** The Stage Gate
  required the host's selection box (AskUserQuestion) and forbade prose, but
  the claude.ai web chat has no such tool -- so a gate ended on a bare colon
  with no answerable question and the flow stalled. Both skills now share one
  central **"Asking the User"** procedure in `meta/control.md` (pattern
  borrowed from ASE's `ase-dialog`): use the selection box where the host has
  one, otherwise render a numbered Markdown dialog, wait, and map the reply
  (number / label / free text) back to an option -- never ending a turn on a
  half-asked question. Every `<gate/>` routes through it. The Progress Task
  List is likewise host-conditional: skipped (not faked as text) where no
  task-list facility exists.


## 0.11.0 (plugin 0.11.0)

- **Skills announce their version and check for updates.** On first activation
  in a conversation, `ppt` and `ppt-prepare` now print a one-line version
  banner (read from their bundled `VERSION`) and run a best-effort check
  against the GitHub releases of `Brusdeylins/ppt-skill`; when a newer release
  exists they add a single "update available" line (npm for the CLI / re-upload
  the ZIP or update the plugin for the skill). The check uses Node's `fetch`
  with a 3s timeout and degrades silently offline -- it never gates or delays
  the work. `ppt-prepare` now carries its own `VERSION` file (written by
  `plugin:sync` alongside `ppt`'s), and the bundle contract test asserts both.


## 0.10.1 (plugin 0.10.1)

- **Fix: `state` now emits `customProps`.** 0.10.0 read custom document
  properties into the model but the `state` command dropped them from its JSON
  envelope, so a skill calling the CLI never saw them -- defeating the
  self-describing-deck feature end to end. The command now includes
  `result.customProps`. (The round-trip test exercised `readDeckState`
  directly and missed the gap; a test now asserts the command envelope.)


## 0.10.0 (plugin 0.10.0)

**Self-contained decks** -- a deck now carries everything a skill needs to keep
working on it, so one person can hand a `.pptx` to another and the skill picks
up seamlessly, with no side files:

- **Seed-from-deck: `slide.add` no longer needs `--template` on an existing
  deck.** `new` already bakes the full template (masters, layouts, theme) into
  every deck and the post-pass never garbage-collects layouts, so pptc now
  derives its seed (one slide per layout) from the deck's OWN embedded layouts.
  A template is needed only to CREATE a deck (`new`) or to introduce a layout
  the deck does not already carry. This also removes a latent mismatch: layout
  addressing already read the deck's layouts, while the seed came from the
  external template -- now both come from the same source.
- **Custom document properties: `meta.props` gains a `custom` map.** Arbitrary
  name/value pairs are written to `docProps/custom.xml` (with the content-type
  override and package relationship wired idempotently, so `verify` stays clean
  and re-applies never duplicate). PowerPoint preserves them across edit/save,
  so they travel inside the file. `state` now returns them as `customProps`.
- **Skills `ppt` + `ppt-prepare`:** `ppt` persists its setup (image style,
  info-graphic style, deck language, title, topic) INTO the deck as custom
  properties (`pptcImageStyle`, `pptcInfoStyle`, `pptcDeckLang`, `pptcTitle`,
  `pptcTopic`) and reads them back from `state` -- the deck is self-describing,
  the separate `<deck>.md` sidecar is no longer required. `slide.add` guidance
  updated: no `--template` needed for an existing deck.
- **`ppt-prepare` hands off a SINGLE file.** The deck setup (language, title,
  topic) now rides in the plan's header, so PHASE 8 writes only
  `<deck>-plan.md` -- the separate seeded sidecar is gone. On claude.ai, where
  each skill runs in its own non-shared sandbox, `ppt` now asks the user to
  attach the plan instead of silently building without it; README documents
  the platform difference.


## 0.9.1 (plugin 0.9.1)

Review fixes to the `ppt-prepare` and `ppt` skill prompts (no engine change):

- `ppt-prepare` PHASE 5 renamed **"Titles" → "Slide Headlines"** and reworded
  throughout: it produces one on-slide **headline per content slide**, NOT the
  single presentation/deck title (that is `ppt`'s deck-setup job). Removes the
  ambiguity that led the skill to set the deck title instead of per-slide
  headings.
- `ppt-prepare` PHASES 7–8: the final plan is now **assembled and written
  VERBATIM** from the approved Phase 5 headlines and Phase 6 content. Assembly
  only — no re-summarizing or re-densifying (densification closes in Phase 4).
  Fixes Phase 8 re-condensing content instead of carrying Phase 6 over.
- `ppt-prepare` PHASE 6 + methodology: removed the **"bullets + image is the
  default"** framing. The six content types (key-message, bullets + image,
  table, chart, code block, SVG graphic) are now chosen by FIT to the message
  — **no type is the default** — fixing the skill's bias toward bullets+image.
- `ppt` STEP 6: **preserve approved wording — do not rewrite the plan.** When a
  ppt-prepare plan or the user supplies a slide's content, it is placed
  verbatim; paraphrasing, summarizing or shortening it (e.g. to fit placeholder
  capacity) now requires **explicit user confirmation** rather than happening
  as a silent build-time side effect.


## 0.9.0 (plugin 0.9.0)

Quality milestone: engine and plugin versions realigned to **0.9.0**. Rolls up
all the work previously tracked here as 0.2.16.


- Skill `ppt`: **show the full style catalog at the STEP 3 style gate.** The
  catalog has 10 image + 8 info-graphic styles, but a selection box lists only
  4 -- the skill now presents EVERY style (with its "Best for" note) as a
  readable message first, then takes the pick via the box ("Other" free-text)
  or free text, instead of silently offering a curated few.
- Skill `ppt`: brought the command reference **up to date with the pptc surface** --
  added `tpl validate` (STEP 2 now validates the chosen template before building,
  exit 7 on a fail-grade issue) and noted that `state --level full` returns table
  geometry/cells/colWidths and autoshape preset/fill/border/font, so existing
  tables and native diagrams can be recreated/edited round-trip without raw XML.
- Packaging: the **internal skill build now REPLACES the neutral default
  template** instead of bundling it alongside. `skill:zip:internal` (and
  `--from`) drops `assets/neutral-template.pptx` when it overlays corporate
  templates, so an in-house package ships **only** your templates (no generic
  Office fallback to pick by accident); the public `skill:zip` still ships the
  neutral default. `ppt` STEP 2 template selection is now count-based (one →
  use it; several → menu), no longer assuming a neutral default is present.
- Skills `ppt` + `ppt-prepare`: **self-study / teaching decks skip speaker
  notes** -- a deck read without a presenter must be self-contained, so the
  explaining text lives ON the slide (Phase 6), not in a notes pane no reader
  opens. `ppt-prepare` PHASE 7 omits the per-slide notes line for the teaching
  genre; `ppt` writes no `notes` for a self-study deck.
- Skill `ppt-prepare`: **teaching / explanatory deck genre** added to the
  methodology (the original was decision-deck-centric). Genre is *derived from
  the presentation type* (a "teaching" type → explanatory deck), NOT a new
  question or phase. The methodology now flags, per phase, where teaching
  diverges: a **learning objective** instead of a decision asked for (CTA
  optional), a **worked-example spine** with rising-complexity staging as an
  alternative to pure SCR, and **preempting known confusions** (disambiguation
  slide / clarification callout) beside the jargon test.
- Skills `ppt` + `ppt-prepare`: **two content capabilities the engine already
  had but the skills never offered** (gap found against the original m4a
  methodology). (1) **Charts** are now a first-class content type -- quantitative
  data (trends, magnitudes, parts of a whole) plans/builds as a native, data-bound
  `el.add` chart instead of a hand-drawn SVG; added to the `ppt-prepare` layout
  vocabulary and methodology, and to `ppt`'s STEP 6 build rules. (2) **Inserting a
  user-provided image** -- `ppt` STEP 7 now has an explicit branch to place an image
  the user supplies (`slide.fill image` / `el.add image`) instead of only writing a
  generation prompt; the "no image generation" non-goal now clarifies insertion is
  allowed.
- **Toolchain bumped to current majors** (dev-only -- no runtime change): TypeScript
  6.0, ESLint 10.5 (now needs `@eslint/js` as an explicit devDependency), `@types/node`
  26, plus vitest 4.1.9 and typescript-eslint 8.61.1 patches. Lint, type-check, build and
  all 73 tests pass. The runtime dependencies (`pptx-automizer`, `pptxgenjs`,
  `@xmldom/xmldom`, `jszip`, `zod`) were already on their latest published versions.
- **`state --level full` introspection** (read-plane completeness): the shape
  model now exposes what was previously only reachable by unzipping the OOXML.
  Tables report their **`frame`** (the graphicFrame's `p:xfrm`, previously
  `null`) and **`colWidths`** alongside the existing cell matrix; autoshapes
  report **`shape`** (preset geometry), **`fill`**, **`border`**, **`borderPt`**,
  **`fontSize`**, **`fontColor`** and **`fontFace`** (theme colors resolved) --
  the read mirror of el.add's write vocabulary. This lets an agent re-create or
  modify a styled table/diagram from `state` alone, without reading raw XML.
- Engine **architecture**: moved the two leaf modules `cli/args.ts` and
  `core/errors.ts` into the `infra` foundation (`infra/args.ts`,
  `infra/errors.ts`). This removes the real `cli` ⇄ `commands` package cycle
  (commands no longer import *up* into `cli` for the arg parser) and the
  `infra` → `core` upward dependency (the error taxonomy now lives in the
  foundation everyone imports *downward*), so the code finally matches the
  strictly-downward layer model documented in `ARCHITECTURE.md`. Pure internal
  relocation -- no CLI, exit-code or ops-schema change. Also tabled the missing
  exit code **8 (`E_INTEGRITY`)** in the README.
- Skills `ppt` + `ppt-prepare`: both `SKILL.md` now **import** their
  `meta/control.md` eagerly via `@${CLAUDE_SKILL_DIR}/meta/control.md` (the
  control-tag conventions were previously only referenced in prose, so they
  could be skipped). README and `plugin/README.md` updated to document **both**
  shipped skills (the plugin was renamed `powerpoint` -> `ppt`, and
  `ppt-prepare` is now shipped, not "future").
- Skills `ppt` + `ppt-prepare`: **control-tag cleanup** so each `meta/control.md`
  is genuinely self-contained (as it claims). Defined the previously
  undocumented `<objective>` tag; added the optional `<step condition="...">`
  attribute (the task-list prose already assumed skippable steps); and added
  `<define>`/`<expand>` for reusable blocks so `<template>` is now **output
  only**. The two `SKILL.md` no longer overload the self-closing `<template/>`
  as a back-reference -- repeated formats use `<define>`/`<expand>`, one-shot
  output stays an inline `<template>`.
- Skill `ppt`: **header cleanup + alignment** with `ppt-prepare` -- the
  copyright moved out of the YAML front-matter into an HTML comment, and the
  front-matter now sets `user-invocable`, `disable-model-invocation`,
  `model: opus` and `effort: high` (a gated/looped skill needs high effort or
  gates get skipped).
- Skills `ppt` + `ppt-prepare`: **prose trimmed for fewer prompt tokens**
  without behavior change -- removed text that the now-imported
  `meta/control.md` already defines, collapsed the repeated
  `node .../pptc.mjs` prefix in the command reference into a single lead-in,
  and tightened verbose rules. No instruction, gate, quality criterion or
  image-prompt rule was dropped.
- Skill `ppt-prepare`: **split the final phase** -- PHASE 7 is now "Speaker
  Notes & Q&A" (produce notes, Q&A and assemble the plan), and the hand-off
  is its own **PHASE 8: Handoff** with a delivery selection box: *Save &
  finish* (the `<deck>-plan.md` is the deliverable) vs. *Save & hand off to
  `ppt`* (also seed the deck sidecar and point the user to the build). Eight
  phases total; the phase-marker table covers PHASE 4–8.
- Skill `ppt-prepare`: new ground rule **"never guess content -- research,
  then ask"**: on an unclear fact or content gap, do not invent it; where web
  research helps, research first and present the findings as the selection-box
  options, never a silent guess.
- Skills `ppt` + `ppt-prepare`: the reference-files rule now just points at the
  `references/` directory and states the **lazy-load** convention (each step's
  `Read …` line pulls in only what it needs) instead of cataloguing every file
  -- the per-step `Read` lines are the single source. The `ppt` style-catalog
  re-read moved from the catalogue into STEP 7 so nothing was lost.

## 0.2.15 (plugin 0.2.20)

- New command **`pptc verify <deck> [--strict]`**: checks a finished deck
  against every known PowerPoint "repair" trigger and reports the findings
  (`--strict` makes any finding exit 8). The validation now SHIPS in the
  engine (`src/engine/verify.ts`), so it travels with the skill bundle and
  runs on any machine -- the skill verifies its own deliverable instead of
  the user discovering corruption in PowerPoint.
- Engine: **`apply` self-verifies its output before writing**. The post-pass
  repairs the known triggers; if any survive, the apply fails atomically with
  the new `E_INTEGRITY` (exit 8) and the deck is left untouched -- a corrupt
  `.pptx` is never written.
- Skill `ppt`: STEP 6 verifies every write and STEP 8 runs a final integrity
  gate; the report now carries an integrity line. The validator is now a
  single source of truth shared by the engine and the test suite, and a new
  **bundle-sync contract test** fails if the skill's embedded engine drifts
  from the built one (forgetting `npm run plugin:sync` no longer ships a stale
  engine to fresh machines).
- Engine **fix (repair prompt)**: `apply` now prunes **dangling `/slide`
  relationships** from `presentation.xml.rels`. pptx-automizer mints a
  fresh `…-created` slide relationship on every re-import but wires only
  one into the `sldIdLst`; the rest piled up pointing at parts that still
  exist (so the existing dead-target sweep missed them) until a
  thousand-strong rels file made PowerPoint demand a repair on open. The
  post-pass now drops every `/slide` rel no `<p:sldId>` references, and the
  integrity validator fails on any such dangling rel.
- Engine **fix (repair prompt)**: the post-pass now enforces a **1:1
  slide ↔ notesSlide** mapping. The pptc notes part name is derived from the
  slide part basename, but automizer renumbers slide parts across applies,
  so a stale notes rel could leave two slides pointing at one notesSlide
  (which carries a single back-reference) -- invalid, and a repair trigger.
  Shared pptc notes parts are now cloned so each slide owns its notes; the
  integrity validator fails on any notesSlide referenced by more than one
  slide.
- Engine **fix (repair prompt)**: the post-pass now keeps
  **`docProps/app.xml`** in sync with the actual slides. automizer writes
  that extended-properties part once and does not grow its slide list on a
  later apply, so a deck that gained slides ended up declaring fewer than it
  had (`<Slides>`, the slide-title `HeadingPairs` count and the
  `TitlesOfParts` list) -- another repair trigger. The slide count, the
  slide-title group and the title list are rewritten from the real slides;
  the integrity validator fails on a stale count or a `HeadingPairs` /
  `TitlesOfParts` size mismatch.

## 0.2.14 (plugin 0.2.19)

- Engine: `tpl describe` now also reports the capacity of the **title**
  placeholder (`~N lines of ~M chars`), so action titles can be written
  against a real budget instead of overflowing silently.
- Engine: the capacity estimate subtracts a **~10% safety buffer**, so
  authored text lands with headroom rather than at the very edge (bold or
  wide title faces overflow before the raw geometric fit).
- Skill `ppt`: `content-rules.md` documents the title-capacity budget and
  the `W_TEXT_OVERFLOW` consequence; both skills keep a **Progress Task
  List** (the flow's phases/steps shown as a live task list).
- New skill **`ppt-prepare`**: a story-first, gated content-planning
  process (7 phases — briefing, core message, storyline, slide messages,
  titles, content & layout, speaker notes) that produces an approved
  per-slide plan and hands it to `ppt`. Its methodology fixes the layout
  vocabulary (key-message = one short sentence; bullets + image; table;
  code block in a monospace face; SVG graphic; tables/SVG on blank
  layouts; closing = 1-2 words).

## 0.2.13 (plugin 0.2.18)

- Engine: the `img.prompts` box header now states the picture
  placeholder's aspect ratio (`IMAGE PROMPT · <ratio>`), derived from the
  geometry; the ratio is no longer written into the prompt text (so it
  cannot nudge the image model) and the redundant `idx` was dropped.
- Skill and plugin renamed `powerpoint` -> `ppt` (to match `pptc`); the
  marketplace stays `ppt-skill`, install is now
  `/plugin install ppt@ppt-skill`. The plugin is a container for multiple
  skills (a content-prep `ppt-prepare` skill is planned).
- The skill discovers its OWN bundled templates: when no template is named
  it scans `assets/` and offers a choice if more than the neutral default
  was packaged in.
- Packaging: `npm run skill:zip` builds an upload-ready, self-contained
  skill ZIP (folder at the ZIP root) under `deploy/` for claude.ai web
  upload. `skill:zip:internal` / `skill:zip -- --from <dir>` overlay your
  own company templates from a git-ignored source -- never committed,
  never part of a public release.
- Style gate hardened: image/info-graphic styles are NEVER inferred from
  topic or tone; the skill presents the catalog menu and blocks slide
  creation until both are chosen (a gate separate from the outline gate).
- Style catalog expanded (pencil sketch, comic, watercolor, cinematic,
  duotone, blueprint, low-poly, ... with "best for" guidance), and step
  banners are phase-colored (analyze vs. write).
- Skill safety rules: prompt boxes are additive and a generated image is
  never deleted or overwritten as a side effect; a removed prompt box is
  the normal post-generation state and is re-created on request.
- README restructured: a dedicated "Packaging & Distribution" chapter and
  a "claude.ai (web)" install path for non-developers.

## 0.2.12 (plugin 0.2.17)

- Documentation sync: README and `--help` now cover the 0.2.9-0.2.11
  features (`W_ELEMENT_OVERLAP`, picture-placeholder `overlays`); the
  README gained a "PowerPoint Skill" section with install instructions
  and the plugin ships its own README. A new doc-sync contract test
  fails the build when any command, op, lint code or the skill install
  goes undocumented.

## plugin 0.2.16 (CLI unchanged at 0.2.11)

- Image-prompt phrasing rules (research-backed, Google docs): overlay
  regions are expressed as negative space ("vast empty canvas creating
  significant negative space"), never explained with typographic
  vocabulary (title/footer/caption trigger pseudo-text rendering);
  prompts end with "No text. No letters. No symbols." unless text is
  deliberately embedded as a short quoted string.

## 0.2.11 (plugin 0.2.15)

- Picture-placeholder overlay report: `tpl inspect` exposes per picture
  placeholder which text shapes sit on top and where (`overlays`, region
  relative to the image: "bottom area, left part"); `tpl describe`
  narrates it ("keep these regions calm in images"). The skill turns
  every overlay into an image-prompt clause so covered regions stay calm
  and free of relevant detail.

## plugin 0.2.14 (CLI unchanged at 0.2.10)

- Plugin-root `bin/pptc`: a thin PATH wrapper around the bundle in
  `skills/powerpoint/scripts/` -- `pptc <command>` works in the user's
  shell while the plugin is enabled. The skill keeps invoking the
  explicit scripts path (deterministic in every install variant).

## 0.2.10 (plugin 0.2.13)

- Engine fix (repair trigger): the post-pass relationship-id counter now
  starts above every `rIdPptc` left by earlier applies -- re-wiring
  hyperlinks (e.g. refilling a placeholder that contains links) used to
  collide with ids from the previous apply, producing duplicate
  relationship ids. Existing duplicates self-heal on the next apply;
  the integrity validator now checks rel-id uniqueness.
- Plugin: the bundled CLI moved from `skills/powerpoint/bin/` to
  `skills/powerpoint/scripts/` -- `scripts/` is the documented
  supporting-file convention, and `bin/` officially denotes the
  plugin-root PATH directory.

## plugin 0.2.12 (CLI unchanged at 0.2.9)

- "Placeholders first" rule: text content always goes into layout
  placeholders via slide.fill (rich text covers monospace, sizes,
  colors, hyperlinks); el.add is reserved for tables/charts/shapes/
  images/connectors and sanctioned overlays.

## 0.2.9 (plugin 0.2.11)

- New lint `W_ELEMENT_OVERLAP`: an `el.add` element that covers a
  text-bearing shape (placeholders incl. the footer/slide-number area,
  textboxes, tables, charts) is reported with the covered shape's name;
  `--strict` escalates to exit 7. Prompt boxes and picture areas are
  exempt; shapes removed earlier in the same ops document do not count.
- Reader: layout placeholders without their own geometry now inherit
  the master's matching placeholder frame -- capacity and overlap lint
  work on inheritance-based templates (e.g. Office default); layouts
  expose the reserved footer/slide-number/date frames.
- Plugin: element-placement rule (never cover text fields) and the
  concurrent-edit protocol (fresh `state` before every write, exit-6
  recovery: re-read, re-verify targets, rebuild, retry once).

## 0.2.8 (plugin 0.2.10)

- The image-prompt box label is English ("IMAGE PROMPT"); the box
  border, the schema error example and the default chart palette use
  neutral Office colors instead of leftover corporate values.

## plugin 0.2.9 (CLI unchanged at 0.2.7)

- Deck sidecar: the skill persists the deck setup (title, topic, deck
  language, image and info-graphic style, template notes) in a
  `<deck>.md` next to the deck and restores it when reopening -- the
  deck's memory across sessions.

## 0.2.7 (plugin 0.2.8)

- English throughout: the `tpl describe` narration (suitability hints,
  positions, capacities -- previously German), all help/README examples
  and the skill's output templates are now English. Deck content and
  user-facing skill responses still follow the deck/user language;
  German strings remain only as deliberate language examples and
  unicode-escaping test data.

## 0.2.6 (plugin 0.2.7)

- Detailed CLI help: every command documented with options and an
  example (`pptc help <command>`), plus the topics `help ops` (the
  write API: ops document anatomy, all ops with payload sketches) and
  `help selectors`. `--help`/`-h` anywhere prints plain text -- humans
  typing --help never see the JSON envelope.
- README: "The Ops Document" section explaining structure, fill
  payloads, expectRev and the recommended agent loop; --plain/--help
  console examples.

## 0.2.5 (plugin 0.2.6)

- Sections handling: `new` strips the template's PowerPoint sections
  (they only group the stripped example slides); `apply` prunes section
  slide references whose slide was removed, so user-created sections
  survive edits without stale refs. Integrity validator checks section
  references.
- `--plain` extended to all read commands: `help`, `state`, `tpl list`
  and `tpl validate` join `tpl describe` (readable text instead of the
  JSON envelope; write commands stay envelope-only).

## 0.2.4 (plugin 0.2.5)

- `tpl describe --plain` prints the raw Markdown without the JSON
  envelope -- a human console escape hatch (agents keep the envelope).

## plugin 0.2.4 (CLI unchanged at 0.2.3)

- The skill bundles a NEUTRAL fallback template (Microsoft's default
  Office design, same provenance as the test fixture) with a sidecar:
  used when the user names no template, always announced as the neutral
  default with a hint that a corporate template can be supplied instead.
  No corporate material involved.

## 0.2.3

- Engine: two more repair triggers eliminated -- (1) `CT_Presentation`
  element order is normalized in the seed builder (notesMaster/handout
  id lists must precede `sldIdLst`; sloppy templates violate this),
  (2) charts, embeddings and media are garbage-collected: automizer
  duplicates chart parts and accumulates stale slide relationships on
  every re-apply (1 chart became 64 across 6 applies); unused slide
  rels are pruned and unreachable assets removed. This also closes the
  documented "media of removed slides remains" limitation.
- Integrity validator: new checks for presentation element order and
  orphan chart/embedding/media parts; stress suite asserts exactly one
  chart part survives repeated applies.
- Test fixture: presentation element order fixed.

## 0.2.2

- Engine: sweep stale and duplicate content-type overrides (another
  PowerPoint "repair" trigger) -- in the post-pass after every apply and
  in the seed builder, so templates carrying override debris of their own
  produce clean decks; `removeParts` now strips ALL matching overrides.
- Tests: new file-integrity validator (`test/util/integrity.ts`) asserting
  every written deck against the known repair triggers -- well-formed XML,
  unique shape ids, no dead relationship targets, no dead/duplicate
  content-type overrides, full content-type coverage, resolvable sldIds;
  wired into the integration, elements and contract suites.
- Test fixture: removed phantom slideMaster2/3 overrides from
  `neutral-template.pptx`.
- Plugin: element slides (tables/charts/shapes via `el.add`) always use
  the blank-role layout (title + empty surface), never layouts with text
  placeholders in the content area.

## 0.2.1

- Plugin: the skill no longer bundles a template -- it works exclusively
  with external `.potx`/`.pptx` templates (path, directory scan with
  selection menu, or explicit ask). Template-specific knowledge lives in
  a sidecar Markdown next to the template file.

## 0.2.0

- Engine fixes (PowerPoint "repair" triggers): drop stale presentation
  relationships after re-apply slide renames, re-point notesSlide
  back-references at the renamed parent slide, and make `cNvPr` shape
  ids unique per slide.
- Built slides now keep the layout's footer, slide-number and date
  placeholders with their content (footer text, `slidenum` field), so
  decks show footer and page number like PowerPoint-inserted slides.
- Decks built from templates saved in master/layout view now open in
  normal view (`lastView` stripped); seed format bumped to 4.
- Claude Code plugin: new `plugin/skills/powerpoint` skill (template-aware
  deck building plus color-faithful image prompts) with bundled `pptc.mjs`,
  marketplace/plugin manifests and the `plugin:sync` npm script.
- Docs: full per-command CLI reference in the README plus an
  "Architecture in a Nutshell" section (automizer / PptxGenJS / zip
  post-pass); ARCHITECTURE.md covers the repair-cleanup duties.

## 0.1.0

Initial release.

- Kernel CLI: `tpl list/describe/inspect/validate`, `state`, `new`, `apply`,
  `schema`, `update`, `help` plus the micro-edit sugar commands
  `text`, `note`, `footer`, `rm`, `move`.
- Ops vocabulary: `slide.add/fill/rm/move/copy`, `el.add/set/rm`,
  `img.prompts`, `meta.props` -- validated by Zod, planned as a whole,
  applied atomically.
- Template intelligence: `tpl describe` derives layouts, ASCII minimaps,
  semantic positions, text capacities and image aspect ratios generically
  from OOXML geometry.
- Determinism: stable slide ids and title selectors, `$ref` for new slides,
  `rev` tokens with optimistic locking, exit codes per error class,
  capacity lint with `--strict`.
- Engine: pptx-automizer + PptxGenJS interop with on-demand seed decks per
  template (content-hash cached) and a zip-level post-pass for notes,
  footers, backgrounds, placeholder images, hyperlinks and properties.
