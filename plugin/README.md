# ppt -- Claude Code plugin

Template-aware PowerPoint engineering for Claude Code, built on the
bundled [pptc](../README.md) CLI. The plugin ships **two skills**:

- **`ppt`** -- builds and edits PPTX decks (the main skill).
- **`ppt-prepare`** -- plans the *content* first (story, MECE arguments,
  headline titles, a per-slide plan) and hands the approved plan to `ppt`.

## What the skills do

- **Deck building, outline-first**: storyline and outline are gated for
  your approval before any slide is created.
- **Content guardrails**: action titles, one message per slide,
  capacity-checked text (pptc lint, exit 7 on overflow), speaker notes,
  footer with title/year/AI-image note, agenda kept in sync with
  chapters.
- **Color-faithful image prompts**: one Nano-Banana-Pro prompt per
  picture placeholder, using the template's real theme colors, the
  placeholder's aspect ratio and its overlay regions (covered areas
  become negative space). Prompts are written onto the slides as
  removable boxes -- no image API is called.
- **Edit-safe**: optimistic locking against concurrent PowerPoint
  edits; slides addressed by stable ids and titles, never indices.
- **Per-deck memory**: language, styles and topic persist as custom document
  properties INSIDE the `.pptx` (so a deck is self-describing and can be handed
  to another person who continues via the skill).
- **Version-aware**: on first activation each skill prints a one-line version
  banner and does a best-effort check against the GitHub releases, flagging
  when a newer version is available (silent offline, never blocks the work).

## Install

```
/plugin marketplace add Brusdeylins/ppt-skill
/plugin install ppt@ppt-skill
```

Requires Node.js >= 20 (the bundled pptc is a Node program). While the
plugin is enabled, `bin/` puts a `pptc` command on the Bash tool's PATH.

## Templates

The skill prefers your external `.potx`/`.pptx` (path or directory scan).
The public build bundles a neutral Office-default template as the fallback;
when no template is named, the skill uses the bundled template (or, if
several are bundled, offers a choice). Template-specific knowledge (layout
roles, footer pattern, design constraints) lives in a Markdown sidecar next
to the template file -- see `skills/ppt/assets/neutral-template.md` for the
pattern.

Company templates are never part of the public release: an in-house
distribution lives in a downstream company repo that vendors this plugin
and **replaces** the neutral default with the corporate templates in the
skill's `assets/`.

## Layout

```
plugin/
  .claude-plugin/plugin.json   manifest
  bin/pptc                     PATH wrapper
  meta/control.md              shared control-tag definitions (all skills)
  skills/ppt/
    SKILL.md                   the build skill (imports ../../meta/control.md)
    VERSION                    bundled version (checked against GitHub releases)
    references/                content rules, prompt formula, styles
    scripts/pptc.mjs           bundled pptc build
    assets/                    neutral fallback template + sidecar
  skills/ppt-prepare/
    SKILL.md                   the content-planning skill (imports ../../meta/control.md)
    VERSION                    bundled version (checked against GitHub releases)
    references/                methodology + storyline patterns
```
