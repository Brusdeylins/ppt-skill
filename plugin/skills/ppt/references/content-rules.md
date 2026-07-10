Content Rules
=============

Guardrails for clean slides. These are prompt rules enforced while
writing slide content, plus the validation that backs them up.

Titles
------

-   **Action titles**: every slide title states the takeaway, not the
    topic -- `"Revenue beats forecast by 12%"`, not
    `"Q3 results"`. Reading ONLY the titles, in order, must tell
    the deck's complete story.
-   **Unique titles**: no two slides share a title (titles are also
    the stable `title:` address for later edits).
-   **Title capacity**: the title placeholder has a budget too --
    `tpl describe` reports `~N lines of ~M chars` for it (a safety buffer
    is already subtracted). Keep the action title within it; a title that
    would wrap past its box triggers `W_TEXT_OVERFLOW`, so shorten it (or,
    if the message truly needs the length, pick a layout with a roomier
    title). Title layouts are often single-line -- watch the char budget.

Body
----

-   **One message per slide**: the slide's message is fixed as ONE
    sentence in the outline before any content is written. If it does
    not fit one sentence, split the slide.
-   **6x6 guideline**: at most ~6 bullets per slide, ~6-8 words per
    bullet, indent `level` at most 2. No prose paragraphs on slides --
    details go into the speaker notes.
-   **Capacity over instinct**: `pptc tpl describe` states per
    placeholder `~N lines of ~M chars`. Write text AGAINST that
    budget. Validation: `pptc apply ... --dry-run --strict` fails with
    `W_TEXT_OVERFLOW` (exit 7) on overflow -- then SHORTEN or SPLIT,
    never rely on PowerPoint auto-shrink. That applies to SELF-AUTHORED
    text only: approved plan/user wording is never shortened without
    explicit confirmation (SKILL.md STEP 6).
-   **Consistent font sizes -- take them from the template, never invent.**
    `tpl inspect` reports each placeholder's `fontSizePt`. Use the BODY
    size for body/explanatory text and the FOOTER size (the small one) for
    source/link footnotes. Code keeps its one sidecar size; all tables
    share ONE size deckwide. Scattered ad-hoc sizes are visual noise --
    avoid size variation that has no reason.

Language and footer
-----------------

-   **Deck language**: fixed once in deck setup (ask when not obvious
    from the request), independent of the conversation language. ALL
    deck-facing text -- titles, bullets, notes, footer, AI note -- is
    written/translated into the deck language. Image prompts stay
    English.
-   **Footer on every slide**: the template's footer pattern (see
    template sidecar) with the actual deck title and the CURRENT year.
    Never ship the template's placeholder title (e.g. "Insert title") or
    a stale year; ask for the title when it cannot be derived.
-   **AI-image note**: every slide whose layout has picture
    placeholders carries a note that the images are AI-generated, in
    the deck language (de: "Bilder mit KI generiert", en: "Images
    AI-generated"), appended to the footer. Layouts without a footer
    placeholder (typically title/closing) get the note as a small
    discreet textbox near the bottom edge instead (`el.add`).

Elements (tables, charts, free shapes)
--------------------------------------

-   **Placeholders first.** Anything that is TEXT goes into a layout
    placeholder via `slide.fill` -- rich text covers fonts, sizes,
    colors, bullets and hyperlinks. Pick the layout whose placeholder
    fits the content (`tpl describe` capacities) instead of placing a
    textbox on a blank layout. `el.add` is ONLY for what placeholders
    cannot hold: tables, charts, shapes, images, connectors -- and the
    sanctioned overlays (prompt boxes, AI note where no footer
    placeholder exists). A code block is TEXT, NOT an exhibit.
-   **Code blocks / file trees are placeholder text, not `el.add`
    textboxes.** Put them in a content body placeholder (a
    title-plus-content layout) via `slide.fill`, with every run in the
    template's **code font** (from the template sidecar, e.g. a
    monospace family member) and **bullets turned off** (`bullet: false`),
    preserving leading spaces for indentation. Do NOT build a free
    textbox "code card" -- that bypasses the template. (A `ppt-prepare`
    plan may label such a slide layout type "code block"; that maps to a
    BODY PLACEHOLDER filled with the code font, not to `el.add`.)
-   Slides that receive free elements via `el.add` ALWAYS use a layout
    WITHOUT text placeholders in the content area (role "blank" --
    title plus empty surface). Elements must never sit on top of body
    placeholders; that layout role comes from the template sidecar or,
    failing that, the layout whose only placeholder is the title.
-   Position elements on the template's own grid: `tpl inspect`/`describe`
    report a **`contentArea`** (the body region snapped to the master's
    guides) and the raw `guides`. Place a full-width table at the
    `contentArea` frame; centre a narrower textbox/diagram within it. This
    aligns every element to the template instead of guessing coordinates.
-   **Stay inside the contentArea -- never cross the guides.** EVERY
    `el.add` element (table, SVG/native diagram, textbox) must lie FULLY
    within the `contentArea`: its top at or below `contentArea.y` (do NOT
    start in the title band above it), its bottom at or above
    `contentArea.y + h`, and its left/right within the side guides. When a
    slide stacks a table or diagram AND an explanation textbox, the two
    TOGETHER must fit inside the contentArea height -- shrink rows or text,
    or step font sizes down within the template's scale (never below the
    lint floor), to fit; never let an element overflow the guides or the
    slide.
-   **Uniform tables.** Across the deck all tables share ONE look: header
    fill = the template's primary accent (default `accent1`; honor a
    sidecar primary-slot override, see `color-roles.md`) with white text,
    single-line headers, a light `altRowBg`, a subtle `border`, and a
    FIXED `headerHeight`/`rowHeight` (table `style`) so header bands and
    row heights match on every slide.
-   **Never cover existing text fields** (placeholders incl. the
    footer/slide-number area, textboxes, tables, charts) with a new
    element -- prompt boxes and the AI-note textbox (where no footer
    placeholder exists) are the only sanctioned overlays, and even the
    AI note must not sit on text. pptc enforces this: `W_ELEMENT_OVERLAP`
    names the covered shape; under `--strict` the apply fails (exit 7).
    On a finding, REPOSITION the element, never ignore it.
-   **Element font sizes -- from the template, readable, few.** Free
    `el.add` elements (chips, cards, captions, diagram labels,
    annotations) carry no template size of their own -- NEVER invent one.
    Derive every element size from the template's own scale: `tpl
    inspect` reports each placeholder's `fontSizePt`. Use the BODY size
    for primary element text (card titles, central diagram labels), one
    step below it for secondary text (chips, card subtexts, annotations),
    and the template's SMALLEST placeholder size for captions, helper
    labels and hints; sizes the sidecar fixes itself (e.g. the code font)
    always win. Keep to **3-4 sizes per slide**; scattered ad-hoc sizes
    are visual noise. As a pure legibility backstop pptc enforces a
    floor: `W_FONT_TOO_SMALL` names any run/element below the minimum
    (default 8pt, tune with `--min-font-pt`, `0` disables); under
    `--strict` the apply fails (exit 7). Footer/slide-number/date
    placeholders and prompt boxes are exempt -- footnote-scale text
    belongs in the footer placeholder, not a tiny free textbox.

Self-explanatory slides (decks READ without a presenter)
--------------------------------------------------------

When the deck is self-study material -- read with no speaker to fill the
gaps (the briefing says so) -- every slide must stand on its own. A bare
exhibit (code, table, diagram, image) under a title is NOT enough.
(Assertion-Evidence, Alley; "slidedoc", Duarte; multimedia principles,
Mayer.)

-   **Assertion title + takeaway -- land a point.** The title is a
    full-sentence claim; the slide also carries a one-line plain-language
    **takeaway** -- the "so what / what this means for you" -- besides the
    exhibit. Every slide closes on a POINT: a conclusion (the
    implication) or a pointed question ("wouldn't it be better if ...?").
    Showing the status quo alone is not a finished slide.
-   **Explain every exhibit, for a beginner:**
    -   *Code:* a lead-in sentence (what it shows) + highlight the key
        line + a one-line plain-language takeaway; keep <= ~15 lines.
    -   *Table:* a lead-in/takeaway sentence; gloss what the columns and
        entries MEAN for a beginner (units, plain words); highlight the
        key row/cell. Explanatory prose that does not fit a placeholder
        goes in a **textbox beside the table, in the template's body font
        and size** (sidecar) -- a sanctioned exception to "placeholders
        first" for explaining an `el.add` exhibit.
    -   *Diagram:* label parts NEXT TO them (spatial contiguity), never in
        a separate legend.
-   **Signal the point.** Bold/colour/highlight the ONE element that
    carries the message (key code line, key cell). One key point per slide.
-   **Sparse slide -> side-image layout.** If a slide carries little text
    (a few short bullets or one short statement), do NOT ship a near-empty
    full-width text slide: choose a TEMPLATE LAYOUT WITH A SIDE PICTURE
    placeholder (bullets + image) and a MEANINGFUL image -- or an
    infographic that proves the message -- never decoration. Accept a
    SHORTER assertion title when the side-image layout's title area is
    narrow; the image earning its half of the slide beats empty space.
    Reserve full-width text layouts for genuinely text-dense slides.
-   **Cite on the slide.** Put the source (a short `Source: <url>`) on the
    slide where the claim/data/quote appears; a references-only-at-the-end
    slide is not enough for a read-alone deck. One representative source
    per slide; the full list still closes the deck.
-   **Word budget.** Read-alone slides may carry more text than a live
    slide (~100-250 words, the slidedoc range) -- enough to be
    self-explanatory, but skimmable: visual hierarchy, bold keywords, one
    idea per slide.

Speaker notes
-------------

-   **Presented decks:** every content slide gets notes via the `notes`
    field: 40-70 words covering the core message, key numbers, and the
    transition to the next slide.
-   **Self-study / teaching decks SKIP notes** -- the slide is read without
    a presenter, so it must be self-contained; the explaining text belongs
    ON the slide, not in a notes pane no one opens.
-   Slides with charts/images additionally get 1-2 sentences
    describing the visual (doubles as accessibility text) -- on a self-study
    deck this description is on the slide, not in notes.

Structural slides
-----------------

-   **Agenda** derives from the chapter titles, one bullet each. When
    chapters are added/renamed/removed, update the agenda slide in the
    same ops document (`slide.fill` on the agenda slide).
-   **Chapter divider** carries the chapter title AND a one-line chapter
    **conclusion** as a subtitle (an advance organizer: what the chapter
    concludes), synthesized from the chapter's slides and their takeaways.
    Its image synthesizes the WHOLE chapter (see `prompt-formula.md`), not
    just the title.
-   **Chapter summary.** Each chapter ENDS with a "Key takeaways" slide:
    2-4 short bullets distilled from the chapter's slides (close on the
    chapter conclusion). This is segmenting + retrieval for self-study decks.
-   **Closing slide** carries a call-to-action, not a summary.
-   Business decks follow SCQA where it fits: situation, complication,
    question, answer -- recommendation early, not last.

Final checklist (report unanswered items to the user)
------------------------------------------------------

- [ ] Titles alone tell a coherent story (read in order)
- [ ] Every slide has exactly one message
- [ ] No placeholder over capacity (`--dry-run --strict` passes)
- [ ] Speaker notes on every content slide (presented deck) / no notes
      (self-study deck)
- [ ] Agenda matches the actual chapters
- [ ] Image prompts use template hex codes and the deck-wide styles
- [ ] All slide titles unique
- [ ] All deck text in the deck language (no conversation-language leaks)
- [ ] Footer carries the deck title and the current year on every slide
- [ ] AI-image note present on every slide with picture placeholders
- [ ] Self-study deck: each slide stands alone -- assertion title + a
      plain-language takeaway; no bare code/table/diagram exhibit
- [ ] Code/table/diagram explained for a beginner; key line/cell signalled
- [ ] A representative source/URL is cited on each slide that makes a
      claim or shows data (not only on the closing references slide)
- [ ] Sparse slides use a meaningful image/infographic, not empty space
