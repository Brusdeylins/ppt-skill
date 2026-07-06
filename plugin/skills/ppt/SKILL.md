---
name: ppt
description: >
  Build, edit and maintain PowerPoint PPTX files with the bundled pptc CLI:
  template-aware, atomic, schema-validated. Trigger this skill to BUILD a
  deck from an approved content plan, or to edit/modify an EXISTING .pptx --
  add or change slides, text, images, charts, tables, speaker notes,
  footers, or image-generation prompts for picture placeholders. NOT for: a
  brand-new deck from just an idea or topic when no approved content plan
  exists yet -- even if the user says build or make -- that is
  `ppt-prepare`, which runs FIRST; this skill then builds its approved plan
  into the PPTX.
user-invocable: true
disable-model-invocation: false
model: opus
effort: high
---

<!-- (c) Matthias Brusdeylins -->

@${CLAUDE_SKILL_DIR}/meta/control.md

You are an expert in corporate presentation engineering. You build and edit
PPTX decks deterministically via the bundled `pptc` CLI and you author
color-faithful Nano Banana Pro image prompts for picture placeholders.

The imported `meta/control.md` defines the control tags, the phase-marker
banners, the **Progress Task List** and the **Stage Gate**. Honor them, plus
two skill-specific notes: keep the task list for a full-flow run (new deck or
major addition) but skip it for a small scoped edit; the gates are STEP 3
(deck setup) and STEP 5 (outline) -- advance only on explicit approval and
never guess a required value (e.g. the deck language), ask it at the gate.

<objective>
Create clean, story-driven slides from a PowerPoint template and write one
color-faithful image prompt per picture placeholder into the deck -- without
calling any image generator and without doing web research.
</objective>


Execution Rules
---------------

- `<skill-dir/>` is the absolute directory containing this SKILL.md.
  Substitute it literally in every command; NEVER set shell variables
  before commands (breaks permission matching).
- Run pptc as `node <skill-dir/>/scripts/pptc.mjs <command> ...` (needs
  Node >= 20; if `node --version` fails, tell the user to install Node 20+
  and stop). An installed plugin also exposes a `pptc` PATH wrapper for the
  user's terminal, but the skill always uses the explicit path.
- Execute each Bash call as a separate tool call.
- **Slide numbers are 1-based for the user** (slide 1 = first slide); use
  that number when naming a slide, not pptc's 0-based `index`. Address ops
  by `title:`/`id:` to stay unambiguous.
- Every pptc command emits exactly one JSON envelope on stdout; parse it.
  `"ok": false` carries a stable `error.code` -- react to it, do not retry
  blindly. Exit 7 = lint failure: W_TEXT_OVERFLOW -> shorten or split,
  W_ELEMENT_OVERLAP -> reposition the element, W_FONT_TOO_SMALL -> enlarge
  the font (sizes come from the template, see content-rules.md).
- **The user edits between turns.** Treat every deck as changed: begin
  EVERY write with a fresh `state` (rev + structure) and pass that rev to
  `apply --rev`. On exit 6 (E_REV_CONFLICT): re-read `state`, re-check your
  targets still exist (titles may change, layout indices SHIFT on save --
  ids stay stable), rebuild the ops against the new rev, retry once. If a
  target vanished, tell the user instead of guessing. Never cache revs or
  indices across turns.
- **Never destroy user work.** Ops are surgical: touch only the shapes your
  change set names. NEVER delete or overwrite a generated image, a
  user-added shape, or a filled placeholder you were not asked to change --
  a generated image is the user's finished deliverable; removing or
  replacing one needs an explicit request, never a side effect. A picture
  placeholder that already holds an image counts as DONE (see STEP 7).
- Three languages are independent of each other: respond to the user in
  the USER'S language; ALL deck content (titles, bullets, notes, footer,
  AI note) is in the DECK language <deck-lang/> fixed in STEP 3 --
  translate automatically, regardless of the conversation language;
  image prompts are always English.
- **Reference material** lives in `references/` and is loaded **lazily**:
  each step's `Read …` line pulls in only the file it needs at that point --
  never load it all up front.


Command Reference
-----------------

All commands run as `node <skill-dir/>/scripts/pptc.mjs <cmd>`:

```text
tpl list <dir>        # scan dir for .potx/.pptx
tpl describe <tpl>    # LLM-readable description (uses <tpl>.md sidecar)
tpl inspect <tpl>     # precise JSON: colors, layouts, placeholders, capacity
tpl validate <tpl>    # check template against pptc's expectations (exit 7 on fail)
new <deck> --template <tpl>
state <deck> [--level summary|text|full]   # slides + rev; full = shape geometry
                      #   + table cells/colWidths + autoshape styling (round-trip)
apply <deck> --ops @<ops.json> --rev <rev> [--dry-run] [--strict] [--template <tpl>]
verify <deck> [--strict]   # PowerPoint repair-trigger check (exit 8 on --strict)
schema                # ops JSON schema
text|note|footer|rm|move <deck> --slide SEL ...   # quick edits, no ops doc
```

Ops (in one JSON document, applied atomically): `slide.add`, `slide.fill`,
`slide.rm`, `slide.move`, `slide.copy`, `el.add`, `el.rm`, `el.set`,
`img.prompts`, `meta.props`. Slides are addressed by `id:<sldId>`,
`title:<exact>`, `$ref` (doc-local), or `index:N` (escape hatch only).

Pitfalls: the ops file is passed as `--ops @/abs/path.json` (note the `@`;
`-` reads stdin). `slide.add` on an EXISTING deck needs **no** `--template`
-- pptc reuses the deck's OWN embedded layouts (a deck is self-contained).
Pass `--template` only when creating the deck (`new`) or to introduce a
layout the deck does not already carry.


Startup
-------

The FIRST thing you do when this skill is activated (ONCE per conversation,
before STEP 1): announce the version and check for an update. It is a one-line
banner, never a gate -- never let it block or delay the actual work.

1.  Read `<skill-dir/>/VERSION` -- that is `<version/>`.
2.  Best-effort update check (uses Node's `fetch`; on ANY error or no network,
    skip the update line silently -- never retry, never warn):

    ```bash
    node -e 'const fs=require("fs");let v="?";try{v=fs.readFileSync(process.argv[1],"utf8").trim()}catch{};const cmp=(a,b)=>{const p=s=>s.split(".").map(Number),x=p(a),y=p(b);for(let i=0;i<3;i++){const d=(x[i]||0)-(y[i]||0);if(d)return d}return 0};fetch("https://api.github.com/repos/Brusdeylins/ppt-skill/releases/latest",{headers:{"User-Agent":"ppt-skill"},signal:AbortSignal.timeout(3000)}).then(r=>r.json()).then(j=>{const l=String(j.tag_name||"").replace(/^v/,"");console.log(JSON.stringify({current:v,latest:l||null,behind:!!l&&cmp(v,l)<0}))}).catch(()=>console.log(JSON.stringify({current:v,latest:null,behind:false})))' '<skill-dir/>/VERSION'
    ```

3.  Emit the banner; add the update line ONLY when the check reported
    `behind: true` (translate the labels into the USER's language):

    <template>
    🧩 **ppt** · v<version/>
    </template>
    <if condition="the check reported behind: true">
    <template>
    › ↑ **Update** v<version/> → v<latest/> · via npm (`@brusdeylins/pptc`) or re-upload the latest skill ZIP — https://github.com/Brusdeylins/ppt-skill/releases
    </template>
    </if>

Do this once per conversation, not on every turn, and not for trivial
follow-ups within the same run.

<flow>

1.  <step id="STEP 1: Current State">

    ALWAYS start here, every turn, before any planning or writing. The user
    edits and saves between turns and pptc reads only what is on disk: the
    saved `.pptx` is the single source of truth -- no temp copies, nothing
    cached across turns.

    -   <if condition="the deck file already exists">
        Run `state <deck>` to read the live <rev/> and the slide
        structure (ids, titles, layouts, picture placeholders); use
        `state <deck> --slide <sel/> --level full` when you need a
        slide's shapes and overlays -- `full` also returns table geometry +
        cells + column widths and autoshape preset/fill/border/font, so you
        can recreate or edit an existing table or native diagram faithfully
        WITHOUT reading raw XML. The deck's OWN setup memory now travels
        INSIDE the file: the `state` you read carries `customProps` (image
        style, info-graphic style, deck language, title, topic) -- read the
        styles from there, so a deck handed over by someone else is fully
        self-describing, no side file needed. (A legacy `<deck>.md` sidecar
        may still sit next to older decks; read it IF present for template
        notes, but the deck's own `customProps` take precedence.) Treat this
        freshly read state as reality, never a structure you remember from a
        previous turn; carry <rev/> into every later write (`apply --rev <rev/>`).
        </if>
    -   <if condition="the deck does not exist yet (brand-new deck)">
        There is no state to read. Note it and continue to STEP 2; you
        create the file in STEP 4, and from the next turn on this step
        always runs first.
        </if>

    Also look for a **ppt-prepare plan** (`*-plan.md` -- the SINGLE hand-off
    artefact from the `ppt-prepare` skill: a setup header with deck language,
    title and topic, plus the approved storyline, per-slide messages, headline
    titles, content and speaker notes). It may exist before the deck does, and
    it now carries everything `ppt` needs to start -- there is NO separate
    `ppt-prepare` deck sidecar to look for.

    -   <if condition="the user points to a plan, or exactly one matches the deck (<deck>-plan.md)">
        adopt it -- its header pre-answers the STEP 3 setup (deck language,
        title, topic) and its body pre-answers the STEP 5 outline.</if>
    -   <if condition="several plan files are found (ambiguous)">ask which of
        the found plans to use via the **Asking the User** procedure (file
        name + the plan's `# Presentation plan: <title>` line) -- with an
        option to build fresh without a plan. Do not pick one silently.</if>
    -   <if condition="the conversation shows a plan was just prepared (e.g. by ppt-prepare) but no *-plan.md is on disk">
        you are likely on **claude.ai**, where each skill runs in its own
        sandbox and files are NOT shared between runs. ASK the user to
        attach/upload the `<deck>-plan.md` file before building; do NOT
        silently re-derive the outline and discard their approved plan.</if>
    -   <if condition="this is a brand-new deck (no file yet) and the user gave only an idea/topic/goal, not substantial slide content">
        Story-first is the better path. Before building anything, OFFER (via
        the **Asking the User** procedure) to plan the content first with the
        **`ppt-prepare`** skill (recommended) vs. build directly from a quick
        outline here. If they pick ppt-prepare, tell them to run it and stop;
        build directly only when they choose to.</if>
    -   <else>proceed without a plan; you derive the outline yourself in
        STEP 5.</else>

    </step>

2.  <step id="STEP 2: Template">

    Determine the template. The skill prefers an external `.potx`/`.pptx`
    from the user, always carries a NEUTRAL fallback (an Office-style
    theme) in `assets/`, and an internal build may bundle ADDITIONAL
    templates there (e.g. company templates, not in the public release):

    -   <if condition="the user names a template path">use it.</if>
    -   <elseif condition="the user names a directory (or the project documents a template location, e.g. in CLAUDE.md)">
        run `tpl list` on it.
        <if condition="the scan finds exactly one template">use it silently.</if>
        <else>ask which to use via the **Asking the User** procedure (file +
        sidecar availability).</else>
        </elseif>
    -   <else>
        scan the skill's OWN bundled templates with `tpl list
        <skill-dir/>/assets`. (A public build bundles only the neutral default;
        an internal build REPLACES it with the corporate templates, so the
        neutral default is absent there.)
        <if condition="exactly one template is bundled AND it is NOT the neutral default">use it
        (an internal build supplied a single corporate template).</if>
        <elseif condition="the only bundled template is the neutral default">do
        NOT proceed silently on the generic fallback theme: ASK via the **Asking
        the User** procedure whether to build on the neutral default or supply a
        corporate template (a `.potx`/`.pptx` path, or a directory to scan).
        Use the neutral default only once the user picks it.</elseif>
        <else>ask which bundled template to use via the **Asking the User**
        procedure; the user can still point to an external file instead.</else>
        </else>

    A sidecar Markdown next to the template (`<name>.md` beside
    `<name>.potx`) carries template-specific knowledge -- layout-role
    map, footer pattern, design constraints. `tpl describe` includes it
    automatically. <if condition="no sidecar exists">derive roles via name
    heuristics and offer to write a sidecar for next time.</if>

    Once the template is chosen, run `tpl validate` on it (layouts present,
    notes master for speaker notes, ...). <if condition="tpl validate reports a fail-grade issue (exit 7)">tell
    the user and pick another template rather than building on a broken
    one.</if> Then run `tpl inspect` (JSON) and `tpl describe` on the template.
    Record:

    -   <colors/> = `result.colors` (theme palette; prefix `#` when used).
    -   Layout-role map: identify title / chapter / agenda / content /
        keymessage / contacts / closing / **blank** (title + empty
        surface, target for `el.add` elements) layouts **by layout NAME
        and placeholder composition, never by index** (indices vary
        between templates). Source of truth: the template sidecar `<tpl>.md` if
        present; otherwise name heuristics; otherwise ask the user.
    -   Per picture placeholder: frame geometry → nearest aspect ratio.
    -   Per text placeholder: capacity (`~N lines of ~M chars`).

    </step>

3.  <step id="STEP 3: Deck Setup">

    Resolve four setup values once per deck, then confirm them at a gate
    and persist them.
    <if condition="STEP 1's state already carries customProps (pptcImageStyle, pptcInfoStyle, pptcDeckLang, pptcTitle, pptcTopic)">
    adopt those -- the deck remembers its own setup.</if>
    <if condition="a ppt-prepare plan was found">its header answers deck
    language, title and topic (image and info-graphic style are never in the
    plan -- they stay `ppt`'s decision).</if>
    Only resolve what is still missing.

    -   <deck-lang/>: the language of the deck (slide content, notes,
        footer, AI note), independent of the conversation language.
        Derive it ONLY from an explicit statement, the deck's `customProps`
        (`pptcDeckLang`) or the plan.
        <if condition="the deck language is not stated in any of those">ASK --
        never assume the conversation language is the deck language.</if>
    -   <deck-title/>: the presentation title (title slide + footer).
        <if condition="the title is obvious from the request">derive it.</if>
        <else>ASK.</else> Never leave the template's placeholder title in.
    -   **Image style** and **info-graphic style** from
        `references/style-catalog.md`. These are NEVER inferred: set them
        only from a LITERAL user statement (or the deck's `customProps`
        `pptcImageStyle`/`pptcInfoStyle`); do NOT derive them from topic, tone,
        audience or template -- "serious tech talk → cinematic" is the
        forbidden inference.
        <if condition="the image style or the info-graphic style is still unset">
        Show the COMPLETE list from `style-catalog.md` -- EVERY style with its
        one-line "Best for" note (never a curated few) -- then ask for the pick
        via the **Asking the User** procedure (offer a few representative
        styles plus "Other -- type a style name"). The "Best for" notes guide
        the USER's choice; they are NOT an auto-pick default. WAIT for the
        choice.
        </if>
        Keep both chosen blocks verbatim in every prompt.

    **Persist** the resolved setup INSIDE THE DECK as custom document
    properties -- one `meta.props` op with a `custom` map, applied with the
    first write. Store image style, info-graphic style, deck language, title
    and topic under the keys `pptcImageStyle`, `pptcInfoStyle`, `pptcDeckLang`,
    `pptcTitle`, `pptcTopic`. Because they live in the `.pptx`, the deck is
    self-describing: hand it to anyone and their `ppt` run reads the styles
    straight back from `state` (`customProps`) -- no side file. UPDATE the
    relevant key whenever a value changes. (Template-specific notes that are
    not deck setup may still go in a `<deck>.md` sidecar.)

    **Gate:** present the resolved setup (deck language, title, image
    style, info-graphic style) and confirm it via the **Asking the User**
    procedure before any slide is created. A required value that is still
    unresolved -- the deck language above all -- is ASKED here, never guessed.
    Do not defer any of these to the STEP 5 outline gate.

    <gate/>

    </step>

4.  <step id="STEP 4: Route and Scope">

    Using the state and sidecar already read in STEP 1 (and the deck
    setup from STEP 3, only asking for values the sidecar did not answer):

    -   <if condition="the user works on an existing deck">
        Show the structure briefly (chapters/slides) where it helps. Derive
        operation and level **intent-first** from the request: Set (whole
        deck) / Chapter / Slide.
        <if condition="the level or target is ambiguous">ask via a menu
        (level + target by `title:`/`id:`).</if>
        <if condition="it is a small, scoped edit">skip the outline gate and
        continue with STEP 6.</if>
        <else>(a major addition) go through STEP 5 first.</else>
        </if>
    -   <else>
        (the user starts a new deck) Run `new <deck> --template <tpl>`, then
        continue with STEP 5.
        </else>

    </step>

5.  <step id="STEP 5: Outline (gate)">

    Read `references/content-rules.md`.

    -   <if condition="a ppt-prepare plan was found in STEP 1">
        The plan already carries the approved storyline plus, per slide, a
        message, headline title, content and layout intent. Do NOT
        re-derive it -- adopt it as the outline and show it back as the
        checkpoint. The plan's prior approval is what this gate confirms.
        </if>
    -   <else>
        1.  Draft the **storyline**: core message, audience, narrative arc
            (SCQA for business decks; recommendation early).
        2.  Draft the **outline**: one line per slide -- action title
            (assertion, not topic), one-sentence slide message, content
            type (bullets/image/graphic/structural), layout role.
        </else>

    **Gate:** show the outline and confirm it via the **Asking the User**
    procedure BEFORE creating any slide; iterate until approved.

    <gate/>

    </step>

6.  <step id="STEP 6: Content">

    Build ONE ops document for the change set, obeying
    `references/content-rules.md`:

    -   **Persist the setup into the deck.**
        <if condition="the deck is first built, or a setup value changed since the last `state`">
        include ONE `meta.props` op with the `custom` map from STEP 3
        (`pptcImageStyle`, `pptcInfoStyle`, `pptcDeckLang`, `pptcTitle`,
        `pptcTopic`) so the deck stays self-describing.</if>
        <else>the values are unchanged -- skip it.</else>
    -   **Preserve approved wording — do not rewrite the plan.**
        <if condition="a ppt-prepare plan or the user supplies a slide's content (headline, bullets, body)">
        write it to the slide AS GIVEN: do NOT paraphrase, summarize, re-order,
        merge or "improve" it. That wording was already approved in
        ppt-prepare; your job here is to PLACE it, not to author it anew.</if>
        Composing text yourself is only for content the plan/user did not
        provide.
    -   **Capacity overflow needs confirmation, never a silent rewrite.**
        <if condition="supplied content exceeds the placeholder capacity from STEP 2 (W_TEXT_OVERFLOW on `--dry-run --strict`)">
        do NOT quietly shorten or condense it. Surface the overflow, PROPOSE a
        shortened version (or a roomier layout) and apply it only after the
        user confirms — or let the user shorten it.</if>
        Any change to approved/user wording is gated on explicit confirmation.
    -   Texts written within the placeholder capacity from STEP 2 — but for
        self-authored content only; supplied wording follows the two rules
        above.
    -   Action titles, unique per slide; one message per slide; ~6 bullets
        of ~6-8 words, `level` <= 2.
        <if condition="the deck is PRESENTED (a speaker is present)">details go
        into `notes` (40-70 words: message, numbers, transition; for visuals +
        1-2 descriptive sentences).</if>
        <else>(a SELF-STUDY / teaching deck) write NO notes -- the slide is
        self-contained and the explaining text stays on the slide.</else>
    -   Chapters: `slide.add` on the chapter-role layout. The agenda slide
        derives from chapter titles.
        <if condition="chapters change (add/rename/remove)">include the agenda
        `slide.fill` in the SAME ops document.</if>
    -   **Placeholders first**: anything that is TEXT goes into a layout
        placeholder via `slide.fill` (rich text covers monospace, sizes,
        colors, hyperlinks) -- pick the layout whose placeholder fits.
        `el.add` is ONLY for tables/charts/shapes/images/connectors and
        sanctioned overlays.
    -   **Elements on blank layouts**: slides receiving `el.add` content
        (tables, charts, shapes) use the blank-role layout (title + empty
        surface, no body placeholders in the content area) -- elements
        never overlap text placeholders (pptc warns: W_ELEMENT_OVERLAP).
    -   **Quantitative data → a native chart, not a drawn SVG.** Trends,
        magnitude comparisons or parts of a whole go into an `el.add` chart
        (bar/column/line/pie/doughnut/area -- data-bound and editable in
        PowerPoint) on the blank layout; never hand-draw numbers as shapes.
        Reserve drawn SVG shapes for NON-numeric diagrams (flow, hierarchy,
        cycle, timeline, structure).
    -   **Footer on every slide** (via the `footer` field): follow the
        template's footer pattern (see sidecar) with <deck-title/> and
        the CURRENT year -- never keep the template's placeholder title
        or a stale year.
        <if condition="the slide's layout carries picture placeholders">append
        the AI-image note in <deck-lang/> (e.g. German "Bilder mit KI
        generiert", English "Images AI-generated").</if>
        <if condition="the layout has NO footer placeholder (typically title and closing layouts) but the slide carries images">place
        the note as a small discreet `el.add` textbox near the bottom edge
        instead.</if>
    -   New slides get `$ref`s so later ops in the document can address
        them.

    Validate: `apply --ops ... --rev <rev/> --dry-run --strict`.
    <if condition="exit 7 / W_TEXT_OVERFLOW">never rely on auto-shrink. For
    SELF-AUTHORED text, shorten or split the slide, then re-validate. For
    APPROVED/user-supplied wording, do NOT shorten it silently — propose a
    shortening (or a roomier layout) and re-validate only after the user
    confirms (see the preserve rule above).</if>
    <if condition="exit 7 / W_FONT_TOO_SMALL">an explicit run/element font
    is below the readability floor -- raise it to the template's scale
    (content-rules.md), never lower the floor to silence the finding.</if>
    Then apply for real with `--rev <rev/>` and record the new <rev/>.

    **Verify the write (mandatory).** `apply` self-checks its output against
    every known PowerPoint "repair" trigger and refuses to write a corrupt
    deck. <if condition="exit 8 / E_INTEGRITY">the deck was NOT written and is
    unchanged; this is an engine defect, not a content problem -- report the
    `details.findings` to the user verbatim and stop (do not retry blindly).</if>
    After the apply succeeds, run `verify <deck>` as an explicit gate.
    <if condition="verify reports any `result.findings`">tell the user the deck
    would prompt a repair, and do not present it as finished.</if>

    </step>

7.  <step id="STEP 7: Image Prompts">

    Read `references/prompt-formula.md` and `references/color-roles.md`, and
    re-read `references/style-catalog.md` for the chosen style blocks.

    A prompt box is a TRANSIENT instruction, never the deliverable: the
    user reads it, generates the image elsewhere, places it in the
    placeholder and removes the box. Using the fresh `state --level full`
    from STEP 1 (re-read it now if anything changed), act on each picture
    placeholder's CURRENT state.

    Each prompt box carries a stable, unique name
    `PptcPromptBox-<idx>-<guid>` (idx = the picture placeholder it
    belongs to; guid = a unique per-instance suffix). To decide whether
    the box for a given picture placeholder is present or missing, match
    the `PptcPromptBox-<idx>` prefix among the slide's shapes -- no
    guessing, and never a name collision on re-apply. Per placeholder:

    -   <if condition="the user PROVIDES an image for this placeholder (a file path or an upload)">
        INSERT it directly, do NOT write a prompt: `slide.fill` the picture
        placeholder with `image: { image: "<path>" }` (or, on a blank layout,
        `el.add` an image at the placeholder frame). This is image INSERTION,
        not generation -- it is allowed (see Non-Goals). Remove any stale
        `PptcPromptBox-<idx>` for that placeholder in the SAME ops document.</if>
    -   <if condition="the placeholder is empty, or its prompt box is still present">
        (re)write the prompt box.</if>
    -   <if condition="the user asks for a NEW prompt and the old box is already gone">
        that is the NORMAL post-generation state, not an error. Write a FRESH
        `img.prompts` box; never assume a box still exists, and never treat its
        absence as something to "restore".</if>
    -   <if condition="the placeholder already holds an IMAGE">
        that is the user's finished work. Do NOT write a box over it on your
        own. Only re-prompt on an explicit user request, and even then the box
        is ADDED as an overlay (`img.prompts` only pushes a shape, it never
        removes the image) -- the existing image stays untouched.</if>

    <for items="each picture placeholder you (re)prompt">
    Perform an individual creative step per placeholder (no template motifs):

    1.  Determine layout role + sidecar constraints (e.g. dark background
        on title/closing layouts with a white line when the style is
        illustration/render; calm motifs on background-image layouts).
        Collect ALL shapes that overlap the picture region, from TWO
        sources, because each sees only half of them:
        -   the placeholder's `overlays` from `tpl inspect` (template- and
            layout-level shapes: footer, slide number, chapter labels);
        -   the SLIDE's own shapes from `state --level full` -- any
            user-added textbox or element sitting over the picture is
            slide-local and does NOT appear in `tpl inspect`. Re-read it
            every turn so boxes the user added since the last prompt are
            honored.
        Use the picture placeholder's `coverage` (reported by `tpl
        inspect`/`describe`) to pick the mode -- see
        `references/prompt-formula.md` → "Background image vs. negative
        space":
        <if condition="coverage < 0.65 -- partial overlay">
        each overlapping shape becomes a NEGATIVE-SPACE clause ("the
        [region] is a vast empty [color] canvas creating significant
        negative space") and the subject moves to a free region.
        </if>
        <else>
        the whole frame is a true BACKGROUND image, not a subject: carry
        NO text (this OVERRIDES the title-text rule), keep ONE even tone (a
        dark backdrop with no bright hotspots, or a light backdrop with no
        dark blocks), and SET the overlay placeholder's text colour to
        contrast that tone — light text (`lt1`) on a dark backdrop, dark
        text (`dk1`) on a light one — via the `slide.fill` run colour, so
        the words stay legible.
        </else>
        NEVER explain why -- typographic words (title/footer/caption/
        label) make the image model render pseudo text. End prompts
        with "No text. No letters. No symbols." unless the prompt
        deliberately embeds quoted text (never on a background image).
    2.  Choose a motif that makes THIS slide's POINT tangible -- its
        message AND its takeaway/conclusion, leaning toward the resolution the
        slide concludes with, not just the problem -- and does not repeat a
        motif already used in the deck.
        <if condition="the slide message, its conclusion, deck topic and role do NOT determine a concrete, non-generic motif -- or a strong metaphor could plausibly go several clearly different ways">
        ASK the user ONE short, targeted question (offer 2-3 motif directions,
        or invite free text) BEFORE composing.</if>
        A wrong motif costs the user a whole image-generation cycle, so a quick
        question beats guessing; NEVER fall back to a generic filler motif just
        to avoid asking.
    3.  <if condition="this is a title-role image, or the scene carries text surfaces">
        embed the slide title/topic as short text-in-image.</if>
    4.  Compose the prompt per the formula, with the deck-wide style
        block and `#`-prefixed hex codes from <colors/> (primary
        `accent1` unless sidecar/user overrides). Do NOT put the aspect
        ratio in the prompt TEXT -- pptc derives it from the placeholder
        geometry and prints it in the box header (`IMAGE PROMPT · <ratio>`);
        the mirrored markdown header shows the same ratio.
    </for>

    Write all prompts into the deck via ONE `img.prompts` op per slide
    (prompt text per picture-placeholder idx) and `apply --rev <rev/>`.
    <if condition="a prompt box for that placeholder already exists (`PptcPromptBox-<idx>` among the slide's shapes)">
    `el.rm` it by name in the SAME ops document, BEFORE the `img.prompts`, so
    two boxes never stack on one placeholder.</if>
    Mirror each prompt to the user — emit one <expand name="image-prompt"/>
    per prompt:

    <define name="image-prompt">
    <template>
    ### 🎨 Image prompt · slide <n/> · placeholder <idx/> · <aspect/> · role <role/>

    ```
    <prompt/>
    ```
    </template>
    </define>

    Translate the fixed labels of this template into the USER's language
    (e.g. German: "Bild-Prompt — Folie … , Platzhalter … , Rolle …").

    </step>

8.  <step id="STEP 8: Report">

    Walk the final checklist in `references/content-rules.md`; list any
    open items. **Final integrity gate:** run `verify <deck>` once more on
    the finished deck; it must report `result.ok: true` (no findings) before
    you call the deck done -- a deck that would prompt a PowerPoint repair is
    never a finished deliverable. Then report with this output:

    <template>
    ────────────────────────────────────────
    ✅ **Deck updated** · `<deck/>` · rev `<rev/>`
    ────────────────────────────────────────
    - 🗂 Slides: <slide-count/> (<changed/> changed/new)
    - 🎨 Image prompts written: <prompt-count/>
    - 🛡 Integrity: <verify-result/>
    - 📋 Open checklist items: <open-items/>
    </template>

    Translate the fixed labels of this template into the USER's language.

    </step>

</flow>


Non-Goals
---------

-   **No image generation**: this skill writes prompts into the deck
    (`img.prompts` boxes); it never calls Gemini/Nano Banana or any
    other image API. Inserting an image the USER provides (a file/upload)
    is NOT generation -- place it directly (see STEP 7).
-   **No destroying generated images or user content** (see Execution
    Rules): prompt boxes are additive; a user-placed image is final and is
    re-prompted only on explicit request.
-   **No rewriting approved content**: when a ppt-prepare plan or the user
    supplies the wording, place it verbatim; paraphrasing, summarizing or
    shortening it (e.g. to fit capacity) happens only after explicit user
    confirmation, never as a build-time side effect.
-   **No web research**: deck content comes from the user's input and
    the conversation context only.
-   **No raw XML editing**: all mutations go through pptc ops.


Maintenance
-----------

`scripts/pptc.mjs` is a build artifact of the pptc project; the skill version
lives in `VERSION` (skill root, same for `ppt-prepare`). To update: run
`npm run plugin:sync` in the pptc repo (builds `dst/pptc.mjs`, copies it over
and writes both `VERSION` files) -- never edit the bundle by hand.
