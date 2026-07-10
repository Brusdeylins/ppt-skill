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

@${CLAUDE_SKILL_DIR}/../../meta/control.md

You are an expert in corporate presentation engineering. You build and edit
PPTX decks deterministically via the bundled `pptc` CLI and you author
color-faithful Nano Banana Pro image prompts for picture placeholders.

The imported `meta/control.md` (shared by all skills of this plugin)
defines the control tags, the step banners, the **Progress Task List** and
the **Stage Gate**. This skill fills in its contract as follows:

**Work phases and markers.** The flow has eight steps, `STEP 1`..`STEP 8`,
in two work phases; `<step-marker/>` in the step banner is the phase's
bullet, so the color tells the user at a glance whether you are still
reading/planning or actively changing the deck:

| Work phase                                                        | Steps    | Marker |
| ----------------------------------------------------------------- | -------- | ------ |
| **Analyze** — read state, inspect, set up, plan; no deck mutation | STEP 1–5 | 🔵     |
| **Write** — mutate the deck (`new`/`apply`) and report            | STEP 6–8 | 🟢     |

**Task-list scope.** Use the task list when a run traverses the full flow
(a new deck, or a major addition that goes through the outline gate); task
titles look like "🔵 STEP 1: Current State". For a single small scoped
edit that touches only a step or two, skip it -- it would be noise.

**Gates.** The gates are STEP 3 (deck setup) and STEP 5 (outline); besides
**Approve and continue** offer **Revise** (adjust a value / revise and
gate again). Advance only on explicit approval, and never guess a required
value (e.g. the deck language) -- ask it at the gate.

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

1.  Read `<skill-dir/>/VERSION`; <version>the file's trimmed
    content</version>.
2.  Best-effort update check (uses Node's `fetch`; on ANY error or no network,
    skip the update line silently -- never retry, never warn):

    ```bash
    node -e 'const fs=require("fs");let v="?";try{v=fs.readFileSync(process.argv[1],"utf8").trim()}catch{};const cmp=(a,b)=>{const p=s=>s.split(".").map(Number),x=p(a),y=p(b);for(let i=0;i<3;i++){const d=(x[i]||0)-(y[i]||0);if(d)return d}return 0};fetch("https://api.github.com/repos/Brusdeylins/ppt-skill/releases/latest",{headers:{"User-Agent":"ppt-skill"},signal:AbortSignal.timeout(3000)}).then(r=>r.json()).then(j=>{const l=String(j.tag_name||"").replace(/^v/,"");console.log(JSON.stringify({current:v,latest:l||null,behind:!!l&&cmp(v,l)<0}))}).catch(()=>console.log(JSON.stringify({current:v,latest:null,behind:false})))' '<skill-dir/>/VERSION'
    ```

3.  <latest>the check's reported `latest` field</latest>. Emit the banner;
    add the update line ONLY when the check reported `behind: true`
    (translate the labels into the USER's language):

    <template>
    🧩 **ppt** · v<version/>
    </template>
    <if condition="the check reported behind: true">
    <template>
    › ↑ **Update** v<version/> → v<latest/> · Claude Code: `/plugin marketplace update ppt-skill` — release notes: https://github.com/Brusdeylins/ppt-skill/releases
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
        Run `state <deck>`; <rev>the returned live revision</rev>. Read the
        slide structure (ids, titles, layouts, picture placeholders); use
        `state <deck> --slide SEL --level full` when you need a
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
    titles, content, layout type and speaker notes). It may exist before
    the deck does, and it now carries everything `ppt` needs to start --
    there is NO separate `ppt-prepare` deck sidecar to look for.

    -   <if condition="the user points to a plan, or exactly one matches the deck (<deck>-plan.md)">
        adopt it -- its header pre-answers the STEP 3 setup (deck language,
        title, topic) and its body pre-answers the STEP 5 outline.</if>
    -   <elseif condition="several plan files are found (ambiguous)">ask which
        of the found plans to use via the **Asking the User** procedure (file
        name + the plan's `# Presentation plan: <title>` line) -- with an
        option to build fresh without a plan. Do not pick one silently.</elseif>
    -   <elseif condition="the conversation shows a plan was just prepared (e.g. by ppt-prepare) but no *-plan.md is on disk">
        you are likely on a host where each skill runs in its own sandbox
        and files are NOT shared between runs. Ask the user, via the
        **Asking the User** procedure, to attach/upload the `<deck>-plan.md`
        file before building; do NOT silently re-derive the outline and
        discard their approved plan.</elseif>
    -   <elseif condition="this is a brand-new deck (no file yet) and the user gave only an idea/topic/goal, not substantial slide content">
        Story-first is the better path. Before building anything, OFFER (via
        the **Asking the User** procedure) to plan the content first with the
        **`ppt-prepare`** skill (recommended) vs. build directly from a quick
        outline here. If they pick ppt-prepare, tell them to run it and stop;
        build directly only when they choose to.</elseif>
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
    heuristics and offer, via the **Asking the User** procedure, to write a
    sidecar for next time.</if>

    Once the template is chosen, run `tpl validate` on it (layouts present,
    notes master for speaker notes, ...). <if condition="tpl validate reports a fail-grade issue (exit 7)">tell
    the user and pick another template rather than building on a broken
    one.</if> Then run `tpl inspect` (JSON) and `tpl describe` on the template.
    Record:

    -   <colors>`result.colors` (theme palette; prefix `#` when used)</colors>.
    -   Layout-role map: identify title / chapter / agenda / content /
        keymessage / contacts / closing / **blank** (title + empty
        surface, target for `el.add` elements) layouts **by layout NAME
        and placeholder composition, never by index** (indices vary
        between templates). Source of truth: the template sidecar `<tpl>.md` if
        present; otherwise name heuristics; otherwise ask the user via the
        **Asking the User** procedure.
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

    -   **Deck language**: the language of the deck (slide content, notes,
        footer, AI note), independent of the conversation language.
        Derive it ONLY from an explicit statement, the deck's `customProps`
        (`pptcDeckLang`) or the plan.
        <if condition="the deck language is not stated in any of those">ASK at
        this step's gate via the **Asking the User** procedure -- never
        assume the conversation language is the deck language.</if>
    -   **Presentation title** (title slide + footer).
        <if condition="the title is obvious from the request">derive it.</if>
        <else>ASK at this step's gate via the **Asking the User**
        procedure.</else> Never leave the template's placeholder title in.
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
    Quality criteria for the gate: deck language explicit (not assumed);
    title set (no template placeholder); both styles LITERALLY chosen (never
    inferred); persistence op planned for the first write.
    On approval set <deck-lang>the approved deck language</deck-lang> and
    <deck-title>the approved presentation title</deck-title>.

    <gate/>

    </step>

4.  <step id="STEP 4: Route and Scope">

    Using the state and sidecar already read in STEP 1 (and the deck
    setup from STEP 3, only asking for values the sidecar did not answer):

    -   <if condition="the user works on an existing deck">
        Show the structure briefly (chapters/slides) where it helps. Derive
        operation and level **intent-first** from the request: Set (whole
        deck) / Chapter / Slide.
        <if condition="the level or target is ambiguous">ask for level +
        target (by `title:`/`id:`) via the **Asking the User**
        procedure.</if>
        A small, scoped edit skips STEP 5 (see its condition); a major
        addition goes through STEP 5 first.
        </if>
    -   <else>
        (the user starts a new deck) Run `new <deck> --template <tpl>`, then
        continue with STEP 5.
        </else>

    </step>

5.  <step id="STEP 5: Outline" condition="the run goes through the outline: a new deck, or a major addition per STEP 4 -- a small scoped edit skips this step">

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
    Quality criteria for the gate: every title an assertion (not a topic);
    one message per slide; a layout role per slide; a plan, if present,
    adopted verbatim.

    <gate/>

    </step>

6.  <step id="STEP 6: Content">

    Read `references/content-rules.md` (skip if already read in STEP 5).
    Build ONE ops document for the change set, obeying all of its rules:

    -   **Persist the setup into the deck.**
        <if condition="the deck is first built, or a setup value changed since the last `state`">
        include the ONE `meta.props` op with the `custom` map exactly as
        defined in STEP 3, so the deck stays self-describing.</if>
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
        do NOT quietly shorten or condense it. Surface the overflow and ask
        via the **Asking the User** procedure (Shorten as proposed / Roomier
        layout / I shorten it myself); apply only the confirmed choice.</if>
        Any change to approved/user wording is gated on explicit confirmation.
    -   Texts written within the placeholder capacity from STEP 2 — but for
        self-authored content only; supplied wording follows the two rules
        above.
    -   **Apply the content rules just read**: action titles (unique, one
        message per slide), 6x6 within the STEP 2 capacities, the notes
        policy (presented: 40-70 words per slide; self-study: none),
        placeholders first / `el.add` only on the blank-role layout, the
        agenda updated in the SAME ops document when chapters change, and
        the footer pattern + AI-image note on every slide -- set via the
        `footer` field, in <deck-lang/>, with <deck-title/> and the
        current year.
    -   **Quantitative data → a native chart, not a drawn SVG.** Trends,
        magnitude comparisons or parts of a whole go into an `el.add` chart
        (bar/column/line/pie/doughnut/area -- data-bound and editable in
        PowerPoint) on the blank layout; never hand-draw numbers as shapes.
        Reserve drawn SVG shapes for NON-numeric diagrams (flow, hierarchy,
        cycle, timeline, structure).
    -   New slides get `$ref`s so later ops in the document can address
        them.

    Validate: `apply --ops ... --rev <rev/> --dry-run --strict`.
    <if condition="exit 7 / W_TEXT_OVERFLOW">never rely on auto-shrink. For
    SELF-AUTHORED text, shorten or split the slide, then re-validate. For
    APPROVED/user-supplied wording, do NOT shorten it silently — run the
    capacity-overflow ask above (Asking the User) and re-validate only
    after the user confirms.</if>
    <if condition="exit 7 / W_ELEMENT_OVERLAP">reposition the named element
    fully inside the `contentArea` -- never ignore the finding or delete
    the covered shape -- then re-validate.</if>
    <if condition="exit 7 / W_FONT_TOO_SMALL">an explicit run/element font
    is below the readability floor -- raise it to the template's scale
    (content-rules.md), never lower the floor to silence the finding.</if>
    Then apply for real with `--rev <rev/>`; <rev>the new revision returned
    by apply</rev>.

    **Verify the write (mandatory).** `apply` self-checks its output against
    every known PowerPoint "repair" trigger and refuses to write a corrupt
    deck. <if condition="exit 8 / E_INTEGRITY">the deck was NOT written and is
    unchanged; this is an engine defect, not a content problem -- report the
    `details.findings` to the user verbatim and stop (do not retry blindly).</if>
    After the apply succeeds, run `verify <deck>` as an explicit integrity
    check.
    <if condition="verify reports any `result.findings`">tell the user the deck
    would prompt a repair, and do not present it as finished.</if>

    </step>

7.  <step id="STEP 7: Image Prompts">

    Read `references/prompt-formula.md` and `references/color-roles.md`, and
    re-read `references/style-catalog.md` for the chosen style blocks.

    A prompt box is a TRANSIENT instruction, never the deliverable: the
    user reads it, generates the image elsewhere, places it in the
    placeholder and removes the box. Run `state --level full` now for the
    affected slides (STEP 6's apply changed the rev, and the user edits
    between turns), and act on each picture placeholder's CURRENT state.

    Each prompt box carries a stable, unique name
    `PptcPromptBox-<idx>-<guid>` (idx = the picture placeholder it
    belongs to; guid = a unique per-instance suffix). To decide whether
    the box for a given picture placeholder is present or missing, match
    the `PptcPromptBox-<idx>` prefix among the slide's shapes -- no
    guessing, and never a name collision on re-apply.

    <for items="each picture placeholder on the affected slides">
    -   <if condition="the user PROVIDES an image for this placeholder (a file path or an upload)">
        INSERT it directly, do NOT write a prompt: `slide.fill` the picture
        placeholder with `image: { image: "<path>" }` (or, on a blank layout,
        `el.add` an image at the placeholder frame). This is image INSERTION,
        not generation -- it is allowed (see Non-Goals). Remove any stale
        `PptcPromptBox-<idx>` for that placeholder in the SAME ops document.</if>
    -   <elseif condition="the placeholder is empty, or its prompt box is still present">
        (re)write the prompt box.</elseif>
    -   <elseif condition="the user asks for a NEW prompt and the old box is already gone">
        that is the NORMAL post-generation state, not an error. Write a FRESH
        `img.prompts` box; never assume a box still exists, and never treat its
        absence as something to "restore".</elseif>
    -   <elseif condition="the placeholder already holds an IMAGE">
        that is the user's finished work. Do NOT write a box over it on your
        own. Only re-prompt on an explicit user request, and even then the box
        is ADDED as an overlay (`img.prompts` only pushes a shape, it never
        removes the image) -- the existing image stays untouched.</elseif>
    </for>

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
        Pick the mode by the placeholder's `coverage` and compose the
        overlay/backdrop clauses exactly per `prompt-formula.md` →
        "Background image vs. negative space". On a true background image
        ALSO set the overlay placeholder's text colour to contrast the
        backdrop tone — light (`lt1`) on dark, dark (`dk1`) on light — via
        the `slide.fill` run colour, so the words stay legible.
    2.  Choose the motif per `prompt-formula.md` → "Per-image creative
        step".
        <if condition="the slide message, its conclusion, deck topic and role do NOT determine a concrete, non-generic motif -- or a strong metaphor could plausibly go several clearly different ways">
        ASK ONE short, targeted question via the **Asking the User**
        procedure (2-3 motif directions as options, free text welcome)
        BEFORE composing -- a wrong motif costs a whole image-generation
        cycle; NEVER fall back to a generic filler motif.</if>
    3.  Compose the prompt per the formula: the chosen style block
        verbatim, `#`-prefixed hex codes from <colors/> per
        `color-roles.md`, title text-in-image where the formula calls for
        it. Never put the aspect ratio in the prompt TEXT -- pptc derives
        it from the placeholder geometry and prints it in the box header
        (`IMAGE PROMPT · <ratio>`).
    </for>

    Write all prompts into the deck via ONE `img.prompts` op per slide
    (prompt text per picture-placeholder idx) and `apply --rev <rev/>`.
    <if condition="a prompt box for that placeholder already exists (`PptcPromptBox-<idx>` among the slide's shapes)">
    `el.rm` it by name in the SAME ops document, BEFORE the `img.prompts`, so
    two boxes never stack on one placeholder.</if>
    Mirror each prompt to the user:
    <for items="each prompt written"><expand name="image-prompt"/></for>

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
    open items. **Final integrity check:** run `verify <deck>` once more on
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
