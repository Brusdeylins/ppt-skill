Control Tags
============

The skills of this plugin use lightweight XML-style control tags inside
their Markdown. They are instructions to the executing agent, not output.
Honor them exactly as defined here; this file is self-contained and
independent of any other plugin or MCP server.

This file defines only the generic machinery. Where a definition below
defers to "the skill", the importing SKILL.md supplies the specifics: the
flow's steps, the work-phase table (which steps carry which marker color),
which steps carry a `<gate/>`, the gate's option set and quality
criteria, and the task-list scope.


Placeholders
------------

-   `<xxx>value</xxx>` *sets* the placeholder named `xxx` to `value`.
    Expands to nothing; do not output anything.

-   `<xxx/>` *reads* the placeholder named `xxx` and expands to its
    current value. A read whose placeholder was never set is *derived*:
    fill it from the step's own results (never from a guess about the
    user's intent).

BARE angle-bracket tokens (no trailing `/>`) inside code spans or command
lines (`state <deck>`, `<name>.md`) are literal CLI/documentation syntax,
NOT control tags. A self-closing `<xxx/>` READ expands everywhere,
including inside commands and code (`--rev <rev/>`). A multi-word
descriptive token (`<what this step does, one short clause>`) is a
fill-in slot for you to write.


Objective
---------

-   `<objective>...</objective>`:
    The skill's single overarching goal. Read it as binding intent, not
    output. Expands to nothing.


Flow Constructs
---------------

-   `<flow>...</flow>`:
    A *sequential flow* of `<step>`s which MUST be executed in exactly
    the given order. Expands to its body.

-   `<step id="<id/>" [condition="..."]>...</step>`:
    One distinct step of a `<flow>`. Execute its body completely --
    including any closing `<gate/>` -- before moving to the next step.
    With a `condition`, run the step only when it holds, else skip it.
    Expands to its body.

-   `<if condition="...">...</if>`, followed optionally by
    `<elseif condition="...">...</elseif>` and/or `<else>...</else>`:
    Conditional expansion. Expand the body of the first branch whose
    condition is met, otherwise the `<else>` body, otherwise nothing.

-   `<for items="...">...</for>`:
    Repeat the body once per item; inside the body, `<item/>` expands
    to the current item. When the items are structured, the body reads an
    item's fields as `<field/>` placeholders (e.g. `<title/>`, `<n/>`).
    A `<break/>` stops the repetition early.

-   `<while condition="...">...</while>`:
    Repeat the body as long as the condition is met. A `<break/>`
    stops the repetition early.


Output Templates
----------------

-   `<template>...</template>`:
    Fixed-form user-visible output. Where a template exists, output its
    content *exactly* as given (including newlines), expanding only
    control constructs and `<xxx/>` placeholders, and removing trailing
    spaces; a construct occupying whole lines consumes its trailing
    newline. Steps that PROPOSE or draft content (messages, storylines,
    summaries) present it as normal prose -- but never replace an
    existing template with prose of your own.


Reusable Blocks
---------------

-   `<define name="x">...</define>`:
    Names a reusable block `x`. Expands to nothing.

-   `<expand name="x"/>`:
    Expands to the body of the matching `<define name="x">` -- use it to
    emit a block in more than one place without repeating it.


Step Announcement
-----------------

All user-facing output uses ONE clean style: an icon + **bold** label, `·` as
the separator, a thin `────────` rule to open a section (OPEN asks for user
input are framed by a rule above AND below; choice asks are an
AskUserQuestion dialog, not text), `›` before a prompt line. No drawn
boxes, no ASCII art.

The moment a `<step>` begins executing, FIRST emit a one-line banner so
the user can see exactly where you are, THEN carry out the step:

    ────────────────────────────────────────
    <step-marker/> **<step-id/>** · <what this step does or decides, one short clause>

`<step-id/>` is the step's `id`. `<step-marker/>` is the colored bullet of
the step's WORK PHASE, from the phase table in the skill's SKILL.md (each
phase maps a range of steps to one marker color). The color tells the user
at a glance which kind of work is active; the step id in the banner names
the exact step within the phase.

Emit the banner exactly once per step entry; when a step is skipped by
its `condition`, an `<if>` or a route, do not emit its banner. In the
step banner the phase marker is the only decoration -- invent no glyphs
beyond those this file and the skill's own templates define.


Progress Task List
------------------

So the user always sees where they are in the flow, maintain a visible
task list of the steps via the host's task-list facility WHEN THE HOST
PROVIDES ONE (in Claude Code: the Task tools -- TaskCreate / TaskUpdate /
TaskList). Where the host has no task-list facility (e.g. the claude.ai
web chat), SKIP the task list entirely -- do not fake it as text; the step
banner and the gate question orient the user on their own:

1.  **At flow start**, create one task per `<step>` of the `<flow>`, in
    order. A task's subject is the step's marker and `id` followed by the
    step's work in IMPERATIVE form; where the facility carries an active
    form (Claude Code: `activeForm`), phrase that in PRESENT CONTINUOUS --
    e.g. subject "🔵 PHASE 1: Capture the briefing", activeForm
    "Capturing the briefing". A required description field (Claude Code:
    `description`) carries what the step produces, in one line. Where the
    facility supports dependencies (Claude Code: TaskUpdate's
    `addBlockedBy`), chain each task as blocked by its predecessor, so
    the list also SHOWS the flow's order. Create the list once per run;
    reuse an existing list rather than duplicating.
2.  **On entering a step**, set its task to in_progress (right after the
    step banner); **the moment the step finishes** -- its `<gate/>`
    approved, or for a gateless step its body done -- set it completed.
    Never batch completions across steps, and exactly ONE task -- the
    current step -- is in_progress at any time while the flow runs.
3.  **A step skipped** by its `condition`, an `<if>` or a route is set
    completed with "(skipped)" appended to its subject -- the subject is
    what the list view shows -- so the map stays honest.
4.  **Run scope:** every run that traverses the flow gets the list; where
    the skill's SKILL.md exempts small scoped runs, skip it there -- it
    would be noise.

The task list mirrors the flow as an always-visible map; it never replaces
the step banner or the gate's question -- it sits beside them.


Asking the User
---------------

EVERY time a step needs input from the user -- a `<gate/>`, a pick, a
disambiguation, a clarification, a missing value -- the work STOPS at ONE
explicit ask (a blocking AskUserQuestion dialog, or an open frame ending
the turn), so the user always SEES that action is required. NEVER leave a
half-asked question (a bare colon or a trailing "..." with no options),
and never bury a question in prose:

1.  PREFER a choice, and present it with the **AskUserQuestion tool** --
    never as a plain-text question list. Fill the tool's fields like
    this:
    -   `question`: the complete decision question, ending in "?".
    -   `header`: a short chip naming the decision, at most 12
        characters (e.g. "Template", "Deck lang", "Motif").
    -   `options`: 2-4 entries, each a `label` of 1-5 words plus a
        `description` explaining what the choice means or trades off.
        You PROPOSE, the user DECIDES: put the recommended option FIRST
        and append "(Recommended)" to its label. Do NOT add an "Other"
        option -- the UI appends a free-text "Other" automatically.
    -   `multiSelect`: true only when the choices are not mutually
        exclusive (e.g. "which sections to include").
2.  Only when options would be artificial (a genuinely open question --
    a name, a fact only the user knows), ask in plain Markdown with this
    frame, listing at most ~3 short sub-points of what is needed, then
    END the turn:

    <template>
    ────────────────────────────────────────
    ❓ **<the question>**

    - <needed value, one line>
    - <needed value, one line>

    › _<what kind of answer is expected>_
    ────────────────────────────────────────
    </template>

3.  ONE ask per turn: bundle related sub-questions into ONE
    AskUserQuestion call (it carries up to 4 questions) or ONE open
    frame, instead of scattering questions through the prose. Everything
    before the ask is context; everything the user must decide lives in
    the ask.
4.  Act on the reply: a selected option is the decision; free text (the
    UI's "Other", or any reply to an open frame) is acted on as given. A
    partial answer to an open ask is fine: adopt what was given and
    propose the rest.
5.  <if condition="the host does not provide the AskUserQuestion tool (e.g. a plain chat harness)">
    render the choice in the open frame above with numbered options
    (`1. **<label>** — <description>`) and map the reply back: a number
    or an unambiguous label prefix picks that option; other text is an
    "Other" answer.</if>
6.  <if condition="the user declines or cancels">treat it as Cancel: do not
    advance; ask what they want instead.</if>


Stage Gate
----------

-   `<gate/>`:
    A BLOCKING checkpoint at the end of a `<step>` -- the skill's `<flow>`
    places it at its decision points. The flow does NOT advance past it
    until the user explicitly approves. Run it like this:

    1.  Emit a **checkpoint** in the clean style, opened by the thin
        section rule: a header line `◆ **Checkpoint · <step-id/>**`, a
        short summary of what the step produced (its key values), then
        the quality criteria as a list (`✅` met / `⬜` not met).
    2.  Ask via the **Asking the User** procedure -- i.e. an
        AskUserQuestion dialog, one option per outcome -- offering at
        least **Approve and continue** and a revise option (stay in this
        step, apply the change, and gate again). The skill's SKILL.md
        fixes the exact option set and labels, and defines the protocol
        behind any further option (e.g. a skip).
    3.  Act on the choice: on *approve* advance to the next `<step>`; on
        *revise* stay here, apply the change, and gate again; any further
        option follows the protocol the skill defines for it.

    Never advance past a `<gate/>` by assuming or inferring an unanswered
    value -- an unresolved required value is asked at the gate, not
    guessed: the gate's ask then asks for that VALUE (instead of Approve);
    re-gate once it is supplied. Stay inside the current step until its
    gate is approved; do not batch two gated steps into one gate. Steps
    without a `<gate/>` proceed normally.
