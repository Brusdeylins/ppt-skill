Control Tags
============

The skills of this plugin use lightweight XML-style control tags inside
their Markdown. They are instructions to the executing agent, not output.
Honor them exactly as defined here; this file is self-contained and
independent of any other plugin or MCP server.

This file defines only the generic machinery. Where a definition below
defers to "the skill", the importing SKILL.md supplies the specifics: the
flow's steps, the work-phase table (which steps carry which marker color),
which steps carry a `<gate/>`, the gate's option set, and the task-list
scope.


Placeholders
------------

-   `<xxx>value</xxx>` *sets* the placeholder named `xxx` to `value`.
    Expands to nothing; do not output anything.

-   `<xxx/>` *reads* the placeholder named `xxx` and expands to its
    current value.


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
    to the current item. A `<break/>` stops the repetition early.

-   `<while condition="...">...</while>`:
    Repeat the body as long as the condition is met. A `<break/>`
    stops the repetition early.


Output Templates
----------------

-   `<template>...</template>`:
    The ONLY mechanism for user-visible output. Output the template
    content *exactly* as given (including newlines), expanding only
    control constructs and `<xxx/>` placeholders, and removing
    trailing spaces. Do not output explanations or summaries of your
    own unless a template requests them.


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
the separator, a thin `────────` rule to open a section (user choices are
framed by a rule above AND below), `›` before a prompt line. No drawn
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

Emit the banner exactly once per step entry; when a step is skipped by an
`<if>`/route condition, do not emit its banner. The skill's phase markers
are the only decoration -- do not invent other status glyphs or colors.


Progress Task List
------------------

So the user always sees where they are in the flow, maintain a visible
task list of the steps via the host's task-list facility WHEN THE HOST
PROVIDES ONE (in Claude Code: the Task tools -- TaskCreate / TaskUpdate /
TaskList). Where the host has no task-list facility (e.g. the claude.ai
web chat), SKIP the task list entirely -- do not fake it as text; the step
banner and the gate question orient the user on their own:

1.  **At flow start**, create one task per `<step>` of the `<flow>`, in
    order, each titled with the step's marker and `id`. Create the list
    once per run; reuse an existing list rather than duplicating.
2.  **On entering a step**, set its task to in_progress (right after the
    step banner); **when the step finishes** -- its `<gate/>` approved, or
    for a gateless step its body done -- set it completed and move the next
    step to in_progress.
3.  **A step skipped** by an `<if>`/route condition is marked completed
    with a "skipped" note, so the map stays honest.
4.  **Run scope:** every run that traverses the flow gets the list; where
    the skill's SKILL.md exempts small scoped runs, skip it there -- it
    would be noise.

The task list mirrors the flow as an always-visible map; it never replaces
the step banner or the gate's question -- it sits beside them.


Asking the User
---------------

Whenever a step needs the user to choose -- a `<gate/>`, a template or style
pick, a disambiguation -- ask with ONE consistent, harness-portable
procedure. Do NOT call a host-specific selection tool (e.g. Claude Code's
`AskUserQuestion`); render the choice as plain Markdown so the skill works
identically in every LLM harness. NEVER end a turn on a half-asked question
(a bare colon or a trailing "..." with no options):

1.  Frame it as a short question plus 2-4 options, each a
    `Label — one-line description`.
2.  Render the choice framed by a thin rule above and below (so the
    decision stands out from the surrounding prose), then END the turn
    and wait for the reply:

    <template>
    ────────────────────────────────────────
    ❓ **<the question>**

    1. **<label>** — <description>
    2. **<label>** — <description>

    › _Reply with a number — or your own answer._
    ────────────────────────────────────────
    </template>

3.  Map the reply back to an option: a number picks that option; a label
    (or an unambiguous prefix of one) picks its option; any other free
    text is an "Other" answer -- act on it.
4.  <if condition="the user declines or cancels">treat it as Cancel: do not
    advance; ask what they want instead.</if>


Stage Gate
----------

-   `<gate/>`:
    A BLOCKING checkpoint at the end of a `<step>` -- the skill's `<flow>`
    places it at its decision points. The flow does NOT advance past it
    until the user explicitly approves. Run it like this:

    1.  Emit a **checkpoint** in the clean style: a header line
        `◆ **Checkpoint · <step-id>**`, a short summary of what the step
        produced (its key values), then the quality criteria as a list
        (`✅` met / `⬜` not met).
    2.  Ask, via the **Asking the User** procedure, offering at least
        **Approve and continue** and a revise option (stay in this step,
        apply the change, and gate again). The skill's SKILL.md fixes the
        exact option set and labels, and defines the protocol behind any
        further option (e.g. a skip).
    3.  Act on the choice: on *approve* advance to the next `<step>`; on
        *revise* stay here, apply the change, and gate again; any further
        option follows the protocol the skill defines for it.

    Never advance past a `<gate/>` by assuming or inferring an unanswered
    value -- an unresolved required value is asked at the gate, not
    guessed. Stay inside the current step until its gate is approved; do
    not batch two gated steps into one gate. Steps without a `<gate/>`
    proceed normally.
