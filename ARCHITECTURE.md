# pptc Architecture

## Layer model

Dependencies point strictly downward; no module knows a layer above it.

```
cli/        entry, dispatch, envelope rendering, exit-code mapping
  ↓
commands/   one module per command -- thin orchestration only
  ↓
engine/     OOXML adapters: reader, in-place engine, PptxGenJS element
            builders, package-wide steps, self-verify, content guard
  ↓
core/       pure domain logic: model, selectors, ops planning,
            geometry→semantics, lint
  ↓         (no I/O, no engine types -- purity enforced by the module graph)
infra/      file system, payload resolution, hashing, version constant,
            arg parsing, error taxonomy
```

`schema/` (the Zod payload and ops schemas) sits beside `infra/` as a
second leaf: every layer above may import it, it imports nothing but zod.

## Modules

| Module | Responsibility |
|---|---|
| `cli/main.ts` | dispatch, one JSON envelope on stdout, exit codes |
| `cli/help.ts` | usage text and the per-command detailed help registry |
| `commands/*` | per-command orchestration (5-40 lines each); `apply.ts` owns `executeOps`, the single write path also used by `new` and the sugar commands |
| `schema/payloads.ts` | Zod: rich text, table, chart, frame, elements |
| `schema/ops.ts` | Zod: the ~10 ops and the ops document |
| `core/model.ts` | engine-free domain model (Frame, Layout, DeckState, ...) |
| `core/selector.ts` | selector grammar `id:/title:/index:/$ref` + resolver |
| `core/describe/*` | pure geometry→semantics: position bands, ASCII minimap, capacity model, Markdown narration |
| `core/lint.ts` | capacity lint over the describe model |
| `core/ops/registry.ts` | op contract (`plan(ctx, op)`) and the MutationPlan model |
| `core/ops/*` | one plan transformer per op (slide-add, slide-edit, elements) |
| `core/ops/fill-common.ts` | the shared fill payload (placeholders, notes, footer, background, hidden) planned identically for `slide.add` and `slide.fill` |
| `core/ops/planner.ts` | registry assembly, rev check, whole-document planning |
| `engine/reader.ts` | read-only OOXML: TemplateInfo, DeckState, rev hashing |
| `engine/seed.ts` | slides and decks from layouts: `slideFromLayout` (empty slide with cloned placeholders `PptcPh-<idx>`), `buildEmptyDeck` for `new` |
| `engine/inplace/engine.ts` | `runSession`: package repairs, slide structure, slide classification (touched, repairs), one pass per touched slide, slide list, package-wide steps, self-checks (verify, content guard, byte guard), atomic write; `preflight` for `--dry-run` |
| `engine/inplace/package.ts` | part-level package access; rel-graph clone under fresh names (`slide.copy`, transplanted parts) |
| `engine/inplace/slides.ts` | everything done to ONE slide in a single parse/serialize pass: removals, text, picture placeholders, background, footer, notes, transplanted elements, hyperlinks, visibility, links to removed slides, repairs |
| `engine/inplace/repair.ts` | the repairs of the triggers `defects.ts` detects, on any part: duplicate/dead relationships, references to missing relationships (smallest valid removal, named), duplicate shape ids |
| `engine/inplace/transplant.ts` | `el.add`: one PptxGenJS deck per apply (one generator slide per target slide), shapes moved by object name with their parts |
| `engine/text.ts` | DOM rich-text builder (`a:p`/`a:r`/`a:t`) |
| `engine/elements.ts` | ElementSpec → PptxGenJS calls on a generator slide |
| `engine/post.ts` | package-wide steps: sweep unreferenced notes and stale presentation refs, doc props, asset GC, sections, content types, app properties |
| `engine/defects.ts` | the repair triggers of a single part (duplicate or dead relationships, references to missing relationships, duplicate shape ids, notes ownership) -- ONE rule set: the engine detects its input defects with it and repairs exactly these, verify reports them |
| `engine/verify.ts` | read-only integrity check: the part rules of `defects.ts` on every part plus the package-level rules (slide list, content types, notes 1:1, app properties, orphans); backs `pptc verify` and the self-verify before every write (`E_INTEGRITY`) |
| `engine/guard.ts` | content self-check and byte guard before every write: each touched slide must keep its notes and related parts (images, charts, embeddings) unless an op changes them, layouts/masters must not change (`E_INTEGRITY`) |
| `engine/xml.ts` | shared XML DOM helpers: parse/serialize, OOXML namespaces, element queries, relationship references, escaping |
| `engine/parts.ts` | the OPC layer below every engine module: case-insensitive part names, relationships read through the DOM, the relationship-type taxonomy, content types, part removal, asset GC, layout order, and `PartAllocator`: every fresh part name, relationship id, sldId and shape id of an apply |
| `infra/args.ts` | `util.parseArgs` wrapper with usage errors |
| `infra/errors.ts` | `PptcError` with stable codes and the exit-code table |
| `infra/fs.ts` | `@file`/stdin payloads, atomic write, scratch dir, hashing |
| `infra/version.ts` | the CLI version constant |

## Data flow: `apply`

```
args        --ops @file → raw JSON
schema      OpsDocumentSchema.parse           → E_SCHEMA with issue paths
reader      DeckState + deck layouts (+ template layouts when --template)
planner     rev check → per-op plan transformers → MutationPlan
lint        capacity warnings (E_LINT under --strict)
            ── everything above is side-effect-free; --dry-run stops here ──
repair      repair triggers of the input outside the slides (presentation,
            masters, layouts, themes) and slide parts outside the slide list
            are repaired first (W_REPAIRED) -- before the engine writes
structure   in place: kept slides stay where they are, copies are cloned
            from the unchanged input, new slides come from the deck's own
            layout parts; fresh names and ids from the PartAllocator
touch       a kept slide is touched when an op edits or (un)hides it, it
            links to a removed slide, or its content carries repair triggers
            or a notes part of another slide (W_REPAIRED, engine/defects);
            new and copied slides are always touched; untouched slides are
            never written
generate    PptxGenJS renders every el.add element of the apply into one
            in-memory deck
slides      repairs of the slide's other parts (notes, charts, ...), then
            one pass per touched slide: removals, text, picture placeholders,
            background, footer, notes, transplanted elements, hyperlinks,
            visibility, links to removed slides, repairs
package     sldIdLst (kept ids, fresh ids for new and copied slides), removed
            slides and links to them, then engine/post.ts: notes sweep, doc
            props, asset GC, sections, content types, app properties
verify      self-check the produced bytes against the known repair triggers,
            every touched slide against its source (content guard) and every
            part of every untouched slide against the input (byte guard); any
            finding the input did not already carry aborts with
            E_INTEGRITY, nothing is written (inherited: W_INPUT_DEFECT)
write       tmp file + rename (atomic), re-read → rev.after + ref→id map
```

## Engine constraints (why it is built this way)

- **In place, not rebuilt.** 1.0.x rebuilt the whole deck through
  pptx-automizer on every apply; its count-based, slide-local part and id
  allocation moved media, notes, OLE objects, theme overrides, SmartArt,
  tags, links and sections between slides. The in-place engine touches only
  what the plan asks for: an untouched slide's parts are never parsed for
  writing, renamed or re-serialized, and the byte guard proves it on every
  apply.
- **One allocator.** Every fresh part name continues above the highest
  number of its family in the whole package, every relationship id is a
  fresh `rIdPptc<N>`, every sldId is above every id of the input deck (an
  id removed in the same apply is never re-used; across applies the next id
  is the highest one in use plus one, as in PowerPoint -- `--rev` guards
  against stale `id:` selectors). Nothing is ever named from a count.
- **Slides from the deck's own layouts.** A new slide is the layout's
  placeholders cloned onto an empty slide (normalized names `PptcPh-<idx>`;
  footer, slide-number and date placeholders WITH their layout content).
  Decks always carry their masters/layouts (`new` bakes them in from the
  template), so no layout is ever imported; `--template` on `apply` only
  checks that the template's layouts equal the deck's.
- **Copies own their parts.** `slide.copy` clones the slide with its notes,
  charts (with workbook, colors, style), embeddings, SmartArt, tags,
  comments and theme overrides under fresh names; media, the layout and
  external targets stay shared (media are immutable).
- **PptxGenJS generates, the engine places.** All `el.add` elements of an
  apply are rendered into one PptxGenJS deck; each shape is moved into its
  target slide by its `<name>-<uuid>` object name, with the parts it
  references copied under fresh names. `pptxgenjs` is pinned exactly
  (4.0.1) so generated XML stays stable.
- **Repairs where pptc writes, and on package parts.** BOM stripping,
  absolute-to-relative targets and repair-trigger fixes happen on touched
  slides only; a slide whose input carries repair triggers counts as
  touched, is repaired and reported (`W_REPAIRED`). Parts outside every
  slide's content (presentation, masters, layouts, themes, package rels,
  document properties, custom XML) are repaired on every apply before the
  engine writes, and kept consistent (content types, app properties).
  Fresh part names never reuse a name a relationship already points at,
  dead or alive.
- **One rule set for defects.** `engine/defects.ts` defines the repair
  triggers of a part once; the engine detects input defects with it,
  `inplace/repair.ts` repairs exactly these, `verify.ts` reports them. A
  new trigger is added there, with its repair in `repair.ts`. A dead
  reference takes down the smallest element that keeps the XML valid (a
  link, a picture bullet or fill, an extension, a sound, else its shape);
  `W_REPAIRED` names what was removed.
- **Relationships through the DOM.** Relationships, content types and the
  slide list are read and written through the DOM (`parts.ts`
  `relTargets`/`readRels`/`relMap`), never by regex over raw XML --
  attribute quoting and spelling vary between producers. A duplicate
  relationship id resolves to its first relationship everywhere. Only
  `docProps/app.xml` is edited as text (element content, no attributes).
  Part names are resolved case-insensitively by every accessor
  (`findPart`), as OPC demands, and percent-escaped targets resolve to
  their part. XML parts are decoded as UTF-8 or UTF-16 by their byte-order
  mark. Findings carry their rule and part (`defects.ts` `Finding`), so the
  baseline below recognizes an inherited trigger whatever its wording.
- **Only the engine's own triggers block a write.** The output is verified
  against every known repair trigger; a finding the input deck already
  carried and pptc cannot repair (a slide listed twice, malformed custom
  XML, ...) stays as it was and is reported (`W_INPUT_DEFECT`) -- only
  findings the apply would add (by rule and part, counted) fail with
  `E_INTEGRITY`. A repair returns the relationships only the content it
  removed used; the content guard leaves exactly those out.
- **stdout is shielded during the engine pass.** A library may log via
  `console.log`; the engine diverts console output to stderr so the
  one-JSON-envelope contract on stdout holds.

## Extending: a new op

1. Add the Zod schema in `schema/ops.ts` and include it in `OpSchema` /
   `OP_NAMES`.
2. Create the plan transformer in `core/ops/` fulfilling `OpHandler`
   (`plan(ctx, op)` rewrites the MutationPlan; never touch files).
3. Register it in `core/ops/planner.ts` (`HANDLERS`).
4. If the op introduces a new *kind* of mutation, teach
   `engine/inplace/slides.ts` (per slide) or `engine/post.ts` (package-wide)
   to interpret it, and mark the slide as touched in
   `engine/inplace/engine.ts`.
5. Add a unit test for the transformer and extend the integration roundtrip.

## Testing strategy

- `test/unit/` -- pure core with geometry fixtures, no files
- `test/golden/` -- `tpl describe` Markdown pinned as snapshot (the LLM-facing
  contract), and the engine results of fixed op sequences over all ten ops
  (per-slide fingerprints frozen from the 1.0.11 engine; the demo deck's and
  the private templates' goldens stay local in `private-templates/golden/`)
- `test/integration/` -- real write path against the committed neutral
  template fixture, incl. all-or-nothing and rev semantics
- `test/integration/untouched.test.ts` -- every part of every untouched
  slide keeps its bytes across edits, additions, copies, moves and removals
- `test/integration/fuzz-semantic.test.ts` -- seeded random apply sequences
  against an expected model (title, notes, image, chart, OLE payload,
  section, sldId) after every apply; long runs via `PPTC_FUZZ_ROUNDS` and
  `PPTC_FUZZ_SEEDS`
- `test/integration/stress.test.ts` -- agent-like sessions: many applies in
  a row covering every op kind and element type, plus seeded-random edit
  combinations (reproducible LCG); planner rejections are fine, a
  written-but-broken file is not
- `test/contract/` -- the built bundle's envelope shape, exit codes and the
  sugar/read-command happy paths
- `test/util/integrity.ts` -- the file-integrity validator asserting every
  written deck against the known PowerPoint "repair" triggers (well-formed
  XML, unique shape ids, no dead relationship targets, no dead/duplicate
  content-type overrides, full content-type coverage, resolvable sldIds).
  Every test that writes a deck runs `expectIntact`; the stress suite runs
  it after EVERY apply.

The fixture (`test/fixtures/neutral-template.pptx`) carries the eleven
standard Office layouts, five of them renamed to the role names the tests
use (DEFAULT blank, TITLE_SLIDE, CONTENT, TWO_COLUMN, PICTURE with a picture
placeholder), unique placeholder names per layout and passes `tpl validate`.
