# pptc Architecture

## Layer model

Dependencies point strictly downward; no module knows a layer above it.

```
cli/        entry, dispatch, envelope rendering, exit-code mapping
  ↓
commands/   one module per command -- thin orchestration only
  ↓
engine/     OOXML adapters: reader, seed factory, automizer session,
            PptxGenJS element builders, zip post-pass, self-verify
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
| `engine/seed.ts` | seed-deck factory (one empty slide per layout), content-hash cache |
| `engine/session.ts` | automizer pass + post-pass + atomic write |
| `engine/text.ts` | DOM rich-text builder (`a:p`/`a:r`/`a:t`) |
| `engine/elements.ts` | ElementSpec → PptxGenJS calls (automizer interop) |
| `engine/post.ts` | zip post-pass: GC, notes, footer, background, slide visibility (`show="0"`), placeholder images, hyperlink rels, doc props, repair-trigger cleanup (stale rels, shape ids) |
| `engine/verify.ts` | read-only integrity check against the known PowerPoint "repair" triggers; backs `pptc verify` and the self-verify before every write (`E_INTEGRITY`) |
| `engine/xml.ts` | shared XML DOM helpers: parse/serialize, OOXML namespaces, element queries |
| `engine/parts.ts` | zip part access helpers (read/parse archive parts) |
| `infra/args.ts` | `util.parseArgs` wrapper with usage errors |
| `infra/errors.ts` | `PptcError` with stable codes and the exit-code table |
| `infra/fs.ts` | `@file`/stdin payloads, atomic write, cache dir, hashing |
| `infra/version.ts` | the CLI version constant |

## Data flow: `apply`

```
args        --ops @file → raw JSON
schema      OpsDocumentSchema.parse           → E_SCHEMA with issue paths
reader      DeckState + deck layouts (+ template layouts when --template)
planner     rev check → per-op plan transformers → MutationPlan
lint        capacity warnings (E_LINT under --strict)
            ── everything above is side-effect-free; --dry-run stops here ──
session     automizer: rebuild slide list (kept slides re-imported from the
            deck itself, new slides imported from the template's seed deck),
            callbacks apply DOM text + generated elements
post-pass   zip level: GC orphan parts, prune presentation rels whose slide
            part vanished, re-point notesSlide back-references, notes, footer,
            background, slide visibility (re-apply `show="0"`), placeholder
            images, hyperlink rels, doc props, uniquify cNvPr shape ids
verify      self-check the produced bytes against the known repair triggers;
            any finding aborts with E_INTEGRITY, nothing is written
write       tmp file + rename (atomic), re-read → rev.after + ref→id map
```

## Engine constraints (why it is built this way)

- **pptx-automizer imports slides, not layouts.** Hence the *seed deck*: for
  each template, one empty slide per layout with the layout's placeholders
  cloned on (normalized names `PptcPh-<idx>`), generated on demand and cached
  by template content hash. `slide.add layout=N` imports seed slide N+1.
  Footer, slide-number and date placeholders are cloned WITH their layout
  content (footer text, `slidenum` field) so built slides behave like
  PowerPoint-inserted ones; a master-view `lastView` in the template's
  viewProps is stripped so decks open in normal view.
- **`autoImportSlideMasters` stays off.** Templates with OLE objects on
  masters/layouts crash automizer's master import; instead, the deck itself
  always carries its masters/layouts (decks are created from the seed), so
  imported slides re-bind to the root's identical layout numbering.
- **automizer only appends.** Each apply rebuilds the deck: root loaded with
  `removeExistingSlides`, every kept slide re-imported from the same file in
  plan order. Edits on kept slides are modify callbacks on the re-import.
- **The post-pass exists** because automizer cannot express notes, footer
  cloning, backgrounds, slide visibility (it drops the `show="0"` "Hide
  Slide" attribute on every rebuild), images into placeholders, hyperlink
  relationships or doc props. It operates on the written zip via DOM and runs
  before the atomic rename, so the all-or-nothing guarantee covers it.
- **The post-pass also repairs after automizer.** Re-applying renames slide
  parts and leaves relationship debris behind; element merge produces
  colliding shape ids. The post-pass prunes presentation relationships whose
  target part vanished, re-points each notesSlide back-reference at its
  current parent slide, and renumbers duplicate `cNvPr` ids per slide --
  any one of these triggers PowerPoint's "repair" dialog if left in.
- **stdout is shielded during the engine pass.** automizer logs diagnostics
  via `console.log`; the session diverts console output to stderr so the
  one-JSON-envelope contract on stdout holds even on engine complaints.
- **Garbage collection covers slide and notes parts only.** Media of removed
  slides remains as unreferenced parts in the archive -- harmless for
  correctness, slightly larger files after many removals.

## Extending: a new op

1. Add the Zod schema in `schema/ops.ts` and include it in `OpSchema` /
   `OP_NAMES`.
2. Create the plan transformer in `core/ops/` fulfilling `OpHandler`
   (`plan(ctx, op)` rewrites the MutationPlan; never touch files).
3. Register it in `core/ops/planner.ts` (`HANDLERS`).
4. If the op introduces a new *kind* of mutation, teach `engine/session.ts`
   (callback) or `engine/post.ts` (zip level) to interpret it.
5. Add a unit test for the transformer and extend the integration roundtrip.

## Testing strategy

- `test/unit/` -- pure core with geometry fixtures, no files
- `test/golden/` -- `tpl describe` Markdown pinned as snapshot (the LLM-facing
  contract)
- `test/integration/` -- real write path against the committed neutral
  template fixture, incl. all-or-nothing and rev semantics
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
