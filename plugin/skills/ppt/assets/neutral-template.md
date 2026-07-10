**Template notes (bundled fallback):**

This is the bundled fallback template -- a neutral Office-style theme
(16:9). It is used when the user names no template.
Always TELL the user the neutral default is in use and that a corporate
`.potx`/`.pptx` can be supplied instead at any time.

All layouts (address by NAME, indices may shift):

| # | Layout name                | Placeholders            | Role / usage                                        |
|---|----------------------------|-------------------------|-----------------------------------------------------|
| 0 | `TITLE_SLIDE`              | title, subtitle         | Title (cover)                                       |
| 1 | `CONTENT`                  | title, body             | Content; also Agenda (chapters as bullets)          |
| 2 | `Section Header`           | title, body             | Chapter (section divider)                           |
| 3 | `TWO_COLUMN`               | title, 2x body          | Comparison / two columns                            |
| 4 | `Comparison`               | title, 2x heading+body  | Labeled two-column comparison                       |
| 5 | `Title Only`               | title                   | Headline over free space                            |
| 6 | `DEFAULT`                  | (none)                  | Blank -- target for `el.add` (tables/charts/shapes) |
| 7 | `Content with Caption`     | title, body, caption    | Content with explanatory side text                  |
| 8 | `PICTURE`                  | title, picture, caption | Picture slide (image prompts go here)               |
| 9 | `Title and Vertical Text`  | title, vertical body    | Vertical text (East-Asian); rarely used             |
| 10| `Vertical Title and Text`  | vertical title, body    | Vertical text (East-Asian); rarely used             |

Footer pattern (set via `footer` on every slide):

- `<presentation title> | <current year>`
- On slides with picture placeholders append the AI-image note in the
  deck language (wording per `references/content-rules.md`).
- All layouts carry footer and slide-number placeholders.

Image design constraints:

- Primary accent is `accent1` (#4F81BD, muted blue); no further
  template-specific constraints -- standard prompt rules apply.
- No chapter image: `Section Header` has no picture placeholder; image
  prompts only on `PICTURE` slides.
