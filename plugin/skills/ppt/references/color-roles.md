Color Roles
===========

Template-neutral mapping from OOXML `clrScheme` slots to prompt roles.
This file contains NO concrete colors -- the actual hex values come
from the current template at runtime via `pptc tpl inspect` →
`result.colors` (RRGGBB without `#`; always prefix `#` in prompts).

Every OOXML theme defines exactly these slots, so this mapping works
with any `.potx`/`.pptx` template:

| Slot               | Prompt Role                                                                                                         |
|--------------------|---------------------------------------------------------------------------------------------------------------------|
| `accent1`          | PRIMARY brand accent: the dominant, visible color element in every image (clothing, lighting, branding, key object) |
| `accent2`-`accent6`| Secondary palette: diagram elements, graphic details, subtle accents -- never dominant                              |
| `dk1`, `dk2`       | Dark tones: text-on-image, dark backgrounds, shadows                                                                |
| `lt1`, `lt2`       | Light/neutral tones: backgrounds, surfaces, whitespace                                                              |
| `hlink`, `folHlink`| Not used in image prompts                                                                                           |

Usage rules
-----------

-   Exactly ONE primary accent element per image; secondary colors
    only for graphics/diagrams (info-graphics may use 2-3 of them).
-   Always write colors as `#RRGGBB` hex codes into the prompt, e.g.
    `deep blue (#1F4E79)` -- name the hue AND give the code.
-   In ops (element fills, borders, fonts, table and chart colors, runs,
    backgrounds) give the slot NAME instead (`"accent1"`, `"lt1"`): pptc
    writes it as a theme reference, so the element follows the template.
    Hex codes belong into image prompts, which cannot reference a theme.
-   Primary slot default is `accent1`. If a template uses a different
    slot as its brand color, the template sidecar (or the user) may
    override the primary slot; everything else stays the same.
-   On a true BACKGROUND image (high `coverage`), a `dk1`/`dk2` (dark) or
    `lt1`/`lt2` (light) tone DOMINATES the whole frame and the accent
    stays faint or absent. Pair a dark backdrop with light overlay text,
    a light backdrop with dark overlay text -- see
    `prompt-formula.md` → "Background image vs. negative space".
