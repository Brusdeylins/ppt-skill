Image Prompt Formula (Nano Banana Pro)
======================================

Every image prompt follows this formula (English, one paragraph):

**[Style] + [Subject with primary accent] + [Scene / text-in-image] +
[Lighting / background constraint] + [Quality]**

The aspect ratio is NOT part of the prompt text -- pptc derives it from
the picture placeholder's geometry and prints it in the prompt-box header
(`IMAGE PROMPT · <ratio>`).

Element rules
-------------

-   **Style**: the deck-wide style block from `style-catalog.md`,
    verbatim. Photos/motifs use the image style; diagrams use the
    info-graphic style.
-   **Subject with primary accent**: the main subject carries ONE
    visible element in the template's primary accent color, named with
    hue and hex code, e.g. `a deep blue (#1F4E79) blazer`.
-   **Gaze faces the slide centre**: when the image shows people or
    faces, they must look/face TOWARD the slide's content, never off the
    outer edge. Derive the direction from the placeholder geometry: a
    picture placeholder on the LEFT half -> subjects face/look RIGHT
    (toward centre); on the RIGHT half -> they face/look LEFT. Put it in
    the prompt, e.g. `the person is turned toward the left, looking into
    frame` for a right-side placeholder.
-   **Scene / text-in-image**: where the scene plausibly carries text
    (screens, banners, whiteboards, signs) -- and ALWAYS on title-role
    images -- embed the slide title or deck topic:
    `text on screen reads "<title>"`. Keep embedded text short
    (2-5 words); the model renders short text reliably. This is the
    ONLY sanctioned use of text vocabulary -- and it excludes the
    `No text` suffix on that prompt.
-   **Lighting / background constraint**: lighting matching the style
    (e.g. `soft diffused natural light`); apply template sidecar
    constraints here (e.g. `dark muted background` for title/closing
    layouts with a white line, when the style is illustration/render).
-   **Overlay regions become negative space**: `tpl inspect`/`describe`
    report per picture placeholder which shapes sit on top and WHERE
    (`overlays`: e.g. `Title (bottom area, left part)`). Translate every
    overlay into Google's documented negative-space pattern:
    `the [bottom-left area] is a vast empty [color] canvas creating
    significant negative space` -- and position the subject in the
    free region (`positioned in the upper right third`).
    CRITICAL: never EXPLAIN the reason. Typographic vocabulary
    ("title", "footer", "caption", "label", "watermark", "text sits
    there") triggers the model's text rendering and produces pseudo
    lettering. Describe only what the IMAGE looks like.
-   **Text suppression**: unless the prompt deliberately embeds text
    (quoted, short), end with the official suffix clause:
    `No text. No letters. No symbols.` Positive phrasing beats
    negation everywhere else ("plain matte surface", not "no clutter").
-   **Quality**: `sharp focus, 8k` plus style-appropriate terms.
-   **Aspect**: never write the ratio into the prompt text. pptc computes
    it from the placeholder geometry and surfaces it in the box header, so
    the image model is not nudged by a ratio phrase inside the prompt.
-   **Secondary colors**: only for graphic elements within the image
    (dashboards, charts, UI) -- usage rules and hex sources per
    `color-roles.md`.

Per-image creative step
-----------------------

Never template the motif. For each picture placeholder, decide
deliberately:

1.  What is THIS slide's POINT -- its message AND its takeaway / conclusion
    (the conclusion it lands), not just its topic or the problem it states?
2.  What is the placeholder's role (title / chapter / content /
    background / contact)?
3.  Which single motif makes that POINT tangible -- leaning toward the
    resolution the slide concludes with, not only the status quo -- without
    repeating a motif already used in this deck? A slide whose conclusion is
    "capture and reuse" should show capture/reuse, not just the repetition;
    one whose point is "share it" should hint at sharing, not only the
    locked-away problem.

A chapter / divider image synthesizes the WHOLE chapter: make the
chapter's core point tangible -- derived from all its slides and their
conclusions together -- not just the literal chapter title.

Contact placeholders get neutral business portraits.

Background image vs. negative space
-----------------------------------

`tpl inspect`/`describe` report a `coverage` per picture placeholder --
the share of the image the overlay text sits on. It splits two modes:

-   **Partial overlay (coverage < 65%)** -- negative-space mode (above):
    keep the overlaid regions calm and move the subject into the free
    area. The image still has a real subject.
-   **True background image (coverage >= 65%)** -- the text sits on the
    WHOLE image, so it is a backdrop, not a subject:
    -   **No text in the image, ever.** This OVERRIDES the title-text
        rule -- a backdrop carries no embedded words (they collide with
        the real text on top). Always end with `No text. No letters. No
        symbols.`
    -   **One even tone, no clashing luminance.** Pick a single tone and
        hold it across the frame, phrased positively:
        -   a DARK backdrop (deep, low-key, `dk1`/`dk2`) with NO bright
            hotspots -- e.g. `evenly lit deep charcoal (#1F1F1F) field,
            soft low-key gradient, no bright highlights` -- paired with
            LIGHT overlay text; OR
        -   a LIGHT backdrop (high-key, `lt1`/`lt2`) with NO dark blocks
            -- e.g. `clean high-key off-white (#F2F2F2) surface, soft even
            light` -- paired with DARK overlay text.
    -   **Accent stays subtle** -- the primary accent appears only as a
        faint, low-contrast element, or not at all; the tone dominates so
        the text on top stays readable.
    -   **Contrast the overlay text** -- the SKILL sets the placeholder
        text colour to contrast the chosen tone (light on dark, dark on
        light) so the words stay legible (see STEP 7).

Example (title role, 2:3 placeholder, stock photo, accent1 = #1F4E79)
--------------------------------------------------------------------

The 2:3 ratio is shown in the box header by pptc, not in the prompt:

```
commercial photography, shot on Canon EOS R5 85mm f/1.4, professional
business woman in a deep blue (#1F4E79) blazer standing in a modern
glass office lobby, holding a tablet whose screen reads "AI Strategy",
soft diffused natural light from floor-to-ceiling windows, shallow
depth of field, professional color grading, sharp focus, 8k
```
