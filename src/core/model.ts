/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/model: the engine-free domain model of pptc. These types describe
**  decks, slides, placeholders and layouts as plain data. No OOXML, no
**  automizer types ever appear above this layer.
*/

/**  Geometry of a box on a slide, in inches (origin: top-left of the slide).  */
export interface Frame {
    /**  distance from the left slide edge  */
    x: number
    /**  distance from the top slide edge  */
    y: number
    /**  box width  */
    w: number
    /**  box height  */
    h: number
}

/**  Semantic kind of a placeholder, normalized from the OOXML `type` attribute.  */
export type PlaceholderKind =
    | "title" | "subtitle" | "body" | "picture" | "footer" | "slideNumber" | "date" | "other"

/**  placeholder kinds of the footer band -- their small text is by design
     (footer scale), so capacity/font lints and content fills skip them  */
export const isFooterKind = (kind: PlaceholderKind): boolean =>
    kind === "footer" || kind === "slideNumber" || kind === "date"

/**  A placeholder as defined on a slide layout (or instantiated on a slide).  */
export interface Placeholder {
    /**  OOXML placeholder index (`idx` attribute, 0 for the title)  */
    idx: number
    /**  normalized placeholder kind  */
    kind: PlaceholderKind
    /**  shape name as shown in PowerPoint's selection pane  */
    name: string
    /**  resolved geometry in inches, null if fully inherited and unresolvable  */
    frame: Frame | null
    /**  estimated text capacity, null for non-text placeholders  */
    capacity: TextCapacity | null
    /**  text shapes overlapping this PICTURE placeholder: image prompts
         must keep these regions calm (text sits on top)  */
    overlays?: { name: string, region: string }[]
    /**  fraction (0..1) of this PICTURE placeholder's area covered by the
         overlapping text shapes (union, overlap-safe). High coverage
         (about 0.65 or more) marks a true BACKGROUND image: the prompt carries
         no text and keeps one even tone so the overlay text stays legible.
         Undefined for non-picture placeholders or when nothing overlaps.  */
    coverage?: number
}

/**  Estimated text capacity of a text placeholder.  */
export interface TextCapacity {
    /**  approximate number of text lines that fit the box  */
    lines: number
    /**  approximate number of characters per line  */
    charsPerLine: number
    /**  font size in points the estimate is based on  */
    fontSizePt: number
}

/**  One slide layout of a template, fully resolved for description and lint.  */
export interface Layout {
    /**  zero-based layout index in slide-master order (stable address)  */
    index: number
    /**  layout name as defined in the template  */
    name: string
    /**  placeholders in document order, footer/slide-number/date excluded  */
    placeholders: Placeholder[]
    /**  frames of footer/slide-number/date placeholders (overlap lint)  */
    reserved: Frame[]
}

/**  Template-wide data extracted from a .potx/.pptx file.  */
export interface TemplateInfo {
    /**  slide size in inches  */
    slideSize: { w: number, h: number }
    /**  major (heading) and minor (body) theme font names  */
    fonts: { major: string, minor: string }
    /**  theme color map: scheme name to RRGGBB hex  */
    colors: Record<string, string>
    /**  all layouts in slide-master order  */
    layouts: Layout[]
    /**  the master's drawing guides in inches: horizontal guide Y-positions
         and vertical guide X-positions (sorted). Absent when the template
         defines none. Use them to place `el.add` elements on the template's
         own grid instead of guessing coordinates.  */
    guides?: { horizontal: number[], vertical: number[] }
    /**  the clean content rectangle (inches): the largest body placeholder
         snapped to the nearest guides -- the target frame for an `el.add`
         table/textbox/diagram on a title-only layout. Absent when neither a
         body placeholder nor guides are available.  */
    contentArea?: Frame
}

/**  A shape found on an existing slide (read model).  */
/**  One styled text run (read model). Mirrors the run vocabulary of a
     slide.fill / el.add `runs[]` entry, so a run read here passes straight
     back as a write input.  */
export interface RunInfo {
    /**  run text  */
    text: string
    /**  typeface (a:latin) when set on the run; pass as a run `font`  */
    font?: string
    /**  font size in points when set on the run; pass as a run `size`  */
    size?: number
    /**  bold when set on the run  */
    bold?: boolean
    /**  italic when set on the run  */
    italic?: boolean
    /**  font color RRGGBB (theme colors resolved) when set on the run  */
    color?: string
}

/**  One paragraph of a text body (read model), its runs in order. Mirrors a
     slide.fill / el.add paragraph object.  */
export interface ParaInfo {
    /**  the runs of this paragraph in order (empty for a blank line)  */
    runs: RunInfo[]
    /**  indent level (a:pPr lvl) when greater than 0  */
    level?: number
    /**  true/false only when the paragraph sets a bullet explicitly
         (a:buChar / a:buAutoNum / a:buNone); omitted when inherited  */
    bullet?: boolean
}

/**  One shape on a slide as the reader exposes it: identity, classification,
     geometry and -- depending on the shape kind -- text, table, style or
     media facts. The per-slide unit of the `state` read model.  */
export interface ShapeInfo {
    /**  shape name (selection pane)  */
    name: string
    /**  OOXML shape id within the slide  */
    id: number
    /**  shape classification  */
    type: "placeholder" | "textbox" | "picture" | "table" | "chart" | "shape" | "connector" | "group" | "other"
    /**  placeholder idx if the shape is a placeholder, otherwise null  */
    placeholderIdx: number | null
    /**  placeholder kind if the shape is a placeholder, otherwise null  */
    placeholderKind: PlaceholderKind | null
    /**  geometry in inches, null when inherited from the layout  */
    frame: Frame | null
    /**  plain text content, paragraphs joined by newlines, null if none  */
    text: string | null
    /**  table cell texts (rows x cols) when the shape is a table  */
    table?: string[][]
    /**  column widths in inches for a table shape (read-only insight: el.add
         tables auto-distribute columns, this cannot be passed back)  */
    colWidths?: number[]
    /**  preset autoshape geometry (rect, roundRect, diamond, ...) when the
         shape carries an `a:prstGeom`; pass back as el.add `shape` to recreate  */
    shape?: string
    /**  solid fill color RRGGBB (theme colors resolved) when present; mirrors
         el.add `fill`  */
    fill?: string
    /**  outline color RRGGBB (theme colors resolved) when present; mirrors
         el.add `border`  */
    border?: string
    /**  outline width in points when present; mirrors el.add `borderPt`  */
    borderPt?: number
    /**  font size in points of the first text run when present; mirrors
         el.add `fontSize`  */
    fontSize?: number
    /**  font color RRGGBB of the first text run when present; mirrors
         el.add `fontColor`  */
    fontColor?: string
    /**  typeface of the first text run when present; pass it as the `font` of
         a slide.fill / el run to preserve it (e.g. a monospace code block)  */
    fontFace?: string
    /**  media file name under ppt/media for a picture shape, e.g.
         "image11.png", resolved from the blip r:embed relationship; absent
         for non-picture shapes and for placeholders with no image filled  */
    image?: string
    /**  rich text as paragraphs and runs. Present at the 'full' read level
         only when the body carries explicit run formatting (a per-run font,
         size, bold or italic) -- e.g. a monospace code block or an emphasized
         word. The structure round-trips into slide.fill / el run inputs, so a
         styled body can be recreated without reading raw XML. Omitted for
         plain uniform text, where the flat `text` already suffices.  */
    paragraphs?: ParaInfo[]
}

/**  One slide of an existing deck (read model).  */
export interface SlideInfo {
    /**  stable OOXML slide id from `p:sldIdLst` (canonical address)  */
    id: number
    /**  zero-based position in the deck  */
    index: number
    /**  slide title (text of placeholder idx 0), null if empty  */
    title: string | null
    /**  name of the layout this slide uses  */
    layoutName: string
    /**  zero-based layout index in the deck's master, -1 if unresolvable  */
    layoutIndex: number
    /**  all shapes on the slide  */
    shapes: ShapeInfo[]
    /**  speaker notes text, null if none  */
    notes: string | null
    /**  true when the slide is hidden ("Hide Slide" in PowerPoint, the
         `show="0"` attribute on the `<p:sld>` root); preserved across edits  */
    hidden: boolean
    /**  archive part path of the slide (engine address, e.g. "ppt/slides/slide3.xml")  */
    part: string
}

/**  Complete read model of a deck, the agent's source of truth.  */
export interface DeckState {
    /**  absolute path of the deck file  */
    file: string
    /**  revision token (content hash) for optimistic locking  */
    rev: string
    /**  slide size in inches  */
    slideSize: { w: number, h: number }
    /**  all slides in presentation order  */
    slides: SlideInfo[]
    /**  custom document properties (name→value) carried inside the .pptx
         (docProps/custom.xml) -- the deck's self-describing memory, e.g. the
         image/info-graphic style the skill chose. Empty when none are set.  */
    customProps: Record<string, string>
}

/**
 *  Deterministic shape name of a placeholder on a seed slide. Seed slides are
 *  pptc's own artifact, so placeholder names are normalized -- fills address
 *  them independently of how the template names its shapes.
 *
 *  @param idx - OOXML placeholder idx
 *  @returns the normalized shape name
 */
export const seedPlaceholderName = (idx: number): string => `PptcPh-${idx}`

/**  EMU (English Metric Units) per inch, the OOXML length base unit.  */
export const EMU_PER_INCH = 914400

/**
 *  Convert EMU to inches, rounded to two decimals for stable output.
 *
 *  @param emu - length in English Metric Units
 *  @returns length in inches
 */
export const emuToInch = (emu: number): number =>
    Math.round((emu / EMU_PER_INCH) * 100) / 100

/**
 *  Convert inches to EMU.
 *
 *  @param inch - length in inches
 *  @returns length in English Metric Units
 */
export const inchToEmu = (inch: number): number =>
    Math.round(inch * EMU_PER_INCH)
