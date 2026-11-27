/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/automizer-patches: part-name allocation fixes for pptx-automizer.
**  Every apply re-imports all slides into a package that still holds the
**  parts of the previous apply. automizer numbers new parts from a COUNT
**  (or a slide-local rel id), so a name of a part that is still in use is
**  handed out again and its content overwritten -- another slide then
**  shows the wrong image, chart style or embedded object. These patches
**  seed each counter from the HIGHEST number of every part name the
**  counter is used for, and give OLE embeddings a deck-unique name.
**  Further patches: keep picture bullets (`a:buBlip`); resolve pictures
**  whose media is not named `image*`; stop `package` rels (embedded Excel
**  objects) from passing for charts, SmartArt or layouts; leave linked
**  (external) pictures and OLE objects alone; copy SmartArt without a
**  drawing part; accept XML parts that start with a byte-order mark.
**
**  The patched members are private to automizer: installing fails loudly
**  when one is missing, so an automizer upgrade cannot turn them into a
**  silent no-op (see test/contract/bundle-patches.test.ts).
*/

import { CountHelper } from "pptx-automizer/dist/helper/count-helper.js"
import { OLEObject } from "pptx-automizer/dist/shapes/ole.js"
import { Image } from "pptx-automizer/dist/shapes/image.js"
import { Diagram } from "pptx-automizer/dist/shapes/diagram.js"
import { FileHelper } from "pptx-automizer/dist/helper/file-helper.js"
import ArchiveModule from "pptx-automizer/dist/helper/archive/archive.js"
import { PlaceholderNormalizer } from "pptx-automizer/dist/classes/placeholder-normalizer.js"
import { XmlHelper } from "pptx-automizer/dist/helper/xml-helper.js"
import { XmlRelationshipHelper } from "pptx-automizer/dist/helper/xml-relationship-helper.js"

/**  the subset of automizer's archive API used here  */
interface Archive {
    folder(dir: string): Promise<{ relativePath: string }[]>
}

/**  highest number captured by a pattern over the files of some folders  */
const maxNumber = async (archive: Archive, specs: [string, RegExp][]): Promise<number> => {
    let max = 0
    for (const [ dir, re ] of specs)
        for (const file of await archive.folder(dir))
            max = Math.max(max, Number(re.exec(file.relativePath)?.[1] ?? "0"))
    return max
}

/**  replace a member of an automizer object, failing when it is gone  */
const patch = (target: object, name: string, fn: unknown): void => {
    if (typeof (target as Record<string, unknown>)[name] !== "function")
        throw new Error(`pptx-automizer patch target '${name}' not found -- re-check engine/automizer-patches.ts after an upgrade`)
    ;(target as Record<string, unknown>)[name] = fn
}

/*  images: automizer writes `media<N>.<ext>` but counted `image*` files only  */
patch(CountHelper, "countImages", async (archive: Archive): Promise<number> =>
    await maxNumber(archive, [ [ "ppt/media", /^(?:image|media)(\d+)\./i ] ]))

/*  charts: one number names the chart, its style/colors, drawing, theme
    override, workbook and chart media -- count over all of them  */
patch(CountHelper, "countCharts", async (archive: Archive): Promise<number> =>
    await maxNumber(archive, [
        [ "ppt/charts",     /^(?:chart|chartEx|colors|style)(\d+)\.xml$/i ],
        [ "ppt/drawings",   /^drawing(\d+)\.xml$/i ],
        [ "ppt/theme",      /^themeOverride(\d+)\.xml$/i ],
        [ "ppt/embeddings", /-created-(\d+)\.[^.]+$/i ],
        [ "ppt/media",      /-chart-(\d+)\.[^.]+$/i ]
    ]))

/*  SmartArt: data/colors/layout/quickStyle/drawing share one number  */
patch(CountHelper, "countDiagrams", async (archive: Archive): Promise<number> =>
    await maxNumber(archive, [ [ "ppt/diagrams", /^(?:data|colors|layout|quickStyle|drawing)(\d+)\.xml$/i ] ]))

/*  OLE: automizer names the embedding after the slide-local next rel id
    (`oleObject<rId>.bin`), so objects of two slides overwrite each other.
    The rel id doubles as file token, so make it deck-unique: `rIdOle<n>`.  */
interface OleShape {
    sourceRid: string
    createdRid: string
    sourceArchive: unknown
    targetArchive: Archive
    targetSlideRelFile: string
    setTarget(targetTemplate: unknown, targetSlideNumber: number): Promise<void>
    copyOleObjectFile(sourceFilePath: string): Promise<void>
    appendToContentTypes(): Promise<void>
    updateSlideRels(): Promise<void>
    updateSlideXml(): Promise<void>
}
const oleStatic = OLEObject as unknown as {
    getAllOnSlide(archive: unknown, relsPath: string): Promise<{ rId: string, file: string }[]>
    prototype: object
}
patch(oleStatic.prototype, "prepare", async function (this: OleShape, targetTemplate: unknown, targetSlideNumber: number,
    oleObjects?: { rId: string, file: string }[]): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber)
    const all = oleObjects ?? await oleStatic.getAllOnSlide(this.sourceArchive, this.targetSlideRelFile)
    const ole = all.find((obj) => obj.rId === this.sourceRid)
    if (ole === undefined)
        throw new Error(`OLE object with rId ${this.sourceRid} not found.`)
    const n = 1 + await maxNumber(this.targetArchive, [ [ "ppt/embeddings", /^oleObjectrIdOle(\d+)\./ ] ])
    this.createdRid = `rIdOle${n}`
    await this.copyOleObjectFile(`ppt/embeddings/${ole.file.split("/").pop() as string}`)
    await this.appendToContentTypes()
    await this.updateSlideRels()
    await this.updateSlideXml()
})

/*  picture bullets: automizer strips every `a:buBlip` as "unsupported",
    so bullets of kept slides silently turn into plain bullets. Their image
    rel is copied with the slide anyway -- keep the element.  */
interface Normalizer { host: { unsupportedTags: string[] } }
const removeUnsupportedTags = (PlaceholderNormalizer as unknown as { prototype: Record<string, unknown> })
    .prototype["removeUnsupportedTags"] as (this: Normalizer, xml: unknown) => void
patch((PlaceholderNormalizer as unknown as { prototype: object }).prototype, "removeUnsupportedTags",
    function (this: Normalizer, xml: unknown): void {
        const tags = this.host.unsupportedTags
        this.host.unsupportedTags = tags.filter((t) => t !== "a:buBlip")
        try {
            removeUnsupportedTags.call(this, xml)
        }
        finally {
            this.host.unsupportedTags = tags
        }
    })

/*  images: automizer resolves a picture's rel only when its target starts
    with `../media/image`, but it names copied media `media<N>` itself -- so
    el.rm (or any import) of such a picture crashed with "Cannot read
    properties of undefined (reading 'file')". Fall back to the rel id.  */
interface RelItem { rId: string, type: string, file: string, filename: string, element: unknown }
interface XmlHelperStatic {
    getTargetByRelId(archive: unknown, relsPath: string, element: unknown, type: string): Promise<unknown>
    getRelationshipItems(archive: unknown, relsPath: string,
        cb: (element: { getAttribute(name: string): string | null }, rels: RelItem[]) => void): Promise<RelItem[]>
}
const xmlHelper = XmlHelper as unknown as XmlHelperStatic
const getTargetByRelId = xmlHelper.getTargetByRelId.bind(xmlHelper)
const IMAGE_BLIPS: Record<string, string> = { "image": "a:blip", "image:svg": "asvg:svgBlip" }
patch(xmlHelper, "getTargetByRelId", async (archive: unknown, relsPath: string, element: unknown, type: string): Promise<unknown> => {
    const target = await getTargetByRelId(archive, relsPath, element, type)
    const tag = IMAGE_BLIPS[type]
    if (target !== undefined || tag === undefined)
        return target
    const blip = (element as { getElementsByTagName(n: string): { item(i: number): { getAttribute(n: string): string | null } | null } })
        .getElementsByTagName(tag).item(0)
    const rid = blip?.getAttribute("r:embed") ?? null
    if (rid === null)
        return undefined
    const rels = await xmlHelper.getRelationshipItems(archive, relsPath, (rel, items) => {
        const file = rel.getAttribute("Target") ?? ""
        items.push({ rId: rel.getAttribute("Id") ?? "", type: rel.getAttribute("Type") ?? "",
            file, filename: file.split("/").pop() ?? "", element: rel })
    })
    return rels.find((rel) => rel.rId === rid
        && (rel.element as { getAttribute(n: string): string | null }).getAttribute("TargetMode") !== "External")
})

/*  embedded files: automizer's prefix match accepts every `package` rel
    for ANY prefix (meant for chart workbooks), so an embedded Excel object
    on a slide (Insert > Object) was imported as "chart0", "data0", ... and
    every apply failed with "Could not find file ppt/charts/chart0.xml".
    A package rel matches only an embeddings prefix.  */
interface MatchHelper { targetMatchesRelationship(relType: string, subtype: string, file: string, prefix: string): boolean }
const matchHelper = XmlHelper as unknown as MatchHelper
const targetMatches = matchHelper.targetMatchesRelationship.bind(matchHelper)
patch(matchHelper, "targetMatchesRelationship", (relType: string, subtype: string, file: string, prefix: string): boolean =>
    relType === "package" && !prefix.includes("embeddings") ? false : targetMatches(relType, subtype, file, prefix))

/*  the same flaw in the second matcher, which also resolves a slide's
    layout: a package rel listed first would pass for the layout  */
const relHelper = XmlRelationshipHelper as unknown as MatchHelper
const relTargetMatches = relHelper.targetMatchesRelationship.bind(relHelper)
patch(relHelper, "targetMatchesRelationship", (relType: string, subtype: string, file: string, prefix: string): boolean =>
    relType === "package" && !prefix.includes("embeddings") ? false : relTargetMatches(relType, subtype, file, prefix))

/*  linked files: a picture or OLE object linked to an external file
    (TargetMode="External", e.g. file:///C:/logo.png) has nothing to copy,
    but automizer tried to and failed with "Could not find file
    ppt/media/file:///...". Its rel is carried over verbatim with the
    slide, so leave linked targets alone.  */
type RelTarget = { element?: { getAttribute(name: string): string | null } }
const isInternal = (rel: RelTarget): boolean => rel.element?.getAttribute("TargetMode") !== "External"
for (const cls of [ Image, OLEObject ] as unknown as { getAllOnSlide(archive: unknown, relsPath: string): Promise<RelTarget[]> }[]) {
    const getAll = cls.getAllOnSlide.bind(cls)
    patch(cls, "getAllOnSlide", async (archive: unknown, relsPath: string): Promise<RelTarget[]> =>
        (await getAll(archive, relsPath)).filter(isInternal))
}

/*  SmartArt without a drawing part (older PowerPoint, foreign tools): the
    drawing is only a cached rendering, but automizer copied it blindly and
    failed with "Could not find file ppt/diagrams/drawingN.xml". Copy the
    parts that exist; the post-pass drops the then dead drawing rel.  */
interface DiagramShape {
    sourceNumber: number
    targetNumber: number
    sourceArchive: { fileExists(file: string): boolean }
    targetArchive: unknown
}
const fileHelper = FileHelper as unknown as {
    zipCopy(source: unknown, sourceFile: string, target: unknown, targetFile: string): Promise<void>
}
patch((Diagram as unknown as { prototype: object }).prototype, "copyFiles", async function (this: DiagramShape): Promise<void> {
    for (const kind of [ "data", "colors", "drawing", "layout", "quickStyle" ]) {
        const from = `ppt/diagrams/${kind}${this.sourceNumber}.xml`
        if (kind !== "drawing" || this.sourceArchive.fileExists(from))
            await fileHelper.zipCopy(this.sourceArchive, from, this.targetArchive, `ppt/diagrams/${kind}${this.targetNumber}.xml`)
    }
})

/*  byte-order marks: some .NET tools and 6 of PowerPoint's own themes
    start XML parts with a BOM, which xmldom rejects ("processing
    instruction ... is an xml declaration which is only at the start of the
    document") -- every apply on such a deck failed. Drop it before parsing.  */
const archiveClass = ((ArchiveModule as unknown as { default?: unknown }).default ?? ArchiveModule) as { prototype: object }
const parseArchiveXml = (archiveClass.prototype as { parseXml(xml: string): unknown }).parseXml
patch(archiveClass.prototype, "parseXml", function (this: unknown, xml: string): unknown {
    return parseArchiveXml.call(this, xml.replace(/^\uFEFF/, ""))
})
