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
**  counter is used for, and give OLE embeddings a deck-unique name. A last
**  patch keeps picture bullets (`a:buBlip`) that automizer strips.
**
**  The patched members are private to automizer: installing fails loudly
**  when one is missing, so an automizer upgrade cannot turn them into a
**  silent no-op (see test/contract/bundle-patches.test.ts).
*/

import { CountHelper } from "pptx-automizer/dist/helper/count-helper.js"
import { OLEObject } from "pptx-automizer/dist/shapes/ole.js"
import { PlaceholderNormalizer } from "pptx-automizer/dist/classes/placeholder-normalizer.js"

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
