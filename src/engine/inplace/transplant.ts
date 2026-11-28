/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/inplace/transplant: generated elements (`el.add`). PptxGenJS
**  renders all elements of one apply into a single in-memory deck -- one
**  generator slide per target slide -- and every generated shape is moved
**  into its target slide by its unique object name, together with the parts
**  it references (charts with their workbook, media) under fresh names.
**  The generator output is finished first (engine/elements: theme colors,
**  paragraph properties, chart axes), each moved shape after its move.
*/

import PptxGenJsPkg from "pptxgenjs"
import { randomUUID } from "node:crypto"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../../infra/errors.js"
import type { ChartData, ElementSpec } from "../../schema/payloads.js"
import type { SlidePlanEntry } from "../../core/ops/registry.js"
import { GenColors, addElement, finishChart, finishGenerated, finishShape, type GenRoot, type GenSlide } from "../elements.js"
import { PartAllocator, relTarget, relsPartOf } from "../parts.js"
import { NS_R, elements, firstElement, relRefs } from "../xml.js"
import { Package } from "./package.js"

/**  the PptxGenJS API surface used here  */
interface Generator extends GenRoot {
    addSlide(): GenSlide
    write(opts: { outputType: "nodebuffer" }): Promise<unknown>
}

/**
 *  Unwrap the generator class from its module default export. PptxGenJS 4
 *  ships an `exports` map, so esbuild, vitest and plain node all load its
 *  ESM build and get the class directly; the loop only matters if a loader
 *  falls back to the CommonJS build, whose interop may add `.default` levels.
 *
 *  @param mod - the imported default export
 *  @returns the generator constructor
 */
export const unwrapDefault = (mod: unknown): new () => Generator => {
    let candidate: unknown = mod
    while (typeof candidate !== "function" && candidate !== null && typeof candidate === "object"
        && "default" in (candidate as Record<string, unknown>))
        candidate = (candidate as Record<string, unknown>)["default"]
    return candidate as new () => Generator
}

/**  the PptxGenJS generator class  */
export const PptxGenJS = unwrapDefault(PptxGenJsPkg)

/**  a PptxGenJS slide that stamps one object name on every object it adds  */
const named = (slide: GenSlide, objectName: string): GenSlide => ({
    addText: (text, opts): void => { slide.addText(text, { ...opts, objectName }) },
    addTable: (rows, opts): void => { slide.addTable(rows, { ...opts, objectName }) },
    addChart: (type, data, opts): void => { slide.addChart(type, data, { ...opts, objectName }) },
    addShape: (type, opts): void => { slide.addShape(type, { ...opts, objectName }) },
    addImage: (opts): void => { slide.addImage({ ...opts, objectName }) }
})

/**  the generated elements of one apply, ready to be moved into place  */
export interface Generated {
    /**
     *  Move the generated elements of a plan entry into its slide.
     *
     *  @param pkg - the deck package
     *  @param entry - the plan entry the elements belong to
     *  @param slide - the target slide DOM
     *  @param rels - the target slide's relationships
     *  @param part - the target slide part
     */
    transplant(pkg: Package, entry: SlidePlanEntry, slide: Document, rels: Document, part: string): Promise<void>
}

/**
 *  Finish the generator output before anything is moved: every slide and
 *  chart part gets its theme colors and paragraph properties, every chart
 *  its axis formats.
 *
 *  @param source - the generator package (changed in place)
 *  @param entries - the plan entries with elements
 *  @param placed - generator slide part and object names per entry
 *  @param colors - the color tokens of the apply
 */
const finishSource = async (source: Package, entries: SlidePlanEntry[],
    placed: Map<SlidePlanEntry, { part: string, names: string[] }>, colors: GenColors): Promise<void> => {
    const charts = new Map<string, ChartData>()
    for (const entry of entries) {
        const info = placed.get(entry)
        if (info === undefined || !source.has(relsPartOf(info.part)))
            continue
        const slide = await source.xml(info.part)
        const rels = new Map(source.targets(info.part, await source.rels(info.part)).map((t) => [ t.id, t.target ]))
        entry.elements.forEach((planned, i) => {
            const cNvPr = elements(slide, "p:cNvPr").find((el) => el.getAttribute("name") === info.names[i])
            const rid = cNvPr === undefined ? null : firstElement(cNvPr.parentNode?.parentNode as Element, "c:chart")?.getAttributeNS(NS_R, "id") ?? null
            const target = rid === null ? null : rels.get(rid) ?? null
            if (planned.spec.type === "chart" && target !== null && target !== undefined)
                charts.set(target, planned.spec.data)
        })
    }
    const parts = Object.keys(source.zip.files).filter((f) => /^ppt\/(slides\/slide|charts\/chart)\d+\.xml$/.test(f))
    for (const part of parts) {
        const doc = await source.xml(part)
        const chart = charts.get(part)
        const changed = finishGenerated(doc, colors.tokens)
        if ((chart !== undefined && finishChart(doc, chart)) || changed)
            source.write(part, doc)
    }
}

/**
 *  Render the planned elements of all entries with PptxGenJS. Each element
 *  is named `<name>-<uuid>` (or `<uuid>` when unnamed), so el.set/el.rm keep
 *  addressing it by its name prefix.
 *
 *  @param entries - the plan entries of the apply
 *  @param lang - the language of text that states none (the deck language)
 *  @returns the generated elements, null when no entry has any
 */
export const generateElements = async (entries: SlidePlanEntry[], lang: string): Promise<Generated | null> => {
    const withElements = entries.filter((e) => e.elements.length > 0)
    if (withElements.length === 0)
        return null
    const gen = new PptxGenJS()
    const colors = new GenColors(withElements.flatMap((e) => e.elements.map((planned) => planned.spec)))
    const placed = new Map<SlidePlanEntry, { part: string, names: string[] }>()
    for (const [ k, entry ] of withElements.entries()) {
        const slide = gen.addSlide()
        const names = entry.elements.map((planned) => {
            const name = `${planned.name !== null ? `${planned.name}-` : ""}${randomUUID()}`
            addElement(named(slide, name), gen, planned.spec, lang, colors)
            return name
        })
        placed.set(entry, { part: `ppt/slides/slide${k + 1}.xml`, names })
    }
    const source = await Package.open(await gen.write({ outputType: "nodebuffer" }) as Buffer, [])
    await finishSource(source, withElements, placed, colors)

    return {
        async transplant(pkg, entry, slide, rels, part): Promise<void> {
            const info = placed.get(entry)
            const spTree = firstElement(slide, "p:spTree")
            if (info === undefined || spTree === null)
                return
            const genSlide = await source.xml(info.part)
            const genRels = new Map(source.targets(info.part, await source.rels(info.part)).map((t) => [ t.id, t ]))
            const nextId = PartAllocator.shapeIds(slide)
            const copies = new Map<string, string>()

            /*  generated shapes go in front of the existing ones (later in
                the tree = higher in z-order), but before the tree's own
                extension list  */
            const anchor = elements(spTree, "p:extLst").find((el) => el.parentNode === spTree) ?? null
            for (const name of info.names) {
                const cNvPr = elements(genSlide, "p:cNvPr").find((el) => el.getAttribute("name") === name)
                if (cNvPr === undefined)
                    throw new PptcError("E_ENGINE", `generated element '${name}' not found in the generator output`)
                const shape = slide.importNode(cNvPr.parentNode?.parentNode as Element, true) as Element
                const ids = new Map<string, string>()
                for (const { el, attr } of relRefs(shape)) {
                    let id = ids.get(attr.value)
                    const rel = genRels.get(attr.value)
                    if (id === undefined && rel !== undefined) {
                        const target = rel.target === null ? rel.el.getAttribute("Target") ?? ""
                            : relTarget(part, await pkg.clonePart(source, rel.target, copies))
                        id = pkg.addRel(rels, rel.type, target, rel.target === null)
                        ids.set(attr.value, id)
                    }
                    if (id !== undefined)
                        el.setAttributeNS(NS_R, attr.name, id)
                }
                firstElement(shape, "p:cNvPr")?.setAttribute("id", String(nextId()))
                finishShape(shape, entry.elements[info.names.indexOf(name)]?.spec as ElementSpec)
                spTree.insertBefore(shape, anchor)
            }
        }
    }
}
