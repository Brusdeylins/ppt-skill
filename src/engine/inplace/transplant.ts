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
*/

import PptxGenJsPkg from "pptxgenjs"
import { randomUUID } from "node:crypto"
import path from "node:path"
import type { Document, Element } from "@xmldom/xmldom"
import { PptcError } from "../../infra/errors.js"
import type { SlidePlanEntry } from "../../core/ops/registry.js"
import { addElement, type GenRoot, type GenSlide } from "../elements.js"
import { PartAllocator } from "../parts.js"
import { NS_R, elements, firstElement } from "../xml.js"
import { Package } from "./package.js"

/**  the PptxGenJS API surface used here  */
interface Generator extends GenRoot {
    addSlide(): GenSlide
    write(opts: { outputType: "nodebuffer" }): Promise<unknown>
}

/**  unwrap the CJS default export: plain node, esbuild and vitest disagree
     on how many .default levels the interop introduces  */
const unwrapDefault = (mod: unknown): new () => Generator => {
    let candidate: unknown = mod
    while (typeof candidate !== "function" && candidate !== null && typeof candidate === "object"
        && "default" in (candidate as Record<string, unknown>))
        candidate = (candidate as Record<string, unknown>)["default"]
    return candidate as new () => Generator
}
const PptxGenJS = unwrapDefault(PptxGenJsPkg)

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
 *  Render the planned elements of all entries with PptxGenJS. Each element
 *  is named `<name>-<uuid>` (or `<uuid>` when unnamed), so el.set/el.rm keep
 *  addressing it by its name prefix.
 *
 *  @param entries - the plan entries of the apply
 *  @returns the generated elements, null when no entry has any
 */
export const generateElements = async (entries: SlidePlanEntry[]): Promise<Generated | null> => {
    const withElements = entries.filter((e) => e.elements.length > 0)
    if (withElements.length === 0)
        return null
    const gen = new PptxGenJS()
    const placed = new Map<SlidePlanEntry, { part: string, names: string[] }>()
    for (const [ k, entry ] of withElements.entries()) {
        const slide = gen.addSlide()
        const names = entry.elements.map((planned) => {
            const name = `${planned.name !== null ? `${planned.name}-` : ""}${randomUUID()}`
            addElement(named(slide, name), gen, planned.spec)
            return name
        })
        placed.set(entry, { part: `ppt/slides/slide${k + 1}.xml`, names })
    }
    const source = await Package.open(await gen.write({ outputType: "nodebuffer" }) as Buffer, [])

    return {
        async transplant(pkg, entry, slide, rels, part): Promise<void> {
            const info = placed.get(entry)
            const spTree = firstElement(slide, "p:spTree")
            if (info === undefined || spTree === null)
                return
            const genSlide = await source.xml(info.part)
            const genRels = new Map(Package.targets(info.part, await source.rels(info.part))
                .map((t) => [ t.rel.getAttribute("Id") ?? "", t ]))
            const nextId = PartAllocator.shapeIds(slide)
            const copies = new Map<string, string>()

            /*  generated shapes go behind the existing ones, before the
                tree's own extension list  */
            const anchor = elements(spTree, "p:extLst").find((el) => el.parentNode === spTree) ?? null
            for (const name of info.names) {
                const cNvPr = elements(genSlide, "p:cNvPr").find((el) => el.getAttribute("name") === name)
                if (cNvPr === undefined)
                    throw new PptcError("E_ENGINE", `generated element '${name}' not found in the generator output`)
                const shape = slide.importNode(cNvPr.parentNode?.parentNode as Element, true) as Element
                const ids = new Map<string, string>()
                const relink = async (el: Element): Promise<void> => {
                    for (let i = 0; i < el.attributes.length; i++) {
                        const attr = el.attributes.item(i)
                        if (attr === null || attr.namespaceURI !== NS_R)
                            continue
                        let id = ids.get(attr.value)
                        const rel = genRels.get(attr.value)
                        if (id === undefined && rel !== undefined) {
                            const target = rel.target === null ? rel.rel.getAttribute("Target") ?? ""
                                : path.posix.relative(path.posix.dirname(part), await pkg.clonePart(source, rel.target, copies))
                            id = pkg.addRel(rels, rel.type, target, rel.target === null)
                            ids.set(attr.value, id)
                        }
                        if (id !== undefined)
                            el.setAttributeNS(NS_R, attr.name, id)
                    }
                    for (let n = el.firstChild; n !== null; n = n.nextSibling)
                        if (n.nodeType === 1)
                            await relink(n as Element)
                }
                await relink(shape)
                firstElement(shape, "p:cNvPr")?.setAttribute("id", String(nextId()))
                spTree.insertBefore(shape, anchor)
            }
        }
    }
}
