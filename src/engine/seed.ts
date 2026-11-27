/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/seed: slides and decks from layouts. `slideFromLayout` builds the
**  empty slide `slide.add` puts into a deck (placeholders cloned from the
**  layout, named `PptcPh-<idx>`); `buildEmptyDeck` derives the zero-slide
**  deck `pptc new` starts from a template.
*/

import JSZip from "jszip"
import { readFileSync } from "node:fs"
import type { Element } from "@xmldom/xmldom"
import { PptcError } from "../infra/errors.js"
import { seedPlaceholderName } from "../core/model.js"
import { cleanContentTypes, gcAssets, layoutOrder, normalizeRelTargets } from "./parts.js"
import { NS_A, NS_P, elements, firstElement, parseXml, serializeElement, serializeXml } from "./xml.js"

/**  clone the fillable placeholders of a layout onto an empty slide (DOM)  */
export const slideFromLayout = (layoutXml: string): string => {
    const layout = parseXml(layoutXml)
    const sps: string[] = []
    let nextId = 2
    for (const sp of elements(layout, "p:sp")) {
        const ph = firstElement(sp, "p:ph")
        if (ph === null)
            continue
        const type = ph.getAttribute("type") ?? "body"
        const clone = sp.cloneNode(true) as Element
        const cNvPr = firstElement(clone, "p:cNvPr")
        cNvPr?.setAttribute("id", String(nextId++))
        if (type === "ftr" || type === "sldNum" || type === "dt") {
            /*  header/footer placeholders: keep the layout content (footer
                text, slidenum/datetime fields) so slides show footer and
                page number exactly like PowerPoint-inserted slides  */
            sps.push(serializeElement(clone))
            continue
        }
        cNvPr?.setAttribute("name", seedPlaceholderName(Number(ph.getAttribute("idx") ?? "0")))
        const txBody = firstElement(clone, "p:txBody")
        const doc = clone.ownerDocument
        if (txBody !== null && doc !== null) {
            while (txBody.firstChild !== null)
                txBody.removeChild(txBody.firstChild)
            txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"))
            txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"))
            const p = doc.createElementNS(NS_A, "a:p")
            p.appendChild(doc.createElementNS(NS_A, "a:endParaRPr"))
            txBody.appendChild(p)
        }
        sps.push(serializeElement(clone))
    }
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
        + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        + " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
        + "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        + "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
        + "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>"
        + sps.join("")
        + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
}

/**
 *  Create an empty deck from a template: .potx content types converted,
 *  every slide, notes part, section and master-view setting of the template
 *  dropped. Used by `pptc new`.
 *
 *  @param template - path or bytes of the .potx/.pptx template
 *  @returns bytes of a valid, zero-slide .pptx carrying the template's design
 *  @throws PptcError E_TEMPLATE when the template lacks required parts or layouts
 */
export const buildEmptyDeck = async (template: string | Buffer): Promise<Buffer> => {
    const zip = await JSZip.loadAsync(typeof template === "string" ? readFileSync(template) : template)
    await normalizeRelTargets(zip)
    const get = async (part: string): Promise<string> => {
        const f = zip.file(part)
        if (f === null)
            throw new PptcError("E_TEMPLATE", `template is missing required part: ${part}`)
        return await f.async("string")
    }
    if ((await layoutOrder(zip)).length === 0)
        throw new PptcError("E_TEMPLATE", "template contains no slide layouts")

    /*  drop existing slides and notes completely  */
    const pres = parseXml(await get("ppt/presentation.xml"))
    const root = pres.documentElement as Element
    let sldIdLst = firstElement(pres, "p:sldIdLst")
    if (sldIdLst === null) {
        sldIdLst = pres.createElementNS(NS_P, "p:sldIdLst")
        const anchor = firstElement(pres, "p:sldSz")
        root.insertBefore(sldIdLst, anchor)
    }
    while (sldIdLst.firstChild !== null)
        sldIdLst.removeChild(sldIdLst.firstChild)

    /*  drop PowerPoint sections: they group the template's example slides,
        stale section definitions would reference removed sldIds  */
    for (const ext of elements(pres, "p:ext"))
        if (ext.getAttribute("uri") === "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}") {
            const extLst = ext.parentNode as Element
            extLst.removeChild(ext)
            if (firstElement(extLst, "p:ext") === null)
                extLst.parentNode?.removeChild(extLst)
        }

    /*  CT_Presentation is an ordered sequence: notesMasterIdLst (and
        handoutMasterIdLst) MUST precede sldIdLst -- sloppy templates
        violate this and PowerPoint answers with the repair dialog  */
    for (const lst of [ "p:notesMasterIdLst", "p:handoutMasterIdLst" ]) {
        const el = firstElement(pres, lst)
        if (el !== null && el.parentNode === root) {
            let ahead = false
            for (let n = sldIdLst.previousSibling; n !== null; n = n.previousSibling)
                ahead ||= n === el
            if (!ahead)
                root.insertBefore(el, sldIdLst)
        }
    }
    zip.file("ppt/presentation.xml", serializeXml(pres))

    const presRels = parseXml(await get("ppt/_rels/presentation.xml.rels"))
    for (const rel of elements(presRels, "Relationship"))
        if ((rel.getAttribute("Type") ?? "").endsWith("/slide"))
            rel.parentNode?.removeChild(rel)
    zip.file("ppt/_rels/presentation.xml.rels", serializeXml(presRels))
    for (const f of Object.keys(zip.files))
        if (/^ppt\/(slides|notesSlides)\//.test(f))
            zip.remove(f)
    zip.file("[Content_Types].xml", (await get("[Content_Types].xml"))
        .replace("presentationml.template.main+xml", "presentationml.presentation.main+xml"))

    /*  templates saved in master/layout view carry that view along --
        drop lastView so the produced deck opens in normal view  */
    const viewPr = zip.file("ppt/viewProps.xml")
    if (viewPr !== null) {
        const view = parseXml(await viewPr.async("string"))
        view.documentElement?.removeAttribute("lastView")
        zip.file("ppt/viewProps.xml", serializeXml(view))
    }

    /*  templates may carry stale or duplicate content-type overrides of
        their own, a deck used as template the media of its slides: sweep
        them so every derived deck starts clean  */
    await gcAssets(zip)
    await cleanContentTypes(zip)
    return await zip.generateAsync({ type: "nodebuffer" })
}
