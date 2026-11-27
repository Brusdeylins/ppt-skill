/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/seed: the seed-deck factory. pptx-automizer imports slides, not
**  layouts -- so for every template we derive a "seed deck" once: a .pptx
**  containing exactly one empty slide per slide layout (placeholders cloned
**  from the layout). `slide.add` then imports seed slide N+1 for layout N.
**  Seeds are cached by template content hash.
*/

import JSZip from "jszip"
import { readFileSync, readdirSync, rmSync, statSync, utimesSync } from "node:fs"
import path from "node:path"
import type { Element } from "@xmldom/xmldom"
import { PptcError } from "../infra/errors.js"
import { seedPlaceholderName } from "../core/model.js"
import { atomicWrite, cacheDir, contentHash } from "../infra/fs.js"
import { cleanContentTypes } from "./parts.js"
import { gcAssets } from "./post.js"
import { NS_A, elements, firstElement, parseXml, serializeElement } from "./xml.js"

/**  seed format generation: bump to invalidate all cached seeds  */
const SEED_FORMAT = 4

/**  clone the fillable placeholders of a layout onto an empty slide (DOM)  */
const slideFromLayout = (layoutXml: string): string => {
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

/**  matches a presentation-rels slide relationship element  */
const SLIDE_REL_RE = /<Relationship [^>]*Type="[^"]*\/slide"[^>]*\/>/g

/**  strip every slide relationship, slide content-type override and slide part from a deck zip  */
const stripSlides = (
    presRels: string, ct: string
): { presRels: string, ct: string } => ({
    presRels: presRels.replace(SLIDE_REL_RE, ""),
    ct: ct.replace(/<Override PartName="\/ppt\/slides\/[^"]*"[^>]*\/>/g, "")
})

/**  layout part paths in slide-master order  */
const layoutOrder = async (zip: JSZip): Promise<string[]> => {
    const order: string[] = []
    for (const masterPart of Object.keys(zip.files).filter((f) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f)).sort()) {
        const master = parseXml(await (zip.file(masterPart) as JSZip.JSZipObject).async("string"))
        const relsFile = zip.file(`ppt/slideMasters/_rels/${path.posix.basename(masterPart)}.rels`)
        if (relsFile === null)
            continue
        const rels = parseXml(await relsFile.async("string"))
        const relMap = new Map(elements(rels, "Relationship")
            .map((r) => [r.getAttribute("Id") ?? "", r.getAttribute("Target") ?? ""]))
        for (const layoutId of elements(master, "p:sldLayoutId")) {
            const target = relMap.get(layoutId.getAttribute("r:id") ?? "")
            if (target !== undefined && target.includes("slideLayouts/"))
                order.push(path.posix.normalize(path.posix.join("ppt/slideMasters", target)))
        }
    }
    return order
}

/**
 *  Build a seed deck from a template archive: convert .potx content types,
 *  drop all existing slides/notes, then add one empty cloned-placeholder
 *  slide per layout under the canonical names slide1..slideN.
 *
 *  @param templateBytes - raw bytes of the .potx/.pptx template
 *  @returns seed deck bytes and the number of layouts
 *  @throws PptcError E_TEMPLATE when the template has no layouts
 */
export const buildSeed = async (templateBytes: Buffer): Promise<{ bytes: Buffer, layoutCount: number }> => {
    const zip = await JSZip.loadAsync(templateBytes)
    const get = async (part: string): Promise<string> => {
        const f = zip.file(part)
        if (f === null)
            throw new PptcError("E_TEMPLATE", `template is missing required part: ${part}`)
        return await f.async("string")
    }

    let ct = (await get("[Content_Types].xml"))
        .replace("presentationml.template.main+xml", "presentationml.presentation.main+xml")
    let pres = await get("ppt/presentation.xml")
    let presRels = await get("ppt/_rels/presentation.xml.rels")

    /*  drop existing slides and notes completely  */
    pres = pres
        .replace(/<p:sldIdLst\s*\/>/, "<p:sldIdLst></p:sldIdLst>")
        .replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/s, "<p:sldIdLst></p:sldIdLst>")
    if (!pres.includes("<p:sldIdLst>"))
        pres = pres.replace("</p:sldMasterIdLst>", "</p:sldMasterIdLst><p:sldIdLst></p:sldIdLst>")

    /*  drop PowerPoint sections: they group the template's example
        slides, which are stripped -- stale section definitions would
        reference removed sldIds  */
    pres = pres
        .replace(/<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}">.*?<\/p:ext>/s, "")
        .replace(/<p:extLst>\s*<\/p:extLst>/, "")

    /*  CT_Presentation is an ordered sequence: notesMasterIdLst (and
        handoutMasterIdLst) MUST precede sldIdLst -- sloppy templates
        violate this and PowerPoint answers with the repair dialog  */
    for (const lst of ["handoutMasterIdLst", "notesMasterIdLst"]) {
        const el = new RegExp(`<p:${lst}>.*?</p:${lst}>|<p:${lst}\\s*/>`, "s").exec(pres)
        if (el !== null && pres.indexOf(el[0]) > pres.indexOf("<p:sldIdLst"))
            pres = pres.replace(el[0], "").replace(/<p:sldIdLst/, `${el[0]}<p:sldIdLst`)
    }
    presRels = presRels.replace(SLIDE_REL_RE, "")
    for (const f of Object.keys(zip.files))
        if (/^ppt\/(slides|notesSlides)\//.test(f))
            zip.remove(f)
    ct = ct.replace(/<Override PartName="\/ppt\/(slides|notesSlides)\/[^"]*"[^>]*\/>/g, "")

    /*  templates saved in master/layout view carry that view along --
        drop lastView so the produced deck opens in normal view  */
    const viewPr = zip.file("ppt/viewProps.xml")
    if (viewPr !== null)
        zip.file("ppt/viewProps.xml",
            (await viewPr.async("string")).replace(/ lastView="[^"]*"/g, ""))

    /*  one seed slide per layout  */
    const order = await layoutOrder(zip)
    if (order.length === 0)
        throw new PptcError("E_TEMPLATE", "template contains no slide layouts")
    let sldIds = "", rels = "", overrides = ""
    for (let i = 0; i < order.length; i++) {
        const n = i + 1
        const layoutXml = await get(order[i] as string)
        zip.file(`ppt/slides/slide${n}.xml`, slideFromLayout(layoutXml))
        zip.file(`ppt/slides/_rels/slide${n}.xml.rels`,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../${(order[i] as string).substring(4)}"/>`
            + "</Relationships>")
        sldIds += `<p:sldId id="${9000 + n}" r:id="rIdPptcSeed${n}"/>`
        rels += `<Relationship Id="rIdPptcSeed${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/>`
        overrides += `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    }
    zip.file("[Content_Types].xml", ct.replace("</Types>", `${overrides}</Types>`))
    zip.file("ppt/presentation.xml", pres.replace("<p:sldIdLst></p:sldIdLst>", `<p:sldIdLst>${sldIds}</p:sldIdLst>`))
    zip.file("ppt/_rels/presentation.xml.rels", presRels.replace("</Relationships>", `${rels}</Relationships>`))

    /*  templates may carry stale or duplicate content-type overrides
        of their own -- sweep them so every derived deck starts clean; a
        seed derived from a deck also drops the media of its removed slides  */
    await gcAssets(zip)
    await cleanContentTypes(zip)

    return {
        bytes: await zip.generateAsync({ type: "nodebuffer" }),
        layoutCount: order.length
    }
}

/**
 *  Get the cached seed deck for a template, building it on first use.
 *  The cache key is the template's content hash, so template edits
 *  invalidate the seed automatically.
 *
 *  @param templatePath - path of the .potx/.pptx template
 *  @returns absolute path of the cached seed deck
 */
export const ensureSeed = async (templatePath: string): Promise<string> => {
    const bytes = readFileSync(templatePath)
    return await cachedSeed(`seed-${SEED_FORMAT}-${contentHash(bytes)}.pptx`, bytes)
}

/**  seeds unused for this long are dropped from the cache, and at most
     this many seeds are kept (each can carry a template's full media)  */
const SEED_MAX_AGE_MS = 14 * 24 * 3600 * 1000
const SEED_MAX_COUNT  = 24

/**  return the cached seed of a name, building it atomically on a miss (a
     concurrent pptc never reads a half-written seed); every use refreshes
     its mtime; stale seeds and orphaned temp files are pruned  */
const cachedSeed = async (name: string, source: Buffer): Promise<string> => {
    const dir = cacheDir()
    const seedPath = path.join(dir, name)
    try {
        const now = new Date()
        utimesSync(seedPath, now, now)
    }
    catch {
        atomicWrite(seedPath, (await buildSeed(source)).bytes)
    }
    const seeds: { file: string, mtime: number }[] = []
    for (const f of readdirSync(dir).filter((f) => f.startsWith("seed-")))
        try {
            const file = path.join(dir, f)
            const mtime = statSync(file).mtimeMs
            if (f.endsWith(".pptx") && file !== seedPath)
                seeds.push({ file, mtime })
            else if (!f.endsWith(".pptx") && Date.now() - mtime > 3600 * 1000)
                rmSync(file, { force: true })
        }
        catch {
            /*  another pptc pruned it first  */
        }
    seeds.sort((a, b) => b.mtime - a.mtime)
    for (const [ i, seed ] of seeds.entries())
        if (i >= SEED_MAX_COUNT - 1 || Date.now() - seed.mtime > SEED_MAX_AGE_MS)
            rmSync(seed.file, { force: true })
    return seedPath
}

/**
 *  Get the cached seed deck derived from an EXISTING deck's OWN embedded
 *  layouts -- used by `slide.add` when no external template is supplied.
 *  `pptc new` bakes the full template (masters, layouts, theme) into every
 *  deck and the post-pass never garbage-collects layouts, so the deck always
 *  carries the same layouts the template did. The cache key hashes only the
 *  design parts (masters, layouts, theme, presentation), so ordinary content
 *  edits do not rebuild the seed.
 *
 *  @param deckPath - path of the deck whose own layouts seed new slides
 *  @returns absolute path of the cached seed deck
 */
export const ensureSeedFromDeck = async (deckPath: string): Promise<string> => {
    const bytes = readFileSync(deckPath)
    const zip = await JSZip.loadAsync(bytes)
    const designParts = Object.keys(zip.files)
        .filter((f) => /^ppt\/(slideMasters|slideLayouts|theme)\/[^/]+\.xml$/.test(f)
            || f === "ppt/presentation.xml")
        .sort()
    const design: string[] = []

    /*  the slide list and sections of presentation.xml change with every
        slide add/remove but never affect the seed -- keep them out of the key  */
    for (const part of designParts) {
        let xml = await (zip.file(part) as JSZip.JSZipObject).async("string")
        if (part === "ppt/presentation.xml")
            xml = xml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, "")
                .replace(/<p14:sectionLst[\s\S]*?<\/p14:sectionLst>/, "")
        design.push(`${part}:${xml}`)
    }
    return await cachedSeed(`seed-deck-${SEED_FORMAT}-${contentHash(...design)}.pptx`, bytes)
}

/**
 *  Create an empty deck from a template: the seed deck with the slide list
 *  truncated and all seed slide parts removed. Used by `pptc new`.
 *
 *  @param templatePath - path of the .potx/.pptx template
 *  @returns bytes of a valid, zero-slide .pptx carrying the template's design
 */
export const buildEmptyDeck = async (templatePath: string): Promise<Buffer> => {
    const seed = await buildSeed(readFileSync(templatePath))
    const zip = await JSZip.loadAsync(seed.bytes)
    const pres = await (zip.file("ppt/presentation.xml") as JSZip.JSZipObject).async("string")
    const presRels = await (zip.file("ppt/_rels/presentation.xml.rels") as JSZip.JSZipObject).async("string")
    const ct = await (zip.file("[Content_Types].xml") as JSZip.JSZipObject).async("string")
    zip.file("ppt/presentation.xml", pres.replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/s, "<p:sldIdLst/>"))
    const stripped = stripSlides(presRels, ct)
    zip.file("ppt/_rels/presentation.xml.rels", stripped.presRels)
    zip.file("[Content_Types].xml", stripped.ct)
    for (const f of Object.keys(zip.files))
        if (/^ppt\/slides\//.test(f))
            zip.remove(f)
    return await zip.generateAsync({ type: "nodebuffer" })
}
