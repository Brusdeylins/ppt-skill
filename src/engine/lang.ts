/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  engine/lang: the language of the text pptc writes (`a:rPr@lang`, the
**  proofing and hyphenation language in PowerPoint). The deck language is
**  the custom property `pptcDeckLang` when it names a language, else the
**  language the deck's text already has (the presentation's default text
**  style, else the first run language on a slide), `en-US` only as the
**  last resort.
*/

import type JSZip from "jszip"
import { optText, readRels, slideList } from "./parts.js"
import { elements, firstElement, parseXml } from "./xml.js"

/**  the custom document property carrying the deck language  */
export const DECK_LANG_PROP = "pptcDeckLang"

/**  the language when nothing names one  */
export const DEFAULT_LANG = "en-US"

/**  a BCP 47 language tag as PowerPoint writes it ("de-DE")  */
const TAG = /^([A-Za-z]{2,3})(-[A-Za-z0-9]{2,8})*$/

/**  the usual region of a bare language code  */
const REGIONS: Record<string, string> = {
    de: "de-DE", en: "en-US", fr: "fr-FR", es: "es-ES", it: "it-IT", nl: "nl-NL",
    pt: "pt-PT", pl: "pl-PL", sv: "sv-SE", da: "da-DK", fi: "fi-FI", cs: "cs-CZ"
}

/**  language names (English and native) a deck setup may store instead of
     a tag  */
const NAMES: Record<string, string> = {
    german: "de", deutsch: "de", english: "en", englisch: "en", french: "fr",
    "français": "fr", francais: "fr", spanish: "es", "español": "es", espanol: "es",
    italian: "it", italiano: "it", dutch: "nl", nederlands: "nl", portuguese: "pt",
    "português": "pt", polish: "pl", polski: "pl", swedish: "sv", svenska: "sv",
    danish: "da", dansk: "da", finnish: "fi", suomi: "fi", czech: "cs", "čeština": "cs"
}

/**
 *  Normalize a stated language to a tag: a tag keeps its subtags (case
 *  normalized), a bare code or a language name gets its usual region.
 *
 *  @param value - a tag ("de-DE", "de") or a language name ("German")
 *  @returns the language tag, null when the value names no language
 */
export const normalizeLang = (value: string | null | undefined): string | null => {
    const v = (value ?? "").trim()
    const code = NAMES[v.toLowerCase()]
    if (code !== undefined)
        return REGIONS[code] ?? code
    if (!TAG.test(v))
        return null
    const [ lang, ...rest ] = v.split("-")
    const base = (lang as string).toLowerCase()
    if (rest.length === 0)
        return REGIONS[base] ?? base
    return [ base, ...rest.map((s) => s.length === 2 ? s.toUpperCase() : s) ].join("-")
}

/**
 *  Pick the deck language from its sources, in order of precedence.
 *
 *  @param custom - the `pptcDeckLang` custom property, if set
 *  @param presentation - the presentation's default text language, if any
 *  @param firstRun - the first run language found on a slide, if any
 *  @returns the language tag to write
 */
export const pickDeckLang = (custom: string | null | undefined, presentation: string | null | undefined,
    firstRun: string | null | undefined): string => {
    for (const lang of [ normalizeLang(custom), presentation, firstRun ])
        if (lang !== null && lang !== undefined && TAG.test(lang))
            return lang
    return DEFAULT_LANG
}

/**  the value of a custom document property, null when absent  */
const customProp = async (zip: JSZip, name: string): Promise<string | null> => {
    const xml = await optText(zip, "docProps/custom.xml")
    if (xml === null)
        return null
    try {
        const prop = elements(parseXml(xml), "property").find((p) => p.getAttribute("name") === name)
        return prop === undefined ? null : firstElement(prop, "vt:lpwstr")?.textContent ?? null
    }
    catch {
        return null
    }
}

/**  the language of the presentation's default text style: its level 1,
     else its default paragraph properties (where PowerPoint puts it)  */
const presentationLang = async (zip: JSZip): Promise<string | null> => {
    const xml = await optText(zip, "ppt/presentation.xml")
    const style = xml === null ? null : firstElement(parseXml(xml), "p:defaultTextStyle")
    for (const tag of [ "a:lvl1pPr", "a:defPPr" ]) {
        const pPr = style === null ? null : firstElement(style, tag)
        const lang = pPr === null ? null : firstElement(pPr, "a:defRPr")?.getAttribute("lang") ?? null
        if (lang !== null && lang !== "")
            return lang
    }
    return null
}

/**  the first run language found on a slide, in slide order  */
const firstRunLang = async (zip: JSZip): Promise<string | null> => {
    const xml = await optText(zip, "ppt/presentation.xml")
    if (xml === null)
        return null
    for (const { part } of slideList(zip, parseXml(xml), await readRels(zip, "ppt/presentation.xml"))) {
        const slide = part === null ? null : await optText(zip, part)
        if (slide === null)
            continue
        try {
            const lang = elements(parseXml(slide), "a:rPr").map((r) => r.getAttribute("lang") ?? "").find((l) => l !== "")
            if (lang !== undefined)
                return lang
        }
        catch {
            /*  a malformed slide names no language (verify reports it)  */
        }
    }
    return null
}

/**
 *  Resolve the language of the text an apply writes.
 *
 *  @param zip - the input deck
 *  @param customPatch - the custom properties the apply sets, if any
 *  @returns the language tag
 */
export const deckLang = async (zip: JSZip, customPatch: Record<string, string> | null): Promise<string> => {
    const custom = customPatch?.[DECK_LANG_PROP] ?? await customProp(zip, DECK_LANG_PROP)
    if (normalizeLang(custom) !== null)
        return pickDeckLang(custom, null, null)
    const presentation = await presentationLang(zip)
    return pickDeckLang(null, presentation,
        presentation !== null && TAG.test(presentation) ? null : await firstRunLang(zip))
}
