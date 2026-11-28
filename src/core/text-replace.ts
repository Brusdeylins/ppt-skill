/*
**  pptc -- Deterministic PowerPoint CLI for LLM Agents
**  Copyright (c) 2026 Matthias Brusdeylins
**  Licensed under MIT license <https://spdx.org/licenses/MIT>
**
**  core/text-replace: the plain-text find/replace of `text.replace` on a
**  sequence of runs. A match may span several runs (text split by
**  formatting or by PowerPoint's spell checker): the replacement goes into
**  the first touched run, which keeps its formatting, and the matched rest
**  leaves the following runs, which go when nothing is left of them. The
**  planner applies it to planned rich text, the engine to the runs of the
**  slide's text bodies -- both with the same semantics.
*/

import type { Paragraph, RichText } from "../schema/payloads.js"

/**  what to replace  */
export interface ReplaceSpec {
    /**  the plain text to find (never empty)  */
    find: string
    /**  its replacement (may be empty)  */
    replace: string
    /**  match upper/lower case exactly  */
    matchCase: boolean
    /**  match whole words only (no letter, digit or `_` around the match)  */
    wholeWord: boolean
}

/**  the outcome on a run sequence: the new text of every run, null for a
     run the replacement emptied (to be removed), and the match count  */
export interface RunsReplaced {
    texts: (string | null)[]
    count: number
}

/**  the global matcher of a spec  */
const matcher = (spec: ReplaceSpec): RegExp => {
    const escaped = spec.find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    const pattern = spec.wholeWord ? `(?<![\\p{L}\\p{N}_])${escaped}(?![\\p{L}\\p{N}_])` : escaped
    return new RegExp(pattern, spec.matchCase ? "gu" : "giu")
}

/**
 *  Whether a text may hold a match: true for every text that contains a
 *  match in any run sequence of it, so a false answer rules matches out
 *  (whole words are not checked).
 *
 *  @param text - the text (e.g. all run texts of a part, concatenated)
 *  @param spec - what to replace
 *  @returns false when no match is possible
 */
export const mayMatch = (text: string, spec: ReplaceSpec): boolean =>
    matcher({ ...spec, wholeWord: false }).test(text)

/**
 *  Replace every match in a sequence of runs read as one text.
 *
 *  @param texts - the text of each run in order
 *  @param spec - what to replace
 *  @returns the new run texts (null: the run is emptied and goes) and the
 *           number of matches; the texts are the input when nothing matched
 */
export const replaceRuns = (texts: string[], spec: ReplaceSpec): RunsReplaced => {
    const joined = texts.join("")
    const matches = [ ...joined.matchAll(matcher(spec)) ]
    if (matches.length === 0)
        return { texts, count: 0 }

    /*  the run owning each character  */
    const owner: number[] = []
    texts.forEach((t, i) => {
        for (let k = 0; k < t.length; k++)
            owner.push(i)
    })
    const out = texts.map(() => "")
    const touched = texts.map(() => false)
    const hosts = texts.map(() => false)
    let pos = 0
    for (const m of matches) {
        const start = m.index
        const end = start + m[0].length
        for (let k = pos; k < start; k++)
            out[owner[k] as number] += joined[k] as string
        const host = owner[start] as number
        out[host] += spec.replace
        hosts[host] = true
        for (let k = start; k < end; k++)
            touched[owner[k] as number] = true
        pos = end
    }
    for (let k = pos; k < joined.length; k++)
        out[owner[k] as number] += joined[k] as string
    return {
        texts: out.map((t, i) => t === "" && (touched[i] === true || hosts[i] === true) ? null : t),
        count: matches.length
    }
}

/**
 *  Replace in one plain string, line by line (a match never spans a line
 *  break, as it never spans a paragraph).
 *
 *  @param text - the text
 *  @param spec - what to replace
 *  @returns the new text and the number of matches
 */
export const replacePlain = (text: string, spec: ReplaceSpec): { text: string, count: number } => {
    let count = 0
    const lines = text.split("\n").map((line) => {
        const done = replaceRuns([ line ], spec)
        count += done.count
        return done.texts[0] ?? ""
    })
    return { text: lines.join("\n"), count }
}

/**
 *  Replace in rich text: per paragraph across its runs, keeping every
 *  run's formatting; a run the replacement empties goes.
 *
 *  @param text - plain or rich text
 *  @param spec - what to replace
 *  @returns the new text (the input when nothing matched) and the count
 */
export const replaceRichText = (text: RichText, spec: ReplaceSpec): { text: RichText, count: number } => {
    if (typeof text === "string")
        return replacePlain(text, spec)
    let count = 0
    const paragraphs = text.map((p): Paragraph => {
        if (p.runs === undefined) {
            const done = replacePlain(p.text ?? "", spec)
            count += done.count
            return done.count === 0 ? p : { ...p, text: done.text }
        }
        const done = replaceRuns(p.runs.map((r) => r.text), spec)
        count += done.count
        if (done.count === 0)
            return p
        return { ...p, runs: p.runs.flatMap((r, i) => {
            const t = done.texts[i]
            return t === null || t === undefined ? [] : [ { ...r, text: t } ]
        }) }
    })
    return { text: count === 0 ? text : paragraphs, count }
}
