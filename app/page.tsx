"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { extractDocument, type DocumentExtraction } from "./lib/document-text";
import type { UpdatedDocument } from "./lib/document-export";
import { findFigures, readFigureTarget, rebalanceFigureChanges, type FigureCalculation, type FigureChange, type FigureStrategy } from "./lib/figure-engine";
import { LedgerWorkspace, type LedgerSource } from "./ledger-workspace";
import { DocxDocumentPreview } from "./docx-preview";

export const dynamic = "force-static";

type Direction = "forward" | "backward";
type TaskMode = "dates" | "figures" | "both" | "ledger";
type View = "upload" | "processing" | "review" | "complete";
type DateMatch = {
  id: number;
  sourceIndex: number;
  original: string;
  updated: string;
  context: string;
  reason: string;
  confidence: number;
  selected: boolean;
};

function isFigureChange(change: DateMatch | FigureChange): change is FigureChange {
  return "amountMinor" in change && typeof change.amountMinor === "bigint";
}

export type InstructionPlan = {
  months: number;
  direction: Direction;
  amount: number;
  excludeTerms: string[];
  includeTerms: string[];
  preserveMonthEnd: boolean;
  summary: string;
};

const demoText = `SERVICE AGREEMENT 2026
This Service Agreement becomes effective on March 12, 2026 and remains active through 31 December 2026.
The first reporting deadline is 30 April 2026. A quarterly review is scheduled for 06/15/2026.
All renewal requests must be received by 2026-08-01. The final evaluation will take place on November 30, 2026.
The agreement was signed on March 5, 2026.
Cash allocation: GHS 12,500.00
Operating reserve: GHS 7,500.00`;

const monthsFull = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const monthLookup: Record<string, number> = Object.fromEntries(monthsFull.flatMap((month, index) => [[month.toLowerCase(), index], [month.slice(0, 3).toLowerCase(), index]]));
const monthPattern = "(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)";
const weekdayPattern = "(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun)";
// PDF and Word extractors may insert spaces or use typographic dashes even
// when a date looks like `01-Dec-2025` on the page.
const flexibleDateSeparator = "[\\s\\u200B\\u2060]*[-‐‑‒–—−/.][\\s\\u200B\\u2060]*";
const weekdaysFull = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const weekdaysShort = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const numberWords: Record<string, number> = {
  one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8,
  nine: 9, ten: 10, eleven: 11, twelve: 12, thirteen: 13, fourteen: 14,
  fifteen: 15, sixteen: 16, seventeen: 17, eighteen: 18, nineteen: 19,
  twenty: 20, "twenty-one": 21, "twenty-two": 22, "twenty-three": 23, "twenty-four": 24,
};

function daysInMonth(year: number, month: number) {
  return new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
}

function shiftParts(year: number, month: number, day: number, amount: number, preserveMonthEnd = false) {
  const absolute = month + amount;
  const nextYear = year + Math.floor(absolute / 12);
  const nextMonth = ((absolute % 12) + 12) % 12;
  const wasMonthEnd = day === daysInMonth(year, month);
  return {
    year: nextYear,
    month: nextMonth,
    day: preserveMonthEnd && wasMonthEnd ? daysInMonth(nextYear, nextMonth) : Math.min(day, daysInMonth(nextYear, nextMonth)),
  };
}

function ordinal(day: number) {
  if (day % 100 >= 11 && day % 100 <= 13) return `${day}th`;
  return `${day}${day % 10 === 1 ? "st" : day % 10 === 2 ? "nd" : day % 10 === 3 ? "rd" : "th"}`;
}

function matchCase(value: string, template: string) {
  if (template === template.toUpperCase()) return value.toUpperCase();
  if (template === template.toLowerCase()) return value.toLowerCase();
  return value;
}

function monthForTemplate(month: number, template: string) {
  const abbreviated = template.length <= 3 || /^sept$/i.test(template);
  const value = abbreviated ? monthsFull[month].slice(0, 3) : monthsFull[month];
  return matchCase(value, template);
}

function shiftValidParts(year: number, month: number, day: number, amount: number, preserveMonthEnd: boolean) {
  if (!Number.isInteger(year) || year < 1000 || year > 9999 || !Number.isInteger(month) || month < 0 || month > 11 || !Number.isInteger(day) || day < 1 || day > 31) return null;
  const shifted = shiftParts(year, month, Math.min(day, daysInMonth(year, month)), amount, preserveMonthEnd);
  return shifted.year >= 1000 && shifted.year <= 9999 ? shifted : null;
}

function weekdayForDate(year: number, month: number, day: number, template: string) {
  const weekday = new Date(Date.UTC(year, month, day)).getUTCDay();
  const value = template.length > 5 ? weekdaysFull[weekday] : weekdaysShort[weekday];
  return matchCase(value, template);
}

export function shiftDateLabel(value: string, amount: number, preserveMonthEnd = false) {
  const timestampDate = value.match(/^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun)(,?)\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+(\d{1,2})(st|nd|rd|th)?(,)?\s+(\d{2}:\d{2}(?::\d{2})?)\s+((?:GMT|UTC)(?:[+-]\d{4})?)\s+(\d{4})$/i);
  if (timestampDate) {
    const [, rawWeekday, weekdayComma, rawMonth, rawDay, suffix, dayComma, time, timezone, rawYear] = timestampDate;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    const weekdayName = weekdayForDate(shifted.year, shifted.month, shifted.day, rawWeekday);
    const monthName = monthForTemplate(shifted.month, rawMonth);
    const day = suffix ? ordinal(shifted.day) : String(shifted.day).padStart(rawDay.length, "0");
    return `${weekdayName}${weekdayComma} ${monthName} ${day}${dayComma ?? ""} ${time} ${timezone} ${shifted.year}`;
  }

  const weekdayFirst = value.match(/^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun)(,?)\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+(\d{1,2})(st|nd|rd|th)?(,)?\s+(\d{4})$/i);
  if (weekdayFirst) {
    const [, rawWeekday, weekdayComma, rawMonth, rawDay, suffix, dayComma, rawYear] = weekdayFirst;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    const weekdayName = weekdayForDate(shifted.year, shifted.month, shifted.day, rawWeekday);
    const monthName = monthForTemplate(shifted.month, rawMonth);
    const day = suffix ? ordinal(shifted.day) : String(shifted.day).padStart(rawDay.length, "0");
    return `${weekdayName}${weekdayComma} ${monthName} ${day}${dayComma ?? ""} ${shifted.year}`;
  }

  const weekdayDayFirst = value.match(/^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun)(,?)\s+(\d{1,2})(st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?(,)?\s+(\d{4})$/i);
  if (weekdayDayFirst) {
    const [, rawWeekday, weekdayComma, rawDay, suffix, rawMonth, dayComma, rawYear] = weekdayDayFirst;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    const weekdayName = weekdayForDate(shifted.year, shifted.month, shifted.day, rawWeekday);
    const monthName = monthForTemplate(shifted.month, rawMonth);
    const day = suffix ? ordinal(shifted.day) : String(shifted.day).padStart(rawDay.length, "0");
    return `${weekdayName}${weekdayComma} ${day} ${monthName}${dayComma ?? ""} ${shifted.year}`;
  }

  const monthFirst = value.match(/^(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+(\d{1,2})(st|nd|rd|th)?(,)?\s+(\d{4})$/i);
  if (monthFirst) {
    const [, rawMonth, rawDay, suffix, comma, rawYear] = monthFirst;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    const monthName = monthForTemplate(shifted.month, rawMonth);
    const day = suffix ? ordinal(shifted.day) : String(shifted.day).padStart(rawDay.length, "0");
    return `${monthName} ${day}${comma ?? ""} ${shifted.year}`;
  }

  const dayFirst = value.match(/^(\d{1,2})(st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?(,)?\s+(\d{4})$/i);
  if (dayFirst) {
    const [, rawDay, suffix, rawMonth, comma, rawYear] = dayFirst;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    const monthName = monthForTemplate(shifted.month, rawMonth);
    const day = suffix ? ordinal(shifted.day) : String(shifted.day).padStart(rawDay.length, "0");
    return `${day} ${monthName}${comma ?? ""} ${shifted.year}`;
  }

  const dayNamedSeparated = value.match(new RegExp(`^(\\d{1,2})(${flexibleDateSeparator})(${monthPattern})(${flexibleDateSeparator})(\\d{4})$`, "i"));
  if (dayNamedSeparated) {
    const [, rawDay, firstSeparator, rawMonth, secondSeparator, rawYear] = dayNamedSeparated;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    return `${String(shifted.day).padStart(rawDay.length, "0")}${firstSeparator}${monthForTemplate(shifted.month, rawMonth)}${secondSeparator}${shifted.year}`;
  }

  const monthNamedSeparated = value.match(new RegExp(`^(${monthPattern})(${flexibleDateSeparator})(\\d{1,2})(${flexibleDateSeparator})(\\d{4})$`, "i"));
  if (monthNamedSeparated) {
    const [, rawMonth, firstSeparator, rawDay, secondSeparator, rawYear] = monthNamedSeparated;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    return `${monthForTemplate(shifted.month, rawMonth)}${firstSeparator}${String(shifted.day).padStart(rawDay.length, "0")}${secondSeparator}${shifted.year}`;
  }

  const yearNamedSeparated = value.match(new RegExp(`^(\\d{4})(${flexibleDateSeparator})(${monthPattern})(${flexibleDateSeparator})(\\d{1,2})$`, "i"));
  if (yearNamedSeparated) {
    const [, rawYear, firstSeparator, rawMonth, secondSeparator, rawDay] = yearNamedSeparated;
    const shifted = shiftValidParts(Number(rawYear), monthLookup[rawMonth.slice(0, 3).toLowerCase()], Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    return `${shifted.year}${firstSeparator}${monthForTemplate(shifted.month, rawMonth)}${secondSeparator}${String(shifted.day).padStart(rawDay.length, "0")}`;
  }

  const yearFirstNumeric = value.match(new RegExp(`^(\\d{4})(${flexibleDateSeparator})(\\d{1,2})(${flexibleDateSeparator})(\\d{1,2})$`));
  if (yearFirstNumeric) {
    const [, rawYear, firstSeparator, rawMonth, secondSeparator, rawDay] = yearFirstNumeric;
    const shifted = shiftValidParts(Number(rawYear), Number(rawMonth) - 1, Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    return `${shifted.year}${firstSeparator}${String(shifted.month + 1).padStart(rawMonth.length, "0")}${secondSeparator}${String(shifted.day).padStart(rawDay.length, "0")}`;
  }

  const numeric = value.match(new RegExp(`^(\\d{1,2})(${flexibleDateSeparator})(\\d{1,2})(${flexibleDateSeparator})(\\d{4})$`));
  if (numeric) {
    const first = Number(numeric[1]);
    const second = Number(numeric[3]);
    const separatorCharacter = numeric[2].match(/[-‐‑‒–—−/.]/)?.[0] ?? "/";
    const dayFirstFormat = first > 12 || (second <= 12 && !/[/.]/.test(separatorCharacter));
    const rawDay = dayFirstFormat ? numeric[1] : numeric[3];
    const rawMonth = dayFirstFormat ? numeric[3] : numeric[1];
    const shifted = shiftValidParts(Number(numeric[5]), Number(rawMonth) - 1, Number(rawDay), amount, preserveMonthEnd);
    if (!shifted) return value;
    const month = String(shifted.month + 1).padStart(rawMonth.length, "0");
    const day = String(shifted.day).padStart(rawDay.length, "0");
    return dayFirstFormat ? `${day}${numeric[2]}${month}${numeric[4]}${shifted.year}` : `${month}${numeric[2]}${day}${numeric[4]}${shifted.year}`;
  }
  return value;
}

export function interpretInstruction(instruction: string, defaultMonths: number, defaultDirection: Direction): InstructionPlan {
  const normalized = instruction.toLowerCase().replace(/[’']/g, "'");
  const durationPattern = /(?<!\w)([-+−]?\d{1,2}|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty(?:-(?:one|two|three|four))?)\s*(?:calendar\s+)?months?\b/g;
  const numberMatch = Array.from(normalized.matchAll(durationPattern)).at(-1);
  const rawNumber = numberMatch?.[1];
  const parsedMonths = rawNumber ? (Math.abs(Number(rawNumber.replace("−", "-"))) || numberWords[rawNumber] || defaultMonths) : defaultMonths;
  const months = Math.max(1, Math.min(24, parsedMonths));

  let direction = defaultDirection;
  const durationIndex = numberMatch?.index ?? 0;
  const clauseStart = Math.max(normalized.lastIndexOf(".", durationIndex) + 1, normalized.lastIndexOf(";", durationIndex) + 1);
  const nearDuration = numberMatch ? normalized.slice(Math.max(clauseStart, durationIndex - 36), durationIndex + numberMatch[0].length + 24).split(/[.;\n]/, 1)[0] : normalized;
  const requestedDirection = nearDuration.match(/\b(backwards?|back|earlier|subtract|minus|forward|later|add|plus|ahead)\b/);
  if (requestedDirection) direction = /^(back|earlier|subtract|minus)/.test(requestedDirection[0]) ? "backward" : "forward";
  else if (/\b(?:before|previous|prior)\b/.test(nearDuration)) direction = "backward";
  else if (/\bafter\b/.test(nearDuration)) direction = "forward";
  if (rawNumber && /^[-−]/.test(rawNumber)) direction = "backward";
  if (rawNumber?.startsWith("+")) direction = "forward";

  const contextTerms: Array<{ test: RegExp; terms: string[]; label: string }> = [
    { test: /sign(?:ed|ature|ing)?/, terms: ["sign", "signature", "signed"], label: "signature dates" },
    { test: /effective|commencement|start/, terms: ["effective", "commencement", "start"], label: "effective dates" },
    { test: /expir(?:y|ation|es)|end date/, terms: ["expiry", "expiration", "expires", "end date"], label: "expiry dates" },
    { test: /renewal|renew/, terms: ["renewal", "renew"], label: "renewal dates" },
    { test: /payment|invoice|due date/, terms: ["payment", "invoice", "due"], label: "payment dates" },
    { test: /deadline|reporting/, terms: ["deadline", "reporting"], label: "deadline dates" },
  ];
  const exclusionLanguage = /\b(skip|exclude|except|unchanged|do not|don't|leave|ignore)\b/;
  const onlyLanguage = /\b(only|just)\b/;
  const excludeTerms: string[] = [];
  const includeTerms: string[] = [];
  const excludedLabels: string[] = [];
  const includedLabels: string[] = [];

  for (const context of contextTerms) {
    const match = normalized.match(context.test);
    if (!match || match.index === undefined) continue;
    const nearbyInstruction = normalized.slice(Math.max(0, match.index - 42), match.index + match[0].length + 34);
    if (exclusionLanguage.test(nearbyInstruction)) {
      excludeTerms.push(...context.terms);
      excludedLabels.push(context.label);
    } else if (onlyLanguage.test(nearbyInstruction)) {
      includeTerms.push(...context.terms);
      includedLabels.push(context.label);
    }
  }

  const preserveMonthEnd = /month[ -]?end|end of (?:the )?(?:new )?month|last day of (?:the )?month/.test(normalized);
  const directionLabel = direction === "forward" ? "forward" : "backward";
  let summary = `Move ${includedLabels.length ? `only ${includedLabels.join(" and ")}` : "all detected dates"} ${directionLabel} by ${months} month${months === 1 ? "" : "s"}`;
  if (excludedLabels.length) summary += `, leaving ${excludedLabels.join(" and ")} unchanged`;
  if (preserveMonthEnd) summary += ", preserving month-end dates";
  summary += ".";

  return { months, direction, amount: months * (direction === "forward" ? 1 : -1), excludeTerms, includeTerms, preserveMonthEnd, summary };
}

function contextAround(text: string, index: number, length: number) {
  const beforeBreak = Math.max(text.lastIndexOf(".", index - 1) + 1, text.lastIndexOf("\n", index - 1) + 1, index - 70, 0);
  const period = text.indexOf(".", index + length);
  const newline = text.indexOf("\n", index + length);
  const stops = [period, newline].filter((value) => value >= 0);
  const nextBreak = stops.length ? Math.min(...stops) + 1 : text.length;
  const end = Math.min(nextBreak, index + length + 90);
  let context = text.slice(beforeBreak, end).replace(/\s+/g, " ").trim();
  if (context.length > 145) context = `${context.slice(0, 142)}…`;
  return context;
}

function isSingleEdit(value: string, known: string) {
  if (Math.abs(value.length - known.length) > 1) return false;
  let index = 0;
  while (value[index] !== undefined && value[index] === known[index]) index++;
  if (known.length > value.length) return value.slice(index) === known.slice(index + 1);
  if (value.length > known.length) return value.slice(index + 1) === known.slice(index);
  return value.slice(index + 1) === known.slice(index + 1) ||
    (value[index] === known[index + 1] && value[index + 1] === known[index] && value.slice(index + 2) === known.slice(index + 2));
}

const possibleMonthTypo = new RegExp(`\\b\\d{1,2}${flexibleDateSeparator}([a-z]{3,10})${flexibleDateSeparator}\\d{4}\\b`, "gi");

function correctNamedMonth(value: string) {
  const token = Array.from(value.matchAll(possibleMonthTypo))[0]?.[1];
  if (!token || monthLookup[token.toLowerCase()] !== undefined || /^sept$/i.test(token)) return null;
  const candidates = monthsFull.filter((name) => isSingleEdit(token.toLowerCase(), name.toLowerCase()) || isSingleEdit(token.toLowerCase(), name.slice(0, 3).toLowerCase()));
  if (candidates.length !== 1) return null;
  return value.replace(token, matchCase(token.length <= 4 ? candidates[0].slice(0, 3) : candidates[0], token));
}

export function findDates(text: string, plan: InstructionPlan) {
  const clockTime = "\\d{2}:\\d{2}(?::\\d{2})?";
  const timezone = "(?:GMT|UTC)(?:[+-]\\d{4})?";
  const patterns = [
    new RegExp(`\\b${weekdayPattern},?\\s+${monthPattern}\\.?\\s+\\d{1,2}(?:st|nd|rd|th)?[,]?\\s+${clockTime}\\s+${timezone}\\s+\\d{4}\\b`, "gi"),
    new RegExp(`\\b(?:${weekdayPattern},?\\s+)?${monthPattern}\\.?\\s+\\d{1,2}(?:st|nd|rd|th)?[,]?\\s+\\d{4}\\b`, "gi"),
    new RegExp(`\\b${weekdayPattern},?\\s+\\d{1,2}(?:st|nd|rd|th)?\\s+${monthPattern}\\.?[,]?\\s+\\d{4}\\b`, "gi"),
    new RegExp(`\\b\\d{1,2}(?:st|nd|rd|th)?\\s+${monthPattern}\\.?[,]?\\s+\\d{4}\\b`, "gi"),
    new RegExp(`\\b\\d{1,2}${flexibleDateSeparator}${monthPattern}${flexibleDateSeparator}\\d{4}\\b`, "gi"),
    new RegExp(`\\b${monthPattern}${flexibleDateSeparator}\\d{1,2}${flexibleDateSeparator}\\d{4}\\b`, "gi"),
    new RegExp(`\\b\\d{4}${flexibleDateSeparator}${monthPattern}${flexibleDateSeparator}\\d{1,2}\\b`, "gi"),
    new RegExp(`\\b\\d{4}${flexibleDateSeparator}\\d{1,2}${flexibleDateSeparator}\\d{1,2}\\b`, "g"),
    new RegExp(`\\b\\d{1,2}${flexibleDateSeparator}\\d{1,2}${flexibleDateSeparator}\\d{4}\\b`, "g"),
    possibleMonthTypo,
  ];
  const hits: Array<{ value: string; index: number }> = [];
  patterns.forEach((pattern) => {
    for (const match of text.matchAll(pattern)) hits.push({ value: match[0], index: match.index ?? 0 });
  });

  let skipped = 0;
  let needsReview = 0;
  let previousEnd = -1;
  const dates = hits
    .sort((a, b) => a.index - b.index || b.value.length - a.value.length)
    .flatMap((hit) => {
      if (hit.index < previousEnd || hit.value.includes("\n")) return [];
      const correctedMonth = correctNamedMonth(hit.value);
      const source = correctedMonth ?? hit.value;
      const updated = shiftDateLabel(source, plan.amount, plan.preserveMonthEnd);
      if (updated === source) return [];
      previousEnd = hit.index + hit.value.length;
      const normalizedDate = shiftDateLabel(source, 0);
      const correctedDay = source.match(/\d+/g)?.map(Number).join("/") !== normalizedDate.match(/\d+/g)?.map(Number).join("/");
      const weekday = source.match(new RegExp(`^${weekdayPattern}\\b`, "i"))?.[0];
      const correctedWeekday = weekday && !normalizedDate.toLowerCase().startsWith(weekday.slice(0, 3).toLowerCase());
      const context = contextAround(text, hit.index, hit.value.length);
      const normalizedContext = context.toLowerCase();
      const isExcluded = plan.excludeTerms.some((term) => normalizedContext.includes(term));
      const isIncluded = !plan.includeTerms.length || plan.includeTerms.some((term) => normalizedContext.includes(term));
      const matchesInstruction = !isExcluded && isIncluded;
      const correction = correctedMonth || correctedDay;
      const shouldUpdate = matchesInstruction && !correction;
      if (!matchesInstruction) skipped += 1;
      else if (correction) needsReview += 1;
      return [{
        id: 0,
        sourceIndex: hit.index,
        original: hit.value,
        updated,
        context,
        reason: !matchesInstruction ? "Left unchanged because it does not match the instruction."
          : correction ? `Review this proposed correction. Interpreted the source as ${normalizedDate} before shifting.`
          : `Moved ${plan.direction} ${plan.months} month${plan.months === 1 ? "" : "s"}.${correctedWeekday ? " Corrected the weekday using the calendar date." : ""}`,
        confidence: correction ? 70 : /[A-Za-z]/.test(hit.value) || /^\d{4}-/.test(hit.value) ? 99 : 94,
        selected: shouldUpdate,
      }];
    })
    .map((date, index) => ({ ...date, id: index + 1 }));

  return { dates, skipped, needsReview };
}

function describeDateResult(result: ReturnType<typeof findDates>) {
  const count = result.dates.filter((date) => date.selected).length;
  const detected = count ? `I found ${count} date${count === 1 ? "" : "s"} ready to update.` : "No dates were selected automatically.";
  const review = result.needsReview ? ` Review ${result.needsReview} proposed date correction${result.needsReview === 1 ? "" : "s"} before export.` : "";
  const skipped = result.skipped ? ` ${result.skipped} date${result.skipped === 1 ? " was" : "s were"} excluded by your instruction.` : "";
  return detected + review + skipped;
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [task, setTask] = useState<TaskMode>("dates");
  const [file, setFile] = useState<File | null>(null);
  const [fileName, setFileName] = useState("");
  const [extraction, setExtraction] = useState<DocumentExtraction | null>(null);
  const [months, setMonths] = useState(3);
  const [direction, setDirection] = useState<Direction>("forward");
  const [targetTotal, setTargetTotal] = useState("");
  const [figureStrategy, setFigureStrategy] = useState<FigureStrategy>("proportional");
  const [figureCalculation, setFigureCalculation] = useState<FigureCalculation | null>(null);
  const [ledgerSource, setLedgerSource] = useState<LedgerSource | null>(null);
  const [instruction, setInstruction] = useState("");
  const [refinement, setRefinement] = useState("");
  const [view, setView] = useState<View>("upload");
  const [progress, setProgress] = useState(0);
  const [dates, setDates] = useState<Array<DateMatch | FigureChange>>([]);
  const [notice, setNotice] = useState("");
  const [assistantMessage, setAssistantMessage] = useState("");
  const [interpretedInstruction, setInterpretedInstruction] = useState("");
  const [refining, setRefining] = useState(false);
  const [preparingPreview, setPreparingPreview] = useState(false);
  const [preparedDocument, setPreparedDocument] = useState<UpdatedDocument | null>(null);
  const [previewUrl, setPreviewUrl] = useState("");
  const [exportError, setExportError] = useState("");
  const [exportedFileName, setExportedFileName] = useState("");
  const [downloadedCount, setDownloadedCount] = useState(0);
  const previewUrlRef = useRef("");

  const includesDates = task === "dates" || task === "both";
  const includesFigures = task === "figures" || task === "both";
  const isCombinedLedger = task === "both" && ledgerSource !== null;
  const resultType = task === "both" ? "updates" : task === "dates" ? "dates" : "figures";
  const signedMonths = months * (direction === "forward" ? 1 : -1);
  const selectedCount = dates.filter((date) => date.selected).length;
  const selectedFigureCount = dates.filter((change) => change.selected && isFigureChange(change)).length;
  const exampleDate = useMemo(() => shiftDateLabel("12 March 2026", signedMonths), [signedMonths]);
  const fallbackInstruction = `Move every calendar date ${direction} by ${months} month${months === 1 ? "" : "s"}. Preserve the original date format.`;
  const figureFallbackInstruction = targetTotal.trim()
    ? `Make all detected financial figures total ${targetTotal.trim()} ${figureStrategy === "proportional" ? "proportionally" : figureStrategy === "equal" ? "by sharing the difference equally" : "by changing only the final figure"}.`
    : "";
  const effectiveInstruction = instruction.trim() || [includesDates ? fallbackInstruction : "", includesFigures ? figureFallbackInstruction : ""].filter(Boolean).join(" ");
  const canAnalyze = Boolean(fileName && effectiveInstruction && (task !== "both" || readFigureTarget(instruction, targetTotal)));

  useEffect(() => {
    return () => {
      if (previewUrlRef.current) URL.revokeObjectURL(previewUrlRef.current);
    };
  }, []);

  /* eslint-disable react-hooks/set-state-in-effect -- This effect owns the
     asynchronous preview lifecycle and must clear stale blob state whenever
     the selected dates, source file, or view changes. */
  useEffect(() => {
    let cancelled = false;
    if (view === "complete") {
      setPreparingPreview(false);
      return;
    }
    if (previewUrlRef.current) {
      URL.revokeObjectURL(previewUrlRef.current);
      previewUrlRef.current = "";
    }
    setPreviewUrl("");
    setPreparedDocument(null);
    setExportError("");
    if (view !== "review" || !extraction || !selectedCount) {
      setPreparingPreview(false);
      return;
    }

    setPreparingPreview(true);
    const timer = window.setTimeout(async () => {
      try {
        const { createUpdatedDocument } = await import("./lib/document-export");
        const changes = dates.map((change) => ({ ...change, kind: isFigureChange(change) ? "figure" as const : "date" as const }));
        const result = await createUpdatedDocument(file, extraction, changes, selectedFigureCount > 0);
        if (cancelled) return;
        if ((selectedFigureCount || task === "both") && result.updatedCount !== selectedCount) {
          throw new Error("Some selected values could not be written. Review the selection before downloading.");
        }

        const url = URL.createObjectURL(result.blob);
        previewUrlRef.current = url;
        setPreparedDocument(result);
        setPreviewUrl(url);
        setExportError(result.warning ?? "");
      } catch (error) {
        if (!cancelled) setExportError(error instanceof Error ? error.message : "The updated document preview could not be created.");
      } finally {
        if (!cancelled) setPreparingPreview(false);
      }
    }, 350);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [dates, extraction, file, selectedCount, selectedFigureCount, task, view]);
  /* eslint-enable react-hooks/set-state-in-effect */

  function acceptFile(nextFile: File) {
    const extension = nextFile.name.split(".").pop()?.toLowerCase();
    if (!extension || !["docx", "pdf"].includes(extension)) {
      setNotice("Please choose a Word .docx or PDF document. Save older .doc files as .docx first.");
      return;
    }
    if (nextFile.size > 25 * 1024 * 1024) {
      setNotice("This file is larger than 25 MB. Please choose a smaller document.");
      return;
    }
    setFile(nextFile);
    setFileName(nextFile.name);
    setExtraction(null);
    setNotice("");
    setDates([]);
    setFigureCalculation(null);
  }

  function useDemo() {
    setFile(null);
    setFileName("Service_Agreement_2026.docx");
    setExtraction({ kind: "docx", text: demoText, pdfSpans: [] });
    setNotice("");
    setDates([]);
    setFigureCalculation(null);
    if (includesFigures && !targetTotal.trim()) setTargetTotal("GHS 30,000.00");
  }

  function addPrompt(prompt: string) {
    setInstruction((current) => current.trim() ? `${current.trim()} ${prompt}` : prompt);
  }

  async function requestAnalysis(nextInstruction: string, showProcessing: boolean) {
    if (!fileName) return;
    if (showProcessing) {
      setView("processing");
      setProgress(12);
    } else {
      setRefining(true);
    }
    setNotice("");
    const interval = showProcessing
      ? window.setInterval(() => setProgress((current) => Math.min(86, current + (current < 50 ? 10 : 4))), 240)
      : undefined;

    try {
      let currentExtraction = extraction;
      if (!currentExtraction) {
        currentExtraction = file
          ? await extractDocument(file)
          : { kind: "docx", text: demoText, pdfSpans: [] };
        setExtraction(currentExtraction);
      }
      const text = currentExtraction.text;
      if (!text.trim()) {
        throw new Error("No readable text was found. If this is a scanned PDF, use a text-searchable PDF or export it with OCR first.");
      }

      const plan = includesDates ? interpretInstruction(nextInstruction, months, direction) : null;
      const dateResult = plan ? findDates(text, plan) : null;
      if (includesFigures) {
        if (file) {
          const { hasTransactionColumns, readLedger } = await import("./lib/ledger-docx");
          if (hasTransactionColumns(text)) {
            if (currentExtraction.kind !== "docx") throw new Error("This PDF has transaction balance columns. Upload a Word (.docx) copy to recalculate the running balances.");
            const ledger = await readLedger(file);
            const target = readFigureTarget(nextInstruction, targetTotal)?.label;
            setLedgerSource({ name: file.name, file, ledger, target, dateReview: dateResult && plan ? { changes: dateResult.dates, summary: plan.summary } : undefined });
            if (task === "figures") setTask("ledger");
            setView("upload");
            return;
          }
        }
      }
      const figureResult = includesFigures ? findFigures(text, nextInstruction, targetTotal, figureStrategy, dateResult?.dates) : null;
      const updates = [...(dateResult?.dates ?? []), ...(figureResult?.changes ?? [])]
        .sort((a, b) => a.sourceIndex - b.sourceIndex)
        .map((change, index) => ({ ...change, id: index + 1 }));
      setDates(updates);
      setFigureCalculation(figureResult?.calculation ?? null);
      const dateMessage = dateResult ? describeDateResult(dateResult) : "";
      setAssistantMessage([dateMessage, figureResult?.assistantMessage].filter(Boolean).join(" "));
      setInterpretedInstruction([plan?.summary, figureResult?.summary].filter(Boolean).join(" "));
      if (figureResult?.calculation) {
        setTargetTotal(figureResult.calculation.targetDisplay);
        setFigureStrategy(figureResult.calculation.strategy);
        if (task === "figures") setInstruction("");
      }
      if (plan) {
        setMonths(plan.months);
        setDirection(plan.direction);
        if (instruction.trim() || !showProcessing) setInstruction(nextInstruction);
      }
      setProgress(100);
      setView("review");
    } catch (error) {
      setNotice(error instanceof Error ? error.message : "The document could not be read.");
      setView("upload");
    } finally {
      if (interval) window.clearInterval(interval);
      setRefining(false);
    }
  }

  function analyze() {
    void requestAnalysis(effectiveInstruction, true);
  }

  function updateDate(id: number, value: string) {
    setDates((current) => current.map((date) => date.id === id ? { ...date, updated: value } : date));
  }

  function rebalanceSelection(next: Array<DateMatch | FigureChange>) {
    if (!figureCalculation) return next;
    const figures = rebalanceFigureChanges(next.filter(isFigureChange), figureCalculation);
    const byId = new Map(figures.map((change) => [change.id, change]));
    return next.map((change) => byId.get(change.id) ?? change);
  }

  function updateSelection(id: number, selected: boolean) {
    setDates((current) => rebalanceSelection(current.map((item) => item.id === id ? { ...item, selected } : item)));
  }

  function toggleAllSelections() {
    setDates((current) => {
      const selected = selectedCount !== current.length;
      return rebalanceSelection(current.map((item) => ({ ...item, selected })));
    });
  }

  function switchTask(nextTask: TaskMode) {
    if (nextTask === task) return;
    setTask(nextTask);
    setInstruction("");
    setRefinement("");
    setDates([]);
    setFigureCalculation(null);
    setLedgerSource(null);
    setNotice("");
    setAssistantMessage("");
    setInterpretedInstruction("");
    setPreparedDocument(null);
    setPreviewUrl("");
    setExportError("");
  }

  function downloadPreparedDocument(completeAfterDownload = true) {
    if (!preparedDocument) return;
    const anchor = document.createElement("a");
    const downloadUrl = URL.createObjectURL(preparedDocument.blob);
    anchor.href = downloadUrl;
    anchor.download = preparedDocument.fileName;
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 60_000);
    setExportedFileName(preparedDocument.fileName);
    setDownloadedCount(preparedDocument.updatedCount);
    if (completeAfterDownload) setView("complete");
  }

  function restart() {
    setView("upload");
    setFile(null);
    setFileName("");
    setExtraction(null);
    setDates([]);
    setProgress(0);
    setNotice("");
    setAssistantMessage("");
    setInterpretedInstruction("");
    setRefinement("");
    setPreparedDocument(null);
    setPreviewUrl("");
    setPreparingPreview(false);
    setExportError("");
    setExportedFileName("");
    setDownloadedCount(0);
    setFigureCalculation(null);
    setLedgerSource(null);
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#workspace" aria-label="DateShift home"><span className="brand-mark" aria-hidden="true"><i /></span><span>DateShift <em>AI</em></span></a>
        <nav aria-label="Primary navigation"><a className="active" href="#workspace">Workspace</a><a href="#how-it-works">How it works</a><a href="#privacy">Privacy</a></nav>
        <div className="header-actions"><span className="engine-status"><i /> Private mode ready</span><button className="avatar" aria-label="Open account menu">AM</button></div>
      </header>

      {view === "upload" && (
        <>
          <section className="hero" id="workspace">
            <div>
              <span className="eyebrow">PRIVATE DOCUMENT EDITOR</span>
              <h1>{task === "both" ? "Update dates and figures" : task === "ledger" ? "Transaction balances" : task === "dates" ? <>Move every date.<br /><span>Keep the document.</span></> : <>Balance every figure.<br /><span>Reach the exact total.</span></>}</h1>
              <p>{task === "both" ? "Shift dates and recalculate financial figures in a single Word or text-based PDF document." : task === "ledger" ? "Reconcile a Word transaction table, then model how an opening balance, debits, and credits reach your target closing balance." : task === "dates" ? "Upload a Word or PDF document. DateShift updates approved dates inside a new copy while retaining the original document’s remaining content and layout." : "Find financial figures, calculate an exact target total without rounding drift, and write the approved amounts into a new Word or PDF copy."}</p>
              <div className="task-switcher" role="group" aria-label="Document editing feature">
                <button aria-pressed={task === "dates"} className={task === "dates" ? "active" : ""} onClick={() => switchTask("dates")}><span aria-hidden="true">◷</span><b>Update dates</b><small>Shift calendar dates</small></button>
                <button aria-pressed={task === "figures"} className={task === "figures" ? "active" : ""} onClick={() => switchTask("figures")}><span aria-hidden="true">Σ</span><b>Balance figures</b><small>Reach an exact total</small></button>
                <button aria-pressed={task === "both"} className={task === "both" ? "active" : ""} onClick={() => switchTask("both")}><span aria-hidden="true">+</span><b>Dates &amp; figures</b><small>One updated file</small></button>
                <button aria-pressed={task === "ledger"} className={task === "ledger" ? "active" : ""} onClick={() => switchTask("ledger")}><span aria-hidden="true">±</span><b>Transactions</b><small>Calculate running balances</small></button>
              </div>
            </div>
            <div className="trust-card" id="privacy"><span className="shield" aria-hidden="true">✓</span><div><strong>Private on-device mode</strong><small>Your document stays in this browser and is never uploaded to our server.</small></div></div>
          </section>

          {task === "ledger" || isCombinedLedger ? <LedgerWorkspace key={task} initialSource={ledgerSource} onBackToSettings={isCombinedLedger ? () => setLedgerSource(null) : undefined} /> : <>
          <section className="workspace-grid">
            <article className="panel upload-panel">
              <div className="panel-heading"><span className="step">01</span><div><h2>Add your document</h2><p>DOCX or text-based PDF, up to 25 MB</p></div></div>
              <button className={`dropzone ${fileName ? "has-file" : ""}`} onClick={() => inputRef.current?.click()} onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); const dropped = event.dataTransfer.files[0]; if (dropped) acceptFile(dropped); }}>
                <input ref={inputRef} type="file" accept=".docx,.pdf" onChange={(event) => { const chosen = event.target.files?.[0]; if (chosen) acceptFile(chosen); }} />
                <span className="file-icon" aria-hidden="true"><i>{fileName ? "✓" : "+"}</i></span>
                {fileName ? <><strong>{fileName}</strong><small>{file ? `${Math.max(.1, file.size / 1024 / 1024).toFixed(1)} MB · Ready for private analysis` : "Sample document · Ready for analysis"}</small><span className="replace-file">Choose a different file</span></> : <><strong>Drop your document here</strong><small>or <u>browse your files</u></small><span className="format-pills"><i>DOCX</i><i>PDF</i></span></>}
              </button>
              {notice && <p className="notice" role="alert">{notice}</p>}
              {!fileName && <button className="demo-button" onClick={useDemo}><span>✦</span> Try with a sample document</button>}

              {includesDates && (
                <div className="compact-duration">
                  <div className="direction-tabs" aria-label="Default date shift direction"><button className={direction === "forward" ? "active" : ""} onClick={() => setDirection("forward")}>Move forward</button><button className={direction === "backward" ? "active" : ""} onClick={() => setDirection("backward")}>Move backward</button></div>
                  <div className="month-stepper"><button onClick={() => setMonths(Math.max(1, months - 1))} aria-label="Decrease months">−</button><div><strong>{months}</strong><span>month{months === 1 ? "" : "s"} {direction}</span></div><button onClick={() => setMonths(Math.min(24, months + 1))} aria-label="Increase months">+</button></div>
                  <div className="example-shift"><span>DEFAULT EXAMPLE</span><p><s>12 March 2026</s><b>→</b><strong>{exampleDate}</strong></p></div>
                </div>
              )}
              {includesFigures && (
                <div className="figure-controls">
                  <label><span>{task === "both" ? "Target total or ledger closing balance" : "Target total"}</span><input value={targetTotal} onChange={(event) => setTargetTotal(event.target.value)} placeholder="GHS 50,000.00" inputMode="decimal" /></label>
                  <label><span>Adjustment method</span><select value={figureStrategy} onChange={(event) => setFigureStrategy(event.target.value as FigureStrategy)}><option value="proportional">Adjust proportionally</option><option value="equal">Share difference equally</option><option value="last">Change final figure only</option></select></label>
                  <div className="calculation-note"><span>✓</span><p><strong>Exact-total calculation</strong><small>Uses integer minor units so rounding cannot change the requested total.</small></p></div>
                </div>
              )}
            </article>

            <article className="panel chat-panel">
              <div className="panel-heading"><span className="step orange">02</span><div><h2>Tell DateShift what to change</h2><p>Common instructions work in your own words</p></div></div>
              <div className="chat-message assistant-message"><span className="ai-avatar">DS</span><p>{task === "both" ? "Set the date shift and target amount, or include both in your instruction. If your Word file has a transaction table, the target becomes its closing balance." : task === "dates" ? "I recognize common numeric, written-month, weekday, timestamp, and mixed-separator date styles—including 01-Dec-2025. I also understand forward or backward month shifts, exclusions, and month-end preservation." : "Tell me the exact total you need. I can detect currency amounts and financial figures, exclude dates and likely IDs, then adjust the selected values proportionally, equally, or by changing only the final figure."}</p></div>
              {task === "figures" && <button className="ledger-mode-link" onClick={() => switchTask("ledger")}>Working with debits, credits, and running balances? Open Transactions.</button>}
              <label className="instruction-compose" htmlFor="instruction">
                <span>Your instruction</span>
                <textarea id="instruction" maxLength={600} value={instruction} onChange={(event) => setInstruction(event.target.value)} placeholder={task === "both" ? "Move dates forward 3 months and make the cash figures total GHS 50,000.00." : task === "dates" ? fallbackInstruction : "Make all cash figures total GHS 50,000.00 proportionally."} rows={5} />
                <small>{instruction.length}/600</small>
              </label>
              <div className="quick-prompts" aria-label="Suggested instructions">
                {includesDates && <><button onClick={() => addPrompt(`Move all dates ${direction} ${months} months.`)}>All dates</button><button onClick={() => addPrompt("Leave signature dates unchanged.")}>Skip signatures</button><button onClick={() => addPrompt("Keep month-end dates at the end of the new month.")}>Preserve month-end</button></>}
                {includesFigures && <><button onClick={() => addPrompt("Adjust all cash figures proportionally.")}>Proportional</button><button onClick={() => addPrompt("Share the difference equally.")}>Equal difference</button><button onClick={() => addPrompt("Change only the final figure.")}>Final figure only</button></>}
              </div>
              <button className="primary-button" disabled={!canAnalyze} onClick={analyze}>{task === "both" ? "Review dates and figures" : task === "dates" ? "Find & update dates" : "Calculate exact figures"}</button>
            </article>
          </section>

          <section className="preview-strip" id="how-it-works">
            <div className="preview-heading"><span className="spark">✦</span><div><strong>Download an updated document</strong><small>Preview date and figure changes in the same DOCX or text-based PDF you’ll download.</small></div></div>
            <div className="process-steps"><span><i>1</i> Upload</span><b>→</b><span><i>2</i> Instruct</span><b>→</b><span><i>3</i> Review</span></div>
          </section>
          </>}
        </>
      )}

      {view === "processing" && (
        <section className="focus-view processing-view" aria-live="polite">
          <div className="processing-orbit"><span className="document-glyph">DS</span><i /><i /><i /></div>
          <span className="eyebrow">READING YOUR DOCUMENT</span>
          <h2>{task === "both" ? "Finding dates and calculating figures" : task === "dates" ? "Finding dates in context" : "Calculating exact figures"}</h2>
          <p>{task === "both" ? "DateShift is preparing the date shift and checking which financial amounts can be recalculated." : task === "dates" ? "DateShift is extracting the document text locally, interpreting your instruction, and preparing an editable result." : "DateShift is identifying financial figures, excluding date values, and balancing the selected amounts to your exact target."}</p>
          <div className="progress-track"><span style={{ width: `${progress}%` }} /></div>
          <div className="progress-meta"><span>{progress < 50 ? "Reading document content" : progress < 90 ? "Applying your instruction" : "Preparing your review"}</span><strong>{progress}%</strong></div>
          <div className="file-chip"><span>DS</span><div><strong>{fileName}</strong><small>Private browser processing</small></div></div>
        </section>
      )}

      {view === "review" && (
        <section className="review-view" id="workspace">
          <div className="review-topline"><button className="back-button" onClick={() => setView("upload")}>← Back</button><span className="eyebrow">{task === "both" ? "DATE & FIGURE RESULTS" : task === "dates" ? "DATE RESULT" : "FIGURE RESULT"}</span><span className="secure-badge">✓ Private browser processing</span></div>
          <div className="review-heading"><div><h1>Review your <span>{task === "both" ? "changes." : task === "dates" ? "date updates." : "calculated figures."}</span></h1><p>{dates.length ? `${selectedCount} of ${dates.length} proposed ${resultType} selected in ${fileName}.` : `No ${resultType} matched your instruction in ${fileName}.`}</p></div><div className="shift-summary"><small>{task === "dates" ? "DEFAULT ADJUSTMENT" : "FIGURE TOTAL"}</small><strong>{task === "dates" ? <>{direction === "forward" ? "+" : "−"}{months} <span>month{months === 1 ? "" : "s"}</span></> : <>{(selectedFigureCount ? figureCalculation?.targetDisplay : figureCalculation?.currentDisplay) ?? "—"}<span>{selectedFigureCount && figureCalculation ? `from ${figureCalculation.currentDisplay}` : "source total"}</span></>}</strong></div></div>

          <div className="ai-summary"><span className="ai-avatar">DS</span><div><strong>{assistantMessage || "Analysis complete."}</strong><p><b>I understood:</b> {interpretedInstruction || effectiveInstruction}</p></div></div>

          {dates.length ? (
            <>
              <section className="document-preview-card" aria-label="Updated document preview">
                <div className="document-preview-header">
                  <div><span className="preview-file-type">{extraction?.kind === "pdf" ? "PDF" : "DOCX"}</span><div><strong>Updated document preview</strong><small>This is generated from the same file that will download.</small></div></div>
                  <span className={`preview-status ${preparedDocument ? "ready" : ""}`}>{preparingPreview ? "Preparing…" : preparedDocument ? `✓ ${preparedDocument.updatedCount} change${preparedDocument.updatedCount === 1 ? "" : "s"} verified` : selectedCount ? "Preview unavailable" : "Select an update"}</span>
                </div>
                <div className="document-preview-body">
                  {preparingPreview ? (
                    <div className="preview-loading"><i /><i /><i /><p>Writing the selected {resultType} into a private preview…</p></div>
                  ) : preparedDocument && extraction?.kind === "pdf" && previewUrl ? (
                    <iframe className="pdf-preview" src={`${previewUrl}#toolbar=0&navpanes=0&view=FitH`} title="Updated PDF preview" />
                  ) : preparedDocument && extraction?.kind === "docx" ? (
                    <DocxDocumentPreview blob={preparedDocument.blob} />
                  ) : (
                    <div className="preview-empty"><span>⌕</span><p>{selectedCount ? "The preview could not be prepared. Review the message beside the download button." : "Select at least one update to prepare the preview."}</p></div>
                  )}
                </div>
                <div className="document-preview-footer"><span><i /> Rendered from the exact prepared file used by Download</span><small>{fileName}</small></div>
              </section>

              <div className="review-layout">
              <article className="review-list-card">
                <div className="review-list-header"><div><strong>{task === "both" ? "Date and figure updates" : task === "dates" ? "Detected date updates" : "Calculated figure updates"}</strong><span>{selectedCount} of {dates.length} selected</span></div><button onClick={toggleAllSelections}>{selectedCount === dates.length ? "Clear all" : "Select all"}</button></div>
                <div className="date-list">
                  {dates.map((date) => (
                    <div className={`date-row ${date.selected ? "selected" : ""}`} key={date.id}>
                      <label className="check-control" aria-label={`Include ${date.original}`}><input type="checkbox" checked={date.selected} onChange={(event) => updateSelection(date.id, event.target.checked)} /><span>✓</span></label>
                      <div className="date-detail"><div className="context-line"><span>{date.context}</span><b>{task === "both" ? isFigureChange(date) ? "Figure" : "Date" : `${date.confidence}% match`}</b></div><div className="date-change"><s>{date.original}</s><i>→</i><input maxLength={80} readOnly={isFigureChange(date)} aria-label={`Updated value for ${date.original}`} value={date.updated} onChange={(event) => updateDate(date.id, event.target.value)} /></div><small className="date-reason">{date.reason}</small></div>
                    </div>
                  ))}
                </div>
              </article>
              <aside className="export-card">
                <div className="refine-chat"><span className="ai-avatar">DS</span><div><strong>Refine the result</strong><small>Add another instruction</small></div></div>
                <textarea value={refinement} onChange={(event) => setRefinement(event.target.value)} placeholder={task === "both" ? "e.g. Keep the signing date unchanged and use GHS 75,000" : task === "dates" ? "e.g. Keep the signing date unchanged" : "e.g. Use GHS 75,000 and change only the final figure"} rows={3} />
                <button className="secondary-button refine-button" disabled={!refinement.trim() || refining} onClick={() => { const combined = `${effectiveInstruction} ${refinement.trim()}`; setRefinement(""); void requestAnalysis(combined, false); }}>{refining ? "Updating…" : "Apply instruction"}</button>
                <dl><div><dt>Document</dt><dd>{fileName}</dd></div>{includesDates && <div><dt>Dates selected</dt><dd>{selectedCount - selectedFigureCount}</dd></div>}{includesFigures && <div><dt>Figures selected</dt><dd>{selectedFigureCount}</dd></div>}{figureCalculation && selectedFigureCount > 0 && <div><dt>Calculated total</dt><dd>{figureCalculation.targetDisplay}</dd></div>}<div><dt>Output check</dt><dd>{preparedDocument ? "Verified" : preparingPreview ? "Preparing" : "Not ready"}</dd></div><div><dt>Processing</dt><dd>Private</dd></div></dl>
                {exportError && <p className="notice export-error" role="alert">{exportError}</p>}
                <button className="primary-button" disabled={!selectedCount || preparingPreview || !preparedDocument} onClick={() => downloadPreparedDocument()}>{preparingPreview ? "Preparing verified preview…" : `Download updated ${extraction?.kind === "pdf" ? "PDF" : "Word file"}`} <span>↓</span></button>
                <small className="export-note">The download button uses the exact file shown in the preview above.</small>
              </aside>
              </div>
            </>
          ) : (
            <div className="empty-state"><span>⌕</span><h3>No matching {resultType}</h3><p>Change the instruction or choose another document.</p><div><button className="secondary-button" onClick={restart}>Choose another file</button><button className="primary-button" onClick={() => setView("upload")}>Change instruction</button></div></div>
          )}
        </section>
      )}

      {view === "complete" && (
        <section className="focus-view complete-view">
          <div className="success-ring">✓</div><span className="eyebrow">UPDATED DOCUMENT DOWNLOADED</span><h2>Your updated document is ready.</h2><p><strong>{exportedFileName || fileName}</strong> contains {downloadedCount} approved update{downloadedCount === 1 ? "" : "s"}.</p>
          {exportError && <p className="notice" role="alert">{exportError}</p>}
          <div className="completion-summary"><div><small>DOCUMENT FORMAT</small><strong>{extraction?.kind === "pdf" ? "PDF" : "Word (.docx)"}</strong></div><div><small>UPDATES APPLIED</small><strong>{downloadedCount}</strong></div><div><small>FILE STATUS</small><strong className="green">Downloaded</strong></div></div>
          <div className="completion-actions"><button className="primary-button" disabled={!preparedDocument} onClick={() => downloadPreparedDocument(false)}>Download updated file again <span>↓</span></button><button className="secondary-button new-document" onClick={restart}>Process another document <span>→</span></button></div>
        </section>
      )}
    </main>
  );
}
