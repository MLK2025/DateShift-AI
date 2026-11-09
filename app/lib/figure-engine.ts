export type FigureStrategy = "proportional" | "equal" | "last";

export type FigureChange = {
  id: number;
  sourceIndex: number;
  original: string;
  updated: string;
  context: string;
  reason: string;
  confidence: number;
  selected: boolean;
  amountMinor: number;
};

export type FigureCalculation = {
  strategy: FigureStrategy;
  precision: number;
  currentUnits: number;
  targetUnits: number;
  currentDisplay: string;
  targetDisplay: string;
  summary: string;
};

const currencyCode = "(?:GHS|USD|EUR|GBP|NGN|CAD|AUD|ZAR|KES|JPY|CHF)";
const currencyToken = `(?:${currencyCode}|[$€£₵₦¥])`;
const amountNumber = "\\(?-?(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d{1,4})?\\)?";
const amountSource = `(?<![\\w])(?:${currencyToken}\\s*)?${amountNumber}(?:\\s*${currencyCode})?(?![\\w%])`;
const moneyContext = /\b(?:cash|amount|price|cost|subtotal|total|balance|payment|fee|budget|revenue|expense|salary|income|figure|value|debit|credit|net|gross)\b/i;

type ParsedAmount = {
  value: number;
  decimals: number;
  prefix: string;
  suffix: string;
  grouping: boolean;
  parentheses: boolean;
  hasCurrency: boolean;
};

function contextAround(text: string, index: number, length: number) {
  const lineStart = Math.max(text.lastIndexOf("\n", index - 1) + 1, index - 80, 0);
  const nextLine = text.indexOf("\n", index + length);
  const lineEnd = nextLine >= 0 ? nextLine : text.length;
  return text.slice(lineStart, Math.min(lineEnd, index + length + 80)).replace(/\s+/g, " ").trim();
}

function parseAmountLabel(label: string): ParsedAmount | null {
  const number = label.match(/\(?-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,4})?\)?/);
  if (!number || number.index === undefined) return null;
  const raw = number[0];
  const negative = raw.includes("-") || (raw.startsWith("(") && raw.endsWith(")"));
  const normalized = raw.replace(/[(),-]/g, "");
  const value = Number(normalized);
  if (!Number.isFinite(value) || value > 1_000_000_000_000) return null;
  const decimalPart = normalized.split(".")[1] ?? "";
  return {
    value: negative ? -value : value,
    decimals: decimalPart.length,
    prefix: label.slice(0, number.index),
    suffix: label.slice(number.index + raw.length),
    grouping: raw.includes(","),
    parentheses: raw.startsWith("(") && raw.endsWith(")"),
    hasCurrency: new RegExp(currencyToken, "i").test(label),
  };
}

function formatMinorLike(units: number, precision: number, template: string) {
  const parsed = parseAmountLabel(template);
  if (!parsed) return String(units / 10 ** precision);
  const scale = 10 ** precision;
  const absolute = Math.abs(units) / scale;
  let number = absolute.toFixed(precision);
  if (parsed.grouping) {
    const [whole, decimal] = number.split(".");
    number = `${whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",")}${decimal === undefined ? "" : `.${decimal}`}`;
  }
  if (units < 0) number = parsed.parentheses ? `(${number})` : `-${number}`;
  return `${parsed.prefix}${number}${parsed.suffix}`;
}

function dateRanges(text: string) {
  const month = "(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)";
  const weekday = "(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun)";
  const separator = "[\\s\\u200B\\u2060]*[-‐‑‒–—−/.][\\s\\u200B\\u2060]*";
  const patterns = [
    new RegExp(`\\b${weekday},?\\s+(?:${month}\\s+\\d{1,2}|\\d{1,2}\\s+${month})[^\\n]{0,24}?\\d{4}\\b`, "gi"),
    new RegExp(`\\b(?:\\d{1,2}${separator}${month}|${month}${separator}\\d{1,2}|\\d{4}${separator}${month})${separator}\\d{1,4}\\b`, "gi"),
    new RegExp(`\\b(?:\\d{4}${separator}\\d{1,2}${separator}\\d{1,2}|\\d{1,2}${separator}\\d{1,2}${separator}\\d{4})\\b`, "g"),
    new RegExp(`\\b(?:${month}\\.?\\s+\\d{1,2}|\\d{1,2}(?:st|nd|rd|th)?\\s+${month})[,]?\\s+\\d{4}\\b`, "gi"),
  ];
  return patterns.flatMap((pattern) => [...text.matchAll(pattern)].map((match) => ({ start: match.index ?? 0, end: (match.index ?? 0) + match[0].length })));
}

function rebalanceUnits(values: number[], target: number, strategy: FigureStrategy) {
  if (!values.length) return [];
  const current = values.reduce((sum, value) => sum + value, 0);
  if (strategy === "last") {
    const result = [...values];
    result[result.length - 1] += target - current;
    return result;
  }
  if (strategy === "equal" || current === 0) {
    const difference = target - current;
    const base = Math.trunc(difference / values.length);
    let remainder = difference - base * values.length;
    return values.map((value) => {
      const extra = remainder === 0 ? 0 : remainder > 0 ? 1 : -1;
      remainder -= extra;
      return value + base + extra;
    });
  }

  const raw = values.map((value) => value * target / current);
  const result = raw.map((value) => Math.trunc(value));
  let remainder = target - result.reduce((sum, value) => sum + value, 0);
  const order = raw.map((value, index) => ({ index, fraction: value - result[index] }))
    .sort((first, second) => remainder > 0 ? second.fraction - first.fraction : first.fraction - second.fraction);
  let cursor = 0;
  while (remainder !== 0) {
    const step = remainder > 0 ? 1 : -1;
    result[order[cursor % order.length].index] += step;
    remainder -= step;
    cursor += 1;
  }
  return result;
}

function strategyFromInstruction(instruction: string, fallback: FigureStrategy) {
  const normalized = instruction.toLowerCase();
  if (/\b(?:last|final|single|one)\b/.test(normalized) || /total\s+(?:line|figure)\s+only/.test(normalized)) return "last";
  if (/\b(?:equal|equally|evenly|same amount)\b/.test(normalized)) return "equal";
  if (/\bproportion(?:al|ally)?\b/.test(normalized)) return "proportional";
  return fallback;
}

function strategyLabel(strategy: FigureStrategy) {
  if (strategy === "last") return "by changing only the final selected figure";
  if (strategy === "equal") return "by sharing the difference equally";
  return "by adjusting the selected figures proportionally";
}

export function rebalanceFigureChanges(changes: FigureChange[], calculation: FigureCalculation) {
  const fixedUnits = changes.filter((change) => !change.selected).reduce((sum, change) => sum + change.amountMinor, 0);
  const selected = changes.filter((change) => change.selected);
  if (!selected.length) return changes.map((change) => ({ ...change, updated: change.original }));
  const balanced = rebalanceUnits(selected.map((change) => change.amountMinor), calculation.targetUnits - fixedUnits, calculation.strategy);
  let selectedIndex = 0;
  return changes.map((change) => {
    if (!change.selected) return { ...change, updated: change.original, reason: "Left unchanged; the remaining selected figures were recalculated to keep the requested total exact." };
    const updatedUnits = balanced[selectedIndex++];
    return {
      ...change,
      updated: formatMinorLike(updatedUnits, calculation.precision, change.original),
      reason: `Recalculated ${strategyLabel(calculation.strategy)} so the document figures total exactly ${calculation.targetDisplay}.`,
    };
  });
}

export function findFigures(text: string, instruction: string, fallbackTarget: string, fallbackStrategy: FigureStrategy) {
  const instructionAmounts = [...instruction.matchAll(new RegExp(amountSource, "gi"))]
    .map((match) => ({ label: match[0], parsed: parseAmountLabel(match[0]) }))
    .filter((item): item is { label: string; parsed: ParsedAmount } => Boolean(item.parsed));
  const fallbackParsed = parseAmountLabel(fallbackTarget.trim());
  const target = instructionAmounts.at(-1) ?? (fallbackParsed ? { label: fallbackTarget.trim(), parsed: fallbackParsed } : null);
  if (!target) throw new Error("Add the exact target total, for example: Make the cash figures total GHS 50,000.00.");

  const ranges = dateRanges(text);
  const candidates = [...text.matchAll(new RegExp(amountSource, "gi"))].flatMap((match) => {
    const original = match[0];
    const sourceIndex = match.index ?? 0;
    const end = sourceIndex + original.length;
    if (ranges.some((range) => sourceIndex < range.end && end > range.start)) return [];
    const parsed = parseAmountLabel(original);
    if (!parsed) return [];
    const context = contextAround(text, sourceIndex, original.length);
    const digits = original.replace(/\D/g, "");
    const plainInteger = !parsed.hasCurrency && !/[,.]/.test(original);
    if (plainInteger && parsed.value >= 1900 && parsed.value <= 2100) return [];
    if (plainInteger && (digits.length <= 2 || digits.length >= 10) && !moneyContext.test(context)) return [];
    return [{ original, sourceIndex, context, parsed }];
  });

  if (!candidates.length) {
    return {
      changes: [] as FigureChange[],
      calculation: null,
      assistantMessage: "I read the document but did not find editable cash amounts or financial figures. Plain reference numbers, dates, percentages, and likely IDs were excluded.",
      summary: `Target total: ${target.label}.`,
    };
  }

  const precision = Math.min(4, Math.max(target.parsed.decimals, ...candidates.map((candidate) => candidate.parsed.decimals)));
  const scale = 10 ** precision;
  const targetUnits = Math.round(target.parsed.value * scale);
  const rawChanges: FigureChange[] = candidates.map((candidate, index) => ({
    id: index + 1,
    sourceIndex: candidate.sourceIndex,
    original: candidate.original,
    updated: candidate.original,
    context: candidate.context,
    reason: "Ready for exact recalculation.",
    confidence: candidate.parsed.hasCurrency || /[,.]/.test(candidate.original) || moneyContext.test(candidate.context) ? 98 : 84,
    selected: true,
    amountMinor: Math.round(candidate.parsed.value * scale),
  }));
  const currentUnits = rawChanges.reduce((sum, change) => sum + change.amountMinor, 0);
  const strategy = strategyFromInstruction(instruction, fallbackStrategy);
  const displayTemplate = target.parsed.hasCurrency ? target.label : rawChanges.find((change) => parseAmountLabel(change.original)?.hasCurrency)?.original ?? target.label;
  const currentDisplay = formatMinorLike(currentUnits, precision, displayTemplate);
  const targetDisplay = formatMinorLike(targetUnits, precision, displayTemplate);
  const calculation: FigureCalculation = {
    strategy,
    precision,
    currentUnits,
    targetUnits,
    currentDisplay,
    targetDisplay,
    summary: `Make ${rawChanges.length} detected figure${rawChanges.length === 1 ? "" : "s"} total exactly ${targetDisplay} ${strategyLabel(strategy)}.`,
  };
  const changes = rebalanceFigureChanges(rawChanges, calculation);
  return {
    changes,
    calculation,
    assistantMessage: `I found ${changes.length} financial figure${changes.length === 1 ? "" : "s"}. Their current total is ${currentDisplay}; the proposed figures total exactly ${targetDisplay}.`,
    summary: calculation.summary,
  };
}
