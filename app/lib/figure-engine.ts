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
  amountMinor: bigint;
};

export type FigureCalculation = {
  strategy: FigureStrategy;
  precision: number;
  currentUnits: bigint;
  targetUnits: bigint;
  currentDisplay: string;
  targetDisplay: string;
  summary: string;
};

const currencyCode = "(?:GHS|USD|EUR|GBP|NGN|CAD|AUD|ZAR|KES|JPY|CHF)";
const currencyToken = `(?:${currencyCode}|[$€£₵₦¥])`;
const digitsPattern = "(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d{1,4})?";
const amountLabel = `([-−]?)(${currencyToken}?[ \\t]*)(\\(?[-−]?)(${digitsPattern})(\\)?)([ \\t]*${currencyCode})?`;
const amountSource = `(?<![\\w.,/−-])(?:[-−]?${currencyToken}[ \\t]*)?\\(?[-−]?${digitsPattern}\\)?(?:[ \\t]*${currencyCode})?(?![\\w/−-]|[.,]\\d|[ \\t]*%)`;
const cashLabel = /\b(?:cash(?:\s+(?:transaction|allocation|withdrawal|deposit|amount|payment|received|paid))?|(?:transaction |payment |invoice )?amount|price|cost|payment|paid|fee|reserve|budget|revenue|expense|salary|income|debit|credit|deposit|withdrawal)\s*[:=]?\s*$/i;
const nonCashLabel = /\b(?:(?:account|acct\.?|reference|routing|transaction|customer|invoice|receipt|cheque|check|tax|order|serial|phone|policy|document|product|p\.o\.?)\s*(?:number|no\.?|id|code|ref(?:erence)?|#)|ref(?:erence)?|i\.?d\.?|account|routing|iban|swift|phone|qty|quantity|page|version|year|date|rate|percentage|percent|(?:sub)?total(?:\s+(?:debits?|credits?|cash|payments?))?|unit price|unit cost|denomination|balance(?:\s+(?:brought|carried) forward)?)\s*[:=#.]?\s*$/i;
const zero = BigInt(0);
const one = BigInt(1);
const ten = BigInt(10);

type ParsedAmount = {
  units: bigint;
  decimals: number;
  prefix: string;
  suffix: string;
  grouping: boolean;
  currency: string;
};

function contextAround(text: string, index: number, length: number) {
  const lineStart = Math.max(text.lastIndexOf("\n", index - 1) + 1, index - 80, 0);
  const nextLine = text.indexOf("\n", index + length);
  const lineEnd = nextLine >= 0 ? nextLine : text.length;
  return text.slice(lineStart, Math.min(lineEnd, index + length + 80)).replace(/\s+/g, " ").trim();
}

function parseAmountLabel(label: string): ParsedAmount | null {
  const amount = label.trim().match(new RegExp(`^${amountLabel}$`, "i"));
  if (!amount) return null;
  const [, outerSign, prefix, innerSign, digits, closing, suffix = ""] = amount;
  if (innerSign.startsWith("(") !== Boolean(closing) || (outerSign && innerSign) || innerSign.length > 1) return null;
  const [whole, decimalPart = ""] = digits.replace(/,/g, "").split(".");
  if (whole.length > 18) return null;
  const token = (prefix.trim() || suffix.trim()).toUpperCase();
  const currency = ({ "₵": "GHS", "$": "USD", "€": "EUR", "£": "GBP", "₦": "NGN", "¥": "JPY" } as Record<string, string>)[token] ?? token;
  return {
    units: BigInt(whole + decimalPart) * (outerSign || innerSign ? -one : one),
    decimals: decimalPart.length,
    prefix,
    suffix,
    grouping: digits.includes(","),
    currency,
  };
}

function formatMinorLike(units: bigint, precision: number, template: string) {
  const parsed = parseAmountLabel(template);
  const absolute = units < zero ? -units : units;
  const figures = String(absolute).padStart(precision + 1, "0");
  const whole = precision ? figures.slice(0, -precision) : figures;
  const decimal = precision ? `.${figures.slice(-precision)}` : "";
  let number = whole + decimal;
  if (parsed?.grouping) number = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + decimal;
  if (units < zero) number = `(${number})`;
  return `${parsed?.prefix ?? ""}${number}${parsed?.suffix ?? ""}`;
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

function sumUnits(values: bigint[]) {
  return values.reduce((sum, value) => sum + value, zero);
}

function rebalanceUnits(values: bigint[], target: bigint, strategy: FigureStrategy) {
  if (!values.length) return [];
  const current = sumUnits(values);
  if (strategy === "last") {
    const result = [...values];
    result[result.length - 1] += target - current;
    return result;
  }
  if (strategy === "equal" || current === zero) {
    const difference = target - current;
    const count = BigInt(values.length);
    const base = difference / count;
    let remainder = difference % count;
    return values.map((value) => {
      const extra = remainder === zero ? zero : remainder > zero ? one : -one;
      remainder -= extra;
      return value + base + extra;
    });
  }

  const denominator = current < zero ? -current : current;
  const numerators = values.map((value) => value * target * (current < zero ? -one : one));
  const result = numerators.map((numerator) => numerator / denominator);
  const remainder = target - sumUnits(result);
  const step = remainder > zero ? one : -one;
  const order = numerators.map((numerator, index) => ({ index, weight: (numerator % denominator) * step }))
    .sort((first, second) => first.weight === second.weight ? first.index - second.index : first.weight > second.weight ? -1 : 1);
  for (const { index } of order.slice(0, Number(remainder * step))) result[index] += step;
  return result;
}

function strategyFromInstruction(instruction: string, fallback: FigureStrategy) {
  const match = [...instruction.toLowerCase().matchAll(/\b(?:(?:last|final)\s+(?:figure|amount|transaction)|total\s+(?:line|figure)\s+only|equal(?:ly)?|evenly|same amount|proportion(?:al|ally)?)\b/g)].at(-1)?.[0];
  if (match?.startsWith("proportion")) return "proportional";
  if (match && /^(?:last|final|total)/.test(match)) return "last";
  if (match) return "equal";
  return fallback;
}

function strategyLabel(strategy: FigureStrategy) {
  if (strategy === "last") return "by changing only the final selected figure";
  if (strategy === "equal") return "by sharing the difference equally";
  return "by adjusting the selected figures proportionally";
}

export function rebalanceFigureChanges(changes: FigureChange[], calculation: FigureCalculation) {
  const fixedUnits = sumUnits(changes.filter((change) => !change.selected).map((change) => change.amountMinor));
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

export function readFigureTarget(instruction: string, fallbackTarget: string) {
  const instructionTarget = Array.from(instruction.matchAll(new RegExp(`\\b(?:total(?:\\s+of|\\s+to)?|target(?:\\s+of)?|reach|to|use)\\s*[:=]?\\s+(${amountSource})`, "gi"))).at(-1)?.[1];
  const instructionParsed = instructionTarget ? parseAmountLabel(instructionTarget) : null;
  const fallbackParsed = parseAmountLabel(fallbackTarget.trim());
  return instructionParsed && instructionTarget ? { label: instructionTarget, parsed: instructionParsed } : fallbackParsed ? { label: fallbackTarget.trim(), parsed: fallbackParsed } : null;
}

export function findFigures(text: string, instruction: string, fallbackTarget: string, fallbackStrategy: FigureStrategy) {
  const target = readFigureTarget(instruction, fallbackTarget);
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
    const lineStart = text.lastIndexOf("\n", sourceIndex - 1) + 1;
    const preceding = text.slice(Math.max(lineStart, sourceIndex - 70), sourceIndex);
    if (nonCashLabel.test(preceding) || (!parsed.currency && !cashLabel.test(preceding))) return [];
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

  const currencies = new Set([target.parsed.currency, ...candidates.map((candidate) => candidate.parsed.currency)].filter(Boolean));
  if (currencies.size > 1) throw new Error("The figures include different currencies. Use a single currency for each target total.");
  const precision = candidates.reduce((maximum, candidate) => Math.max(maximum, candidate.parsed.decimals), target.parsed.decimals);
  const toUnits = (parsed: ParsedAmount) => parsed.units * ten ** BigInt(precision - parsed.decimals);
  const targetUnits = toUnits(target.parsed);
  const rawChanges: FigureChange[] = candidates.map((candidate, index) => ({
    id: index + 1,
    sourceIndex: candidate.sourceIndex,
    original: candidate.original,
    updated: candidate.original,
    context: candidate.context,
    reason: "Ready for exact recalculation.",
    confidence: candidate.parsed.currency ? 98 : 90,
    selected: true,
    amountMinor: toUnits(candidate.parsed),
  }));
  const currentUnits = sumUnits(rawChanges.map((change) => change.amountMinor));
  const strategy = strategyFromInstruction(instruction, fallbackStrategy);
  const displayTemplate = target.parsed.currency ? target.label : candidates.find((candidate) => candidate.parsed.currency)?.original ?? target.label;
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
