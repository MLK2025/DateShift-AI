import { reconcileLedger, type Ledger } from "./ledger";
import { seededInteger, type MockPattern } from "./mock-transactions";

const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const monthShort = monthNames.map((month) => month.slice(0, 3));
const twoDigits = (n: number) => String(n).padStart(2, "0");
const endOfMonth = (period: number) => new Date(Date.UTC(Math.floor(period / 12), period % 12 + 1, 0));
const periodOf = (date: Date) => date.getUTCFullYear() * 12 + date.getUTCMonth();

function matchCapitalization(text: string, example: string) {
  if (example === example.toUpperCase()) return text.toUpperCase();
  if (example === example.toLowerCase()) return text.toLowerCase();
  return text;
}

export function templateDate(text: string): { date: Date; format: (date: Date) => string } | null {
  const yearFirst = /^(\d{4})([-/.])(\d{1,2})\2(\d{1,2})$/.exec(text.trim());
  const dayFirst = /^(\d{1,2})([-/. ])([A-Za-z]{3,9}|\d{1,2})\2(\d{4})$/.exec(text.trim());
  if (!yearFirst && !dayFirst) return null;
  const year = Number(yearFirst?.[1] ?? dayFirst![4]);
  const monthToken = yearFirst?.[3] ?? dayFirst![3];
  const named = /^[A-Za-z]/.test(monthToken);
  const usOrder = Boolean(dayFirst && !named && Number(monthToken) > 12 && Number(dayFirst[1]) <= 12);
  const month = named ? monthShort.findIndex((value) => value.toLowerCase() === monthToken.slice(0, 3).toLowerCase())
    : Number(usOrder ? dayFirst![1] : monthToken) - 1;
  const day = Number(yearFirst?.[4] ?? (usOrder ? monthToken : dayFirst![1]));
  if (year < 1900 || month < 0 || month > 11 || day < 1 || day > 31) return null;
  const date = new Date(Date.UTC(year, month, Math.min(day, new Date(Date.UTC(year, month + 1, 0)).getUTCDate())));
  const separator = (yearFirst ?? dayFirst)![2];
  return {
    date,
    format(next) {
      const month = next.getUTCMonth();
      const label = named ? matchCapitalization((monthToken.length === 3 ? monthShort : monthNames)[month], monthToken) : twoDigits(month + 1);
      const day = twoDigits(next.getUTCDate());
      const year = next.getUTCFullYear();
      return (yearFirst ? [year, label, day] : usOrder ? [label, day, year] : [day, label, year]).join(separator);
    },
  };
}

const payroll = ["Payroll", "Salary Cr", "Pay Run", "Wages Cr", "Staff Pay", "Monthly Pay"];
type Narration = { channel: string; texts: string[] };
const debits: Narration[] = [
  { channel: "POS", texts: ["Card Purchase", "Retail Purchase", "Grocery Purchase", "Merchant Payment"] },
  { channel: "ATM", texts: ["Cash Withdrawal", "Terminal Withdrawal", "Cash Disbursement"] },
  { channel: "TRF", texts: ["Transfer Out", "Vendor Transfer", "Online Payment", "Utility Payment"] },
  { channel: "WLT", texts: ["Mobile Wallet Debit", "Wallet Transfer Out", "Mobile Bill Payment"] },
  { channel: "FEE", texts: ["Service Charge", "Transfer Commission", "Card Fee", "Account Fee"] },
];
const credits: Narration[] = [
  { channel: "TRF", texts: ["Transfer In", "Incoming Transfer", "Invoice Settlement", "Online Transfer Credit"] },
  { channel: "DEP", texts: ["Cash Deposit", "Counter Deposit", "Deposit Credit"] },
  { channel: "WLT", texts: ["Mobile Wallet Credit", "Wallet Transfer In", "Wallet Settlement"] },
  { channel: "REV", texts: ["Card Refund", "Purchase Reversal", "Fee Reversal", "Merchant Refund"] },
  { channel: "INT", texts: ["Interest Credit", "Interest Payment", "Account Interest"] },
];

function genericNarration(original: string, credit: boolean, random: (limit: number) => number) {
  const list = credit ? credits : debits;
  const channel = /\b(?:wallet|momo|mobile|mtn)\b|\bp2p/i.test(original) ? "WLT"
    : /\b(?:POS|VISA|card purchase)\b/i.test(original) ? "POS"
    : /\b(?:ATM|withdrawal)\b/i.test(original) ? "ATM"
    : /\b(?:commission|charge|fee)\b/i.test(original) ? "FEE"
    : /\b(?:GIP|ACH|EFT|transfer|TRF)\b/i.test(original) ? "TRF"
    : /\b(?:interest)\b/i.test(original) ? "INT" : undefined;
  const matching = random(3) === 0 ? undefined : list.find((entry) => entry.channel === channel);
  const pick = matching ?? list[random(list.length)];
  return { label: pick.texts[random(pick.texts.length)], channel: pick.channel };
}

function styledDescription(original: string, phrase: string, reference: string, channel: string, used: Set<string>) {
  const label = /[a-z]/i.test(original) ? matchCapitalization(phrase, original) : phrase;
  let description = label;
  if (original.includes(";")) description = original.split(";").length > 2
    ? `${label};${reference}; ${channel}` : `${label};${reference}`;
  else if (original.includes("_")) description = `${reference}_${label}`;
  else if (original.includes("/")) description = `${label}/${reference}`;
  else if (original.includes(":")) description = `${label} : ${reference}`;
  else if (/\d/.test(original) || used.has(label.toUpperCase())) description = `${label} ${reference}`;
  used.add(description.toUpperCase());
  return description;
}

export function generateMockLedger(ledger: Ledger, pattern: MockPattern, companyName: string, batch: number) {
  const company = companyName.replace(/\b(?:sample|smp)(?=\b|\d)/gi, "").trim().replace(/\s+/g, " ") || "Example Company";
  const rows = ledger.transactions;
  const parsedDates = rows.map((row) => templateDate(row.date));
  const known = parsedDates.flatMap((value) => value ? [value] : []);
  if (!known.length) throw new Error("The template needs recognizable transaction dates, such as 24-Sep-2026 or 2026-09-24.");
  const periods = known.map(({ date }) => periodOf(date));
  const first = periods.reduce((a, b) => Math.min(a, b));
  const last = periods.reduce((a, b) => Math.max(a, b));
  const periodCount = last - first + 1;
  const openingSlots = ledger.hasOpeningEntry ? 0 : 1;
  const movementCount = rows.length - openingSlots;
  if (movementCount < 1 || (pattern === "salary" && movementCount < periodCount)) {
    throw new Error(pattern === "salary" ? "The template needs at least one transaction row per month, plus any opening balance row." : "The template needs at least one transaction after its opening balance.");
  }
  const startDate = new Date(Date.UTC(Math.floor(first / 12), first % 12, 1));
  const endDate = endOfMonth(last);
  const slots = Array.from({ length: movementCount }, (_, index) => first + Math.floor(index * periodCount / movementCount));
  const monthlyCounts = new Map<number, number>();
  for (const period of slots) monthlyCounts.set(period, (monthlyCounts.get(period) ?? 0) + 1);
  const seen = new Map<number, number>();
  const random = seededInteger(batch + rows.length);
  const payrollAmount = BigInt(165000 + random(180000));
  const zero = BigInt(0);
  const opening = BigInt(220000 + random(90000));
  let balance = opening;
  const batchCode = String((batch % 99) + 1).padStart(2, "0");
  const descriptions = new Set<string>();
  const transactions = rows.map((source, index) => {
    const serial = `${batchCode}${String(index + 1).padStart(4, "0")}`;
    const dateStyle = parsedDates[index] ?? known[0];
    if (index < openingSlots) return {
      date: dateStyle.format(startDate), dateValue: startDate,
      description: styledDescription(source.description, "Opening Bal", serial, "BAL", descriptions),
      reference: `BAL${serial}`, debit: zero, credit: zero, balance, category: "Opening balance",
    };

    const period = slots[index - openingSlots];
    const position = seen.get(period) ?? 0;
    const count = monthlyCounts.get(period)!;
    seen.set(period, position + 1);
    const lastDay = endOfMonth(period).getUTCDate();
    const isPayroll = pattern === "salary" && position === count - 1;
    const day = isPayroll ? lastDay : 1 + Math.floor((position + .5) * (lastDay - 1) / count);
    const date = new Date(Date.UTC(Math.floor(period / 12), period % 12, day));
    const amount = isPayroll ? payrollAmount : BigInt(250 + random(22500));
    const isCredit = isPayroll || balance < amount || index % 5 === 0 || (source.credit > zero && index % 5 !== 1);
    const narration = isPayroll
      ? { label: `${payroll[(period + batch) % payroll.length]} ${company}`, channel: "PAY" }
      : genericNarration(source.description, isCredit, random);
    const debit = isCredit ? zero : amount;
    const credit = isCredit ? amount : zero;
    balance += credit - debit;
    return {
      date: dateStyle.format(date), dateValue: date,
      description: styledDescription(source.description, narration.label, serial, narration.channel, descriptions),
      reference: `${narration.channel}${serial}`, debit, credit, balance, category: isPayroll ? "Payroll" : isCredit ? "Credit" : "Debit",
    };
  });
  const totals = reconcileLedger({ ...ledger, opening, transactions });
  return { opening, closing: totals.closing, debitTotal: totals.debit, creditTotal: totals.credit, transactions, startDate, endDate };
}
