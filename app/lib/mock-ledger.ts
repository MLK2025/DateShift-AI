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
const debits = ["Supply Order", "Service Fee", "Transit Fare", "Stock Purchase", "Test Charge", "Vendor Payment"];
const credits = ["Test Refund", "Transfer In", "Sample Rebate", "Order Reversal", "Wallet Transfer", "Sample Deposit"];

function styledDescription(original: string, phrase: string, reference: string) {
  const label = /[a-z]/i.test(original) && original === original.toUpperCase() ? phrase.toUpperCase() : phrase;
  if (original.includes(";")) return `${label};${reference}`;
  if (original.includes("_")) return `${reference}_${label}`;
  if (original.includes("/")) return `${label}/${reference}`;
  return `${label} ${reference}`;
}

export function generateMockLedger(ledger: Ledger, pattern: MockPattern, companyName: string, batch: number) {
  const company = companyName.trim().replace(/\s+/g, " ") || "Example Company";
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
  const batchCode = batch.toString(36).toUpperCase().slice(-2).padStart(2, "0");
  const transactions = rows.map((source, index) => {
    const reference = `SMP${batchCode}${String(index + 1).padStart(4, "0")}`;
    const shortRef = `SMP${String(index + 1).padStart(3, "0")}`;
    const dateStyle = parsedDates[index] ?? known[0];
    if (index < openingSlots) return {
      date: dateStyle.format(startDate), dateValue: startDate, description: styledDescription(source.description, "Opening Bal", shortRef),
      reference, debit: zero, credit: zero, balance, category: "Opening balance",
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
    const label = isPayroll ? `${payroll[(period + batch) % payroll.length]} ${company}` : (isCredit ? credits : debits)[random(debits.length)];
    const debit = isCredit ? zero : amount;
    const credit = isCredit ? amount : zero;
    balance += credit - debit;
    return {
      date: dateStyle.format(date), dateValue: date, description: styledDescription(source.description, label, shortRef),
      reference, debit, credit, balance, category: isPayroll ? "Payroll" : isCredit ? "Credit" : "Debit",
    };
  });
  const totals = reconcileLedger({ ...ledger, opening, transactions });
  return { opening, closing: totals.closing, debitTotal: totals.debit, creditTotal: totals.credit, transactions, startDate, endDate };
}
