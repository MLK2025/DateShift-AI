export type LedgerTransaction = {
  date: string;
  description: string;
  debit: bigint;
  credit: bigint;
  balance: bigint | null;
};

export type Ledger = {
  currency: string;
  opening: bigint;
  reportedClosing: bigint | null;
  transactions: LedgerTransaction[];
};

export type LedgerMethod = "spread" | "last";

const zero = BigInt(0);
const one = BigInt(1);
const minorUnits = BigInt(100);

export function parseMoney(value: string, currency = "") {
  const normalized = value.trim().replace(/\u2212/g, "-").replace(/\u00a0/g, " ");
  const match = normalized.match(/^(-|\()?\s*([₵$€£]|[A-Z]{3})?\s*(-)?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?)\s*(\))?$/i);
  if (!match || (match[1] === "(") !== Boolean(match[5]) || (match[1] && match[3])) return null;
  const symbol = match[2]?.toUpperCase();
  const code = symbol ? ({ "₵": "GHS", "$": "USD", "€": "EUR", "£": "GBP" }[symbol] ?? symbol) : "";
  if (currency && code && currency !== code) return null;
  const [whole, fractional = ""] = match[4].replace(/,/g, "").split(".");
  if (whole.length > 18) return null;
  const cents = BigInt(whole) * minorUnits + BigInt(fractional.padEnd(2, "0"));
  return match[1] || match[3] ? -cents : cents;
}

export function formatMoney(amount: bigint) {
  const absolute = amount < zero ? -amount : amount;
  const whole = String(absolute / minorUnits).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return `${amount < zero ? "-" : ""}${whole}.${String(absolute % minorUnits).padStart(2, "0")}`;
}

export function reconcileLedger(ledger: Ledger) {
  let balance = ledger.opening;
  let debit = zero;
  let credit = zero;
  let mismatches = 0;
  for (const transaction of ledger.transactions) {
    debit += transaction.debit;
    credit += transaction.credit;
    balance += transaction.credit - transaction.debit;
    if (transaction.balance !== null && transaction.balance !== balance) mismatches++;
  }
  return { debit, credit, closing: balance, mismatches };
}

function distributeAdjustment(weights: bigint[], difference: bigint) {
  if (difference === zero) return weights.map(() => zero);
  const direction = difference < zero ? -one : one;
  const magnitude = difference * direction;
  const totalWeight = weights.reduce((sum, weight) => sum + weight, zero);
  const denominator = totalWeight || BigInt(weights.length);
  const parts = weights.map((weight, index) => {
    const numerator = (totalWeight ? weight : one) * magnitude;
    return { index, units: numerator / denominator, remainder: numerator % denominator };
  });
  const remaining = Number(magnitude - parts.reduce((sum, part) => sum + part.units, zero));
  const byRemainder = [...parts].sort((a, b) => a.remainder === b.remainder ? a.index - b.index : a.remainder > b.remainder ? -1 : 1);
  for (const part of byRemainder.slice(0, remaining)) part.units++;
  return parts.map((part) => part.units * direction);
}

export function projectLedger(ledger: Ledger, target: bigint, specifiedOpening: bigint | null, method: LedgerMethod) {
  if (!ledger.transactions.length) throw new Error("Add at least one transaction to calculate a projection.");
  const source = reconcileLedger(ledger);
  const opening = specifiedOpening ?? target - (source.credit - source.debit);
  const difference = target - (opening + source.credit - source.debit);
  const lastIndex = ledger.transactions.length - 1;
  const adjustments = method === "last"
    ? ledger.transactions.map((_, index) => index === lastIndex ? difference : zero)
    : distributeAdjustment(ledger.transactions.map((row) => row.debit + row.credit), difference);

  let balance = opening;
  let debitTotal = zero;
  let creditTotal = zero;
  const transactions = ledger.transactions.map((row, index) => {
    const adjustment = adjustments[index];
    // When reducing a transaction, retain its direction until it reaches zero.
    // Any additional change becomes an amount in the other column.
    const debitReduction = adjustment > zero ? (row.debit < adjustment ? row.debit : adjustment) : zero;
    const creditReduction = adjustment < zero ? (row.credit < -adjustment ? row.credit : -adjustment) : zero;
    const debit = row.debit - debitReduction + (adjustment < zero ? -adjustment - creditReduction : zero);
    const credit = row.credit - creditReduction + (adjustment > zero ? adjustment - debitReduction : zero);
    balance += credit - debit;
    debitTotal += debit;
    creditTotal += credit;
    return { ...row, debit, credit, balance, adjustment };
  });
  if (balance !== target) throw new Error("The projection did not reach the requested closing balance.");
  return { opening, closing: balance, debitTotal, creditTotal, transactions, difference };
}

export type LedgerProjection = ReturnType<typeof projectLedger>;
