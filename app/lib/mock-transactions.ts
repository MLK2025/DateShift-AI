export type MockPattern = "salary" | "non-salary";
export type MockFlow = "credit" | "debit";
export type MockTransaction = {
  date: string;
  description: string;
  reference: string;
  category: string;
  flow: MockFlow;
  minorUnits: number;
};

const generalSamples: { description: string; category: string; flow: MockFlow }[] = [
  { description: "Office supply order", category: "Purchase", flow: "debit" },
  { description: "Weekend kiosk order", category: "Purchase", flow: "debit" },
  { description: "Sample inventory purchase", category: "Purchase", flow: "debit" },
  { description: "Equipment order return", category: "Refund", flow: "credit" },
  { description: "Returned test shipment", category: "Refund", flow: "credit" },
  { description: "Subscription fee", category: "Fee", flow: "debit" },
  { description: "Workspace service charge", category: "Fee", flow: "debit" },
  { description: "Training allowance transfer", category: "Transfer", flow: "credit" },
  { description: "Internal test wallet transfer", category: "Transfer", flow: "debit" },
  { description: "Example invoice settlement", category: "Transfer", flow: "credit" },
];

const payrollDescriptions = [
  "Month-end salary credit",
  "Monthly payroll deposit",
  "Payroll settlement credit",
  "End-of-month pay run",
  "Regular salary transfer",
  "Monthly wages receipt",
];
const payrollReferenceCodes = ["PAY", "WAGE", "SAL", "RUN", "HR", "CREDIT"];
const sampleDayRanges = [[2, 5], [8, 13], [16, 20], [23, 26], [27, 28]];

export function seededInteger(seed: number) {
  let state = seed >>> 0;
  return (limit: number) => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return Math.floor(state / 2 ** 32 * limit);
  };
}

export function generateMockTransactions(pattern: MockPattern, placeholder: string, batch: number, now = new Date()): MockTransaction[] {
  const company = placeholder.trim().slice(0, 64) || "Example Company";
  const variation = seededInteger(batch);
  const rows: MockTransaction[] = [];
  const payrollAmount = 245000 + variation(85000);
  const year = now.getUTCFullYear();
  const currentMonth = now.getUTCMonth();
  for (let offset = 6; offset > 0; offset -= 1) {
    const month = currentMonth - offset;
    const monthEnd = new Date(Date.UTC(year, month + 1, 0));
    const period = monthEnd.toISOString().slice(0, 7);
    const generalCount = pattern === "salary" ? 4 : 5;
    const generalStart = variation(generalSamples.length);
    for (let i = 0; i < generalCount; i += 1) {
      const [first, last] = sampleDayRanges[i];
      const day = first + variation(last - first + 1);
      const sample = generalSamples[(generalStart + i * 3) % generalSamples.length];
      rows.push({
        ...sample,
        date: `${period}-${String(day).padStart(2, "0")}`,
        description: `Mock ${company} · ${sample.description} · ${period}`,
        reference: `MOCK-TXN-${batch}-${period}-${i + 1}-${100 + variation(900)}`,
        minorUnits: 125 + variation(64976),
      });
    }
    if (pattern === "salary") {
      const style = (6 - offset + batch % payrollDescriptions.length) % payrollDescriptions.length;
      rows.push({
        date: monthEnd.toISOString().slice(0, 10),
        description: `Mock ${company} · ${payrollDescriptions[style]} · ${period}`,
        reference: `MOCK-${payrollReferenceCodes[style]}-${batch}-${period}-${100 + variation(900)}`,
        category: "Payroll",
        flow: "credit",
        minorUnits: payrollAmount,
      });
    }
  }
  return rows;
}

export function filterMockTransactions(
  rows: MockTransaction[],
  { search, month, type }: { search: string; month: string; type: "all" | MockFlow | "payroll" },
) {
  const query = search.trim().toLowerCase();
  return rows.filter((row) =>
    (!month || row.date.startsWith(month))
    && (type === "all" || (type === "payroll" ? row.category === "Payroll" : row.flow === type))
    && (!query || `${row.description} ${row.reference} ${row.category}`.toLowerCase().includes(query)),
  );
}
