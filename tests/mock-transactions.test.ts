import assert from "node:assert/strict";
import test from "node:test";
import { filterMockTransactions, generateMockTransactions } from "../app/lib/mock-transactions";

const sampleDate = new Date("2028-03-08T12:00:00Z");

test("generates one distinct payroll credit on each completed calendar month end", () => {
  const rows = generateMockTransactions("salary", "Test Company", 1, sampleDate);
  const salary = filterMockTransactions(rows, { type: "payroll", month: "", search: "" });
  assert.equal(rows.length, 30);
  assert.deepEqual(salary.map((row) => row.date),
    ["2027-09-30", "2027-10-31", "2027-11-30", "2027-12-31", "2028-01-31", "2028-02-29"]);
  assert.equal(new Set(salary.map((row) => row.minorUnits)).size, 1);
  assert.equal(new Set(rows.map((row) => row.description)).size, rows.length);
  assert.equal(new Set(rows.map((row) => row.reference)).size, rows.length);
  assert.ok(salary.every((row) => row.flow === "credit"));
  assert.ok(rows.every((row) => row.reference.startsWith("MOCK-") && row.description.startsWith("Mock Test Company") && row.minorUnits > 0));
});

test("generates general activity and combines the text, month, and type filters", () => {
  const rows = generateMockTransactions("non-salary", " ", 4, sampleDate);
  assert.equal(rows.length, 30);
  assert.equal(new Set(rows.map((row) => row.date)).size, rows.length);
  assert.ok(rows.every((row) => row.category !== "Payroll" && row.description.startsWith("Mock Example Company")));
  assert.notDeepEqual(rows, generateMockTransactions("non-salary", " ", 5, sampleDate));

  const debit = rows.find((row) => row.flow === "debit")!;
  const filters = { type: "debit" as const, month: debit.date.slice(0, 7), search: ` ${debit.reference.toLowerCase()} ` };
  assert.deepEqual(filterMockTransactions(rows, filters), [debit]);
  assert.deepEqual(filterMockTransactions(rows, { ...filters, type: "payroll" }), []);
});
