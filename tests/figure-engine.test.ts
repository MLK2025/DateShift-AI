import assert from "node:assert/strict";
import test from "node:test";
import { findFigures, rebalanceFigureChanges } from "../app/lib/figure-engine";

function numericValue(label: string) {
  const match = label.match(/-?[\d,]+(?:\.\d+)?/);
  return Number((match?.[0] ?? "0").replace(/,/g, ""));
}

test("balances currency figures proportionally to an exact requested total", () => {
  const result = findFigures(
    "Statement date 01-Dec-2025\nCash allocation GHS 12,500.00\nOperating reserve GHS 7,500.00",
    "Make all cash figures total GHS 30,000.00 proportionally.",
    "",
    "proportional",
  );

  assert.equal(result.changes.length, 2);
  assert.deepEqual(result.changes.map((change) => change.original), ["GHS 12,500.00", "GHS 7,500.00"]);
  assert.deepEqual(result.changes.map((change) => change.updated), ["GHS 18,750.00", "GHS 11,250.00"]);
  assert.equal(result.changes.reduce((sum, change) => sum + numericValue(change.updated), 0), 30_000);
  assert.equal(result.calculation?.currentDisplay, "GHS 20,000.00");
  assert.equal(result.calculation?.targetDisplay, "GHS 30,000.00");
});

test("uses integer minor units so fractional rounding still reaches the target", () => {
  const result = findFigures(
    "Cash $0.01\nCash $0.01\nCash $0.01",
    "Make the total $0.02 proportionally.",
    "",
    "proportional",
  );

  assert.deepEqual(result.changes.map((change) => change.updated), ["$0.01", "$0.01", "$0.00"]);
  assert.equal(result.changes.reduce((sum, change) => sum + numericValue(change.updated), 0), 0.02);
});

test("recalculates the remaining selected figures when one amount is excluded", () => {
  const result = findFigures(
    "Cash GHS 12,500.00\nReserve GHS 7,500.00",
    "Make the total GHS 30,000.00 proportionally.",
    "",
    "proportional",
  );
  assert.ok(result.calculation);
  const selected = result.changes.map((change, index) => ({ ...change, selected: index !== 0 }));
  const rebalanced = rebalanceFigureChanges(selected, result.calculation);

  assert.deepEqual(rebalanced.map((change) => change.updated), ["GHS 12,500.00", "GHS 17,500.00"]);
  assert.equal(rebalanced.reduce((sum, change) => sum + numericValue(change.updated), 0), 30_000);
});

test("can change only the final figure", () => {
  const result = findFigures(
    "Cash GHS 100.00\nReserve GHS 200.00",
    "Make the total GHS 400.00 by changing only the final figure.",
    "",
    "proportional",
  );

  assert.deepEqual(result.changes.map((change) => change.updated), ["GHS 100.00", "GHS 300.00"]);
  assert.equal(result.calculation?.strategy, "last");
});
