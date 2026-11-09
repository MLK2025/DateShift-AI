import assert from "node:assert/strict";
import test from "node:test";
import { findDates, interpretInstruction, shiftDateLabel } from "../app/page";

test("always treats hyphenated numeric dates as DD-MM-YYYY", () => {
  assert.equal(shiftDateLabel("20-12-2026", 3), "20-03-2027");
  assert.equal(shiftDateLabel("05-12-2026", 3), "05-03-2027");
  assert.equal(shiftDateLabel("05-12-2026", -3), "05-09-2026");
});

test("detects written-month dates with document and typographic separators", () => {
  const result = findDates(
    "Issued 01-Dec-2025. Renew 01 ‑ Dec ‑ 2025. Archive DEC/01/2025.",
    interpretInstruction("Move all dates forward 3 months.", 3, "forward"),
  );

  assert.deepEqual(
    result.dates.map(({ original, updated }) => ({ original, updated })),
    [
      { original: "01-Dec-2025", updated: "01-Mar-2026" },
      { original: "01 ‑ Dec ‑ 2025", updated: "01 ‑ Mar ‑ 2026" },
      { original: "DEC/01/2025", updated: "MAR/01/2026" },
    ],
  );
});

test("detects common year-first and weekday day-first formats without accepting invalid dates", () => {
  const result = findDates(
    "Review Wed, 01 Dec 2025; archive 2025/12/01; ignore 31-Feb-2025.",
    interpretInstruction("Move all dates forward 2 months.", 2, "forward"),
  );

  assert.deepEqual(
    result.dates.map(({ original, updated }) => ({ original, updated })),
    [
      { original: "Wed, 01 Dec 2025", updated: "Sun, 01 Feb 2026" },
      { original: "2025/12/01", updated: "2026/02/01" },
    ],
  );
});

test("detects and shifts weekday-prefixed month dates while preserving their style", () => {
  const result = findDates(
    "The review is scheduled for Wed Mar 04 2026.",
    interpretInstruction("Move all dates forward 3 months.", 3, "forward"),
  );

  assert.equal(result.dates.length, 1);
  assert.equal(result.dates[0].original, "Wed Mar 04 2026");
  assert.equal(result.dates[0].updated, "Thu Jun 04 2026");
  assert.equal(shiftDateLabel("Wednesday, March 04, 2026", -1), "Wednesday, February 04, 2026");
});

test("detects timestamp dates after a Date label and preserves their time and timezone", () => {
  const result = findDates(
    "Date : Wed Mar 04 15:57:15 GMT 2026",
    interpretInstruction("Move all dates forward 3 months.", 3, "forward"),
  );

  assert.equal(result.dates.length, 1);
  assert.equal(result.dates[0].original, "Wed Mar 04 15:57:15 GMT 2026");
  assert.equal(result.dates[0].updated, "Thu Jun 04 15:57:15 GMT 2026");
  assert.match(result.dates[0].context, /Date : Wed Mar 04 15:57:15 GMT 2026/);
});
