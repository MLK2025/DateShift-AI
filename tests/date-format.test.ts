import assert from "node:assert/strict";
import test from "node:test";
import { findDates, interpretInstruction, shiftDateLabel } from "../app/page";

test("always treats hyphenated numeric dates as DD-MM-YYYY", () => {
  assert.equal(shiftDateLabel("20-12-2026", 3), "20-03-2027");
  assert.equal(shiftDateLabel("05-12-2026", 3), "05-03-2027");
  assert.equal(shiftDateLabel("05-12-2026", -3), "05-09-2026");
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
