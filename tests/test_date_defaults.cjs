const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function element(id = "") {
  const listeners = new Map();
  return {
    id, value: "", textContent: "", innerHTML: "", disabled: false,
    classList: { add() {}, remove() {} },
    addEventListener(name, handler) { listeners.set(name, handler); },
    dispatch(name) { listeners.get(name)?.({ target: this }); },
    appendChild() {}, click() {}, querySelector() { return element(); },
  };
}

const ids = [
  "fileGrid", "log", "bulkInput", "dropzone", "currentStartDate", "currentEndDate",
  "previousStartDate", "previousEndDate", "periodText", "readyCount", "statusText",
  "generateBtn", "clearBtn", "progress",
];
const fields = Object.fromEntries(ids.map((id) => [id, element(id)]));
const context = vm.createContext({
  document: {
    querySelector(selector) { return fields[selector.slice(1)]; },
    createElement() { return element(); },
  },
  window: {},
});
const script = fs.readFileSync(path.join(__dirname, "..", "static", "app.js"), "utf8");
vm.runInContext(script, context);
function run(code) { return vm.runInContext(code, context); }

function check(currentStart, currentEnd, previousStart, previousEnd) {
  assert.equal(fields.currentStartDate.value, currentStart);
  assert.equal(fields.currentEndDate.value, currentEnd);
  assert.equal(fields.previousStartDate.value, previousStart);
  assert.equal(fields.previousEndDate.value, previousEnd);
}

// Thursday in Shanghai: the current partial week is excluded.
run("autoFillDates(new Date('2026-11-19T01:00:00Z'))");
check("2026-11-09", "2026-11-15", "2026-11-02", "2026-11-08");

// Sunday is not complete until the following Monday begins in Shanghai.
run("autoFillDates(new Date('2026-11-22T10:00:00Z'))");
check("2026-11-09", "2026-11-15", "2026-11-02", "2026-11-08");
run("autoFillDates(new Date('2026-11-22T16:00:00Z'))");
check("2026-11-16", "2026-11-22", "2026-11-09", "2026-11-15");

// A Monday near New Year still yields two full seven-day weeks.
run("autoFillDates(new Date('2027-01-03T16:00:00Z'))");
check("2026-12-28", "2027-01-03", "2026-12-21", "2026-12-27");

fields.currentStartDate.value = "2026-12-29";
fields.currentStartDate.dispatch("input");
fields.previousEndDate.value = "2026-12-26";
fields.previousEndDate.dispatch("input");
run("filesByKey.set('oraDaily', {name: 'Ora外送日报_20261026.xlsx'}); autoFillDates(new Date('2027-01-04T01:00:00Z'))");
check("2026-12-29", "2027-01-03", "2026-12-21", "2026-12-26");

fields.clearBtn.dispatch("click");
assert.equal(run("manuallyEditedDates.size"), 0);
assert.equal(run("filesByKey.size"), 0);
assert.match(fields.currentStartDate.value, /^\d{4}-\d{2}-\d{2}$/);
console.log("date defaults passed");
