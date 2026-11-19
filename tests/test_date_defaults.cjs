const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function element(id = "") {
  const listeners = new Map();
  return {
    id,
    value: "",
    textContent: "",
    innerHTML: "",
    disabled: false,
    classList: { add() {}, remove() {} },
    addEventListener(name, handler) { listeners.set(name, handler); },
    dispatch(name) { listeners.get(name)?.({ target: this }); },
    appendChild() {},
    click() {},
    querySelector() { return element(); },
  };
}

const ids = [
  "fileGrid", "log", "bulkInput", "dropzone", "currentStartDate", "currentEndDate",
  "previousStartDate", "previousEndDate", "periodText", "readyCount", "statusText",
  "generateBtn", "clearBtn", "progress",
];
const fields = Object.fromEntries(ids.map((id) => [id, element(id)]));
const XLSX = {
  read(buffer, options) {
    const key = Buffer.from(buffer).toString();
    if (key === "previous") return { SheetNames: ["V2", "11.2-11.8", "上期"] };
    if (key === "previousSame") return { SheetNames: ["V2", "11.9-11.15", "上期"] };
    if (key === "daily") {
      return {
        SheetNames: ["Ora外送日报"],
        Sheets: { Ora外送日报: {
          "!ref": "A1:A3",
          A1: { v: "date_id" },
          A2: { v: "2026-11-09T00:00:00.000Z" },
          A3: { v: "2026-11-15" },
        } },
      };
    }
    if (key === "dailyNext") {
      return {
        SheetNames: ["Ora外送日报"],
        Sheets: { Ora外送日报: {
          "!ref": "A1:A3",
          A1: { v: "date_id" },
          A2: { v: "2026-11-16" },
          A3: { v: "2026-11-22" },
        } },
      };
    }
    if (key === "dailyPartial") {
      return {
        SheetNames: ["Ora外送日报"],
        Sheets: { Ora外送日报: {
          "!ref": "A1:A3",
          A1: { v: "date_id" },
          A2: { v: "2026-11-09" },
          A3: { v: "2026-11-18" },
        } },
      };
    }
    throw new Error(`Unknown workbook ${key}`);
  },
  utils: {
    decode_range(ref) {
      const last = Number(ref.match(/\d+$/)[0]) - 1;
      return { s: { r: 0, c: 0 }, e: { r: last, c: 0 } };
    },
    encode_cell({ r, c }) { return `${String.fromCharCode(65 + c)}${r + 1}`; },
  },
  SSF: { parse_date_code() { return null; } },
};
const context = vm.createContext({
  document: {
    querySelector(selector) { return fields[selector.slice(1)]; },
    createElement() { return element(); },
  },
  window: {},
  XLSX,
  Buffer,
});
const script = fs.readFileSync(path.join(__dirname, "..", "static", "app.js"), "utf8");
vm.runInContext(script, context);

function file(name, key) {
  return { name, async arrayBuffer() { return Buffer.from(key); } };
}
function run(code) { return vm.runInContext(code, context); }

(async () => {
  context.daily = file("Ora外送日报.xlsx", "daily");
  context.previous = file("旧文件名.xlsx", "previous");
  run("filesByKey.set('oraDaily', daily); filesByKey.set('previous', previous)");
  await run("autoFillDates()");
  assert.equal(fields.currentStartDate.value, "2026-11-09");
  assert.equal(fields.currentEndDate.value, "2026-11-15");
  assert.equal(fields.previousStartDate.value, "2026-11-02");
  assert.equal(fields.previousEndDate.value, "2026-11-08");

  run("filesByKey.clear()");
  context.wide = file("美团门店数据20261019_20261101.xlsx", "unused");
  context.narrow = file("美团订单数据20261026_20261101.xlsx", "unused");
  run("filesByKey.set('mtStore', wide); filesByKey.set('mtOrder', narrow)");
  await run("autoFillDates()");
  assert.equal(fields.currentStartDate.value, "2026-10-26");
  assert.equal(fields.currentEndDate.value, "2026-11-01");
  assert.equal(fields.previousStartDate.value, "2026-10-19");
  assert.equal(fields.previousEndDate.value, "2026-10-25");

  fields.clearBtn.dispatch("click");
  run("filesByKey.set('previous', previous)");
  await run("autoFillDates()");
  assert.equal(fields.currentStartDate.value, "");
  assert.equal(fields.currentEndDate.value, "");
  assert.equal(fields.previousStartDate.value, "2026-11-02");
  assert.equal(fields.previousEndDate.value, "2026-11-08");

  context.dailyPartial = file("Ora外送日报.xlsx", "dailyPartial");
  run("filesByKey.set('oraDaily', dailyPartial)");
  await run("autoFillDates()");
  assert.equal(fields.currentStartDate.value, "2026-11-09");
  assert.equal(fields.currentEndDate.value, "2026-11-15");
  assert.equal(fields.previousStartDate.value, "2026-11-02");
  assert.equal(fields.previousEndDate.value, "2026-11-08");

  fields.currentStartDate.value = "2026-11-10";
  fields.currentStartDate.dispatch("input");
  fields.previousEndDate.value = "2026-11-07";
  fields.previousEndDate.dispatch("input");
  context.dailyNext = file("Ora外送日报.xlsx", "dailyNext");
  run("filesByKey.set('oraDaily', dailyNext)");
  await run("autoFillDates()");
  assert.equal(fields.currentStartDate.value, "2026-11-10");
  assert.equal(fields.currentEndDate.value, "2026-11-22");
  assert.equal(fields.previousEndDate.value, "2026-11-07");

  fields.clearBtn.dispatch("click");
  context.previousSame = file("旧文件名.xlsx", "previousSame");
  run("filesByKey.set('oraDaily', daily); filesByKey.set('previous', previousSame)");
  await run("autoFillDates()");
  assert.equal(fields.previousStartDate.value, "2026-11-02");
  assert.equal(fields.previousEndDate.value, "2026-11-08");

  assert.equal(run("periodFromSheetName('12.29-1.4', '2027-01-05').start"), "2026-12-29");
  assert.equal(run("periodFromSheetName('12.29-1.4', '2027-01-05').end"), "2027-01-04");
  assert.equal(run("dateRangeFromName('美团门店数据20261019_20261025.xlsx').end"), "2026-10-25");
  console.log("date defaults passed");
})().catch((error) => { console.error(error); process.exitCode = 1; });
