import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import "@univerjs/preset-sheets-core/lib/index.css";

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS) },
  presets: [
    UniverSheetsCorePreset({
      container: "app",
    }),
  ],
});

const workbook = univerAPI.createWorkbook({
  id: "calcplus-spike",
  name: "univer.xlsx",
  sheetOrder: ["sheet-01"],
  sheets: {
    "sheet-01": {
      id: "sheet-01",
      name: "Sales",
      rowCount: 1000,
      columnCount: 30,
    },
  },
});

const sheet = workbook.getActiveSheet();

// Header and a sample block that mirrors the custom spike.
const headers = ["Region", "Q1", "Q2", "Q3", "Q4", "Total", "Growth", "Owner", "Notes"];
sheet.getRange("A1:I1").setValues([headers]);
sheet.getRange("A1:I1").setFontWeight("bold").setBackgroundColor("#f3f2f1");

const regions = ["North", "South", "East", "West", "Central"];
const rows = [];
for (let r = 0; r < 200; r++) {
  const row = [regions[r % regions.length]];
  let total = 0;
  for (let c = 1; c <= 4; c++) {
    const v = 1000 + ((r * 37 + c * 113) % 9000);
    total += v;
    row.push(v);
  }
  row.push(total);
  row.push((r % 21 - 10) / 100);
  row.push(["Ana", "Ben", "Chloe", "Dev", "Eli"][r % 5]);
  row.push(r % 17 === 0 ? "Reviewed" : "");
  rows.push(row);
}
sheet.getRange("A2:I201").setValues(rows);
sheet.getRange("B2:E201").setNumberFormat("#,##0");
sheet.getRange("F2:F201").setNumberFormat("#,##0");
sheet.getRange("G2:G201").setNumberFormat("0%");
sheet.getRange("F2:F201").setFontWeight("bold");

// Formulas, so the Univer engine has something to calculate.
sheet.getRange("F2").setFormula("=SUM(B2:E2)");
sheet.getRange("F3").setFormula("=SUM(B3:E3)");

// A frozen header row and first column, matching the custom spike.
sheet.setFreeze({ startRow: 1, startColumn: 1, xSplit: 1, ySplit: 1 });

// Expose a handle for the measurement harness.
window.__univer = {
  univerAPI,
  workbook,
  sheet,
};

requestAnimationFrame(() => requestAnimationFrame(() => {
  window.__readyAt = performance.now();
}));
