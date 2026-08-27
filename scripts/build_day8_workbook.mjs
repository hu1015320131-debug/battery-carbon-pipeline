import fs from "node:fs/promises";
import process from "node:process";
import crypto from "node:crypto";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";


function argValue(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}


async function inspectReference() {
  const input = argValue("--inspect-reference");
  const outputDir = argValue("--preview-dir");
  if (!input || !outputDir) {
    throw new Error("--inspect-reference and --preview-dir are required");
  }
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
  const summary = await workbook.inspect({
    kind: "workbook,sheet,table",
    maxChars: 12000,
    tableMaxRows: 6,
    tableMaxCols: 8,
    tableMaxCellChars: 80,
  });
  await fs.mkdir(outputDir, { recursive: true });
  await fs.writeFile(`${outputDir}/reference_inspection.ndjson`, summary.ndjson, "utf8");
  const sheets = workbook.worksheets.items;
  const selected = [0, Math.floor(sheets.length / 2), sheets.length - 1]
    .filter((value, index, array) => value >= 0 && array.indexOf(value) === index);
  for (const index of selected) {
    const sheet = sheets[index];
    const preview = await workbook.render({
      sheetName: sheet.name,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
    await fs.writeFile(
      `${outputDir}/reference_${String(index + 1).padStart(2, "0")}_${safeName}.png`,
      new Uint8Array(await preview.arrayBuffer()),
    );
  }
  console.log(JSON.stringify({ status: "PASS", sheetCount: sheets.length, sheetNames: sheets.map((sheet) => sheet.name) }));
}


function parseCsv(text) {
  const source = text.replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (quoted) {
      if (char === '"' && source[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field.replace(/\r$/, ""));
    rows.push(row);
  }
  return rows.filter((values) => values.some((value) => value !== ""));
}


function columnName(index) {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}


const INTEGER_FIELDS = new Set([
  "Year", "Source_Row", "Header_Row", "Physical_Row_Count", "Column_Count",
  "Data_Row_Count", "Formula_Count", "Merged_Cell_Count", "Match_Count",
  "Input_Records", "Candidate_Records", "Excluded_Records", "Record_Count",
  "Activity_Candidate_Count", "Result_Candidate_Count", "Display_Decimal_Places",
  "Open_Items_Count", "Source_Field_Count", "Target_Field_Count", "Input_Row",
]);


const DECIMAL_FIELDS = new Set([
  "PCS", "Unit_Weight_g", "Original_Activity_Value", "Total_Weight_g",
  "Total_Weight_kg", "Total_Weight_t", "Activity_Diff_g", "Activity_Diff_Rate",
  "Activity_Data", "Activity_Data_Context", "Activity_Value_Parsed",
  "Activity_Data_Original", "Activity_Data_Normalized_kg", "Activity_Total_kg",
  "EF_Value", "EF_Value_Original", "EF_Value_Parsed",
  "EF_Value_Normalized_kgCO2e_per_kg", "Emission_kgCO2e",
  "Total_Emission_kgCO2e", "Sum_of_Row_Display_Emission",
  "Rounding_Reconciliation_Difference", "Fallback_Confidence_Score",
]);


const HIGH_PRECISION_TEXT_FIELDS = new Set([
  "Raw_Emission_kgCO2e", "Raw_Total_Emission_kgCO2e", "SHA256",
  "Source_SHA256", "Copy_SHA256", "Receipt_Source_SHA256",
  "Activity_Source_SHA256", "Third_Party_Input_Source_SHA256",
  "Raw_Input_SHA256", "Received_Input_SHA256", "Profile_Config_SHA256",
  "Calculation_Config_SHA256",
]);


function typedValue(header, value) {
  if (value === "") return null;
  if (HIGH_PRECISION_TEXT_FIELDS.has(header) || header.endsWith("_ID") || header === "Record_ID") {
    return value;
  }
  if (INTEGER_FIELDS.has(header) && /^-?\d+$/.test(value)) return Number(value);
  if (DECIMAL_FIELDS.has(header) && /^-?\d+(\.\d+)?$/.test(value)) return Number(value);
  if (value === "TRUE") return true;
  if (value === "FALSE") return false;
  return value;
}


function computeColumnWidth(header, rows, columnIndex) {
  let maxLength = Math.min(String(header).length, 30);
  for (const row of rows.slice(0, 80)) {
    const value = row[columnIndex];
    maxLength = Math.max(maxLength, Math.min(String(value ?? "").length, 32));
  }
  return Math.max(10, Math.min(34, maxLength + 2));
}


function styleTitle(sheet, lastColumn, title, subtitle) {
  const last = columnName(lastColumn - 1);
  sheet.mergeCells(`A1:${last}1`);
  sheet.mergeCells(`A2:${last}2`);
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A1:${last}1`).format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", size: 16, name: "Microsoft YaHei" },
    verticalAlignment: "center",
  };
  sheet.getRange(`A2:${last}2`).format = {
    fill: "#D9EAF7",
    font: { italic: true, color: "#5B6573", size: 10, name: "Microsoft YaHei" },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("1:1").format.rowHeight = 30;
  sheet.getRange("2:2").format.rowHeight = 28;
}


function applyStatusFormatting(range) {
  range.conditionalFormats.add("containsText", {
    text: "PASS",
    format: { fill: "#E2F0D9", font: { color: "#375623" } },
  });
  range.conditionalFormats.add("containsText", {
    text: "WARNING",
    format: { fill: "#FFF2CC", font: { color: "#7F6000" } },
  });
  range.conditionalFormats.add("containsText", {
    text: "FAIL",
    format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "BLOCKED",
    format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } },
  });
}


async function addDataSheet(workbook, definition) {
  const sheet = workbook.worksheets.getItem(definition.name);
  const csvText = await fs.readFile(definition.csvPath, "utf8");
  const parsed = parseCsv(csvText);
  if (parsed.length === 0) throw new Error(`No CSV rows for ${definition.name}`);
  const headers = parsed[0];
  const rawRows = parsed.slice(1);
  const dataRows = rawRows.map((row) => headers.map((header, index) => typedValue(header, row[index] ?? "")));
  const last = columnName(headers.length - 1);
  styleTitle(sheet, headers.length, definition.title, definition.subtitle);
  sheet.getRange(`A4:${last}4`).values = [headers];
  if (dataRows.length > 0) {
    sheet.getRangeByIndexes(4, 0, dataRows.length, headers.length).values = dataRows;
  }
  const tableEndRow = Math.max(4, dataRows.length + 4);
  const table = sheet.tables.add(`A4:${last}${tableEndRow}`, true, definition.tableName);
  table.style = "TableStyleMedium2";
  table.showBandedRows = true;
  table.showFilterButton = true;
  sheet.getRange(`A4:${last}4`).format = {
    fill: "#285E8E",
    font: { bold: true, color: "#FFFFFF", size: 10, name: "Microsoft YaHei" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.getRange(`A4:A${tableEndRow}`).format.fill = "#D9EAF7";
  sheet.getRange(`A5:${last}${tableEndRow}`).format.font = { size: 9, name: "Microsoft YaHei" };
  sheet.getRange(`A4:${last}${tableEndRow}`).format.verticalAlignment = "center";
  for (let index = 0; index < headers.length; index += 1) {
    const column = columnName(index);
    sheet.getRange(`${column}:${column}`).format.columnWidth = computeColumnWidth(headers[index], rawRows, index);
    if (INTEGER_FIELDS.has(headers[index])) {
      sheet.getRange(`${column}5:${column}${tableEndRow}`).format.numberFormat = "#,##0";
    } else if (DECIMAL_FIELDS.has(headers[index])) {
      const decimals = headers[index].includes("Rate") ? "0.0000000000" : "#,##0.000000";
      sheet.getRange(`${column}5:${column}${tableEndRow}`).format.numberFormat = decimals;
    }
  }
  const statusColumnPattern = /(Status|Eligible|Decision|Severity|Result|QC)/i;
  for (let index = 0; index < headers.length; index += 1) {
    if (statusColumnPattern.test(headers[index])) {
      const column = columnName(index);
      applyStatusFormatting(sheet.getRange(`${column}5:${column}${tableEndRow}`));
    }
  }
  sheet.freezePanes.freezeRows(4);
  sheet.freezePanes.freezeColumns(1);
  sheet.showGridLines = false;
  return { headers, rawRows, tableEndRow, lastColumn: last };
}


function addOverview(workbook, manifest) {
  const sheet = workbook.worksheets.getItem("01_运行概览");
  styleTitle(sheet, 6, manifest.workbookTitle, manifest.workbookSubtitle);
  const headers = ["检查项", "工作簿公式结果", "期望值", "判定", "含义", "证据"];
  const checks = [
    ["D5正式结果", "=COUNTA('15_D5端到端结果'!A5:A146)", 142, "端到端记录数", "15_D5端到端结果"],
    ["WP5冻结血缘", "=COUNTA('11_WP5冻结血缘'!A5:A146)", 142, "冻结32字段血缘记录数", "11_WP5冻结血缘"],
    ["Demo扩展血缘", "=COUNTA('13_Demo扩展血缘'!A5:A146)", 142, "扩展血缘一一对应", "13_Demo扩展血缘"],
    ["活动量合计", "='12_汇总与限制'!D5", Number(manifest.expected.activityTotalKg), "Decimal活动量", "12_汇总与限制"],
    ["历史因子", "='12_汇总与限制'!F5", Number(manifest.expected.efValue), "历史模拟因子", "12_汇总与限制"],
    ["未舍入排放总量", "='12_汇总与限制'!G5", manifest.expected.rawTotal, "CSV保留十进制真值", "12_汇总与限制"],
    ["官方六位展示总量", "='12_汇总与限制'!H5", Number(manifest.expected.officialTotal), "汇总后HALF_UP六位", "12_汇总与限制"],
    ["行级展示合计", "='12_汇总与限制'!I5", Number(manifest.expected.rowDisplayTotal), "每行六位值相加", "12_汇总与限制"],
    ["舍入勾稽差", "='12_汇总与限制'!J5", Number(manifest.expected.roundingDifference), "行级展示合计减官方总量", "12_汇总与限制"],
    ["正式阻断", "=COUNTIF('14_状态与OpenItems'!D5:D200,\"BLOCKED\")", 0, "不得存在正式阻断", "14_状态与OpenItems"],
  ];
  sheet.getRange("A4:F4").values = [headers];
  sheet.getRange(`A5:A${checks.length + 4}`).values = checks.map((row) => [row[0]]);
  sheet.getRange(`B5:B${checks.length + 4}`).formulas = checks.map((row) => [row[1]]);
  sheet.getRange(`C5:C${checks.length + 4}`).values = checks.map((row) => [row[2]]);
  sheet.getRange("D5").formulas = [["=IF(B5=C5,\"PASS\",\"FAIL\")"]];
  sheet.getRange(`D5:D${checks.length + 4}`).fillDown();
  sheet.getRange(`E5:F${checks.length + 4}`).values = checks.map((row) => [row[3], row[4]]);
  sheet.getRange("A4:F4").format = {
    fill: "#285E8E", font: { bold: true, color: "#FFFFFF", size: 10, name: "Microsoft YaHei" },
    horizontalAlignment: "center", verticalAlignment: "center",
  };
  sheet.getRange(`A5:A${checks.length + 4}`).format.fill = "#D9EAF7";
  sheet.getRange(`A4:F${checks.length + 4}`).format.borders = { preset: "inside", style: "thin", color: "#D9E2F3" };
  sheet.getRange(`B5:C${checks.length + 4}`).format.numberFormat = "#,##0.000000";
  sheet.getRange(`A4:F${checks.length + 4}`).format.font = { name: "Microsoft YaHei", size: 10 };
  sheet.getRange("A:A").format.columnWidth = 24;
  sheet.getRange("B:C").format.columnWidth = 22;
  sheet.getRange("D:D").format.columnWidth = 12;
  sheet.getRange("E:E").format.columnWidth = 34;
  sheet.getRange("F:F").format.columnWidth = 24;
  applyStatusFormatting(sheet.getRange(`D5:D${checks.length + 4}`));
  sheet.freezePanes.freezeRows(4);
  sheet.showGridLines = false;
  return { checkCount: checks.length };
}


async function sha256File(path) {
  const buffer = await fs.readFile(path);
  return crypto.createHash("sha256").update(buffer).digest("hex").toUpperCase();
}


async function buildWorkbook() {
  const buildStarted = Date.now();
  const manifestPath = argValue("--manifest");
  const outputPath = argValue("--output");
  const verificationPath = argValue("--verification");
  const previewDir = argValue("--preview-dir");
  if (!manifestPath || !outputPath || !verificationPath || !previewDir) {
    throw new Error("--manifest, --output, --verification and --preview-dir are required");
  }
  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  const workbook = Workbook.create();
  for (const definition of manifest.sheets) workbook.worksheets.add(definition.name);
  const sheetStats = {};
  for (const definition of manifest.sheets.filter((item) => item.kind === "csv")) {
    sheetStats[definition.name] = await addDataSheet(workbook, definition);
  }
  const overviewStats = addOverview(workbook, manifest);
  const sheetsBuiltAt = Date.now();
  await fs.mkdir(previewDir, { recursive: true });
  const previewStartedAt = Date.now();
  await Promise.all(manifest.sheets.map(async (definition) => {
    const preview = await workbook.render({
      sheetName: definition.name,
      range: "A1:J20",
      scale: 1,
      format: "png",
    });
    const safeName = definition.name.replace(/[\\/:*?"<>|]/g, "_");
    await fs.writeFile(
      `${previewDir}/${safeName}.png`,
      new Uint8Array(await preview.arrayBuffer()),
    );
  }));
  const previewCompletedAt = Date.now();
  const overviewInspect = await workbook.inspect({
    kind: "table,formula",
    sheetId: "01_运行概览",
    range: "A1:F14",
    include: "values,formulas",
    maxChars: 8000,
    tableMaxRows: 20,
    tableMaxCols: 8,
  });
  const formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "Day8 formula error scan",
  });
  const workbookInspectedAt = Date.now();
  await fs.mkdir(outputPath.substring(0, Math.max(outputPath.lastIndexOf("/"), outputPath.lastIndexOf("\\"))), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  const workbookExportedAt = Date.now();

  const readback = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
  const workbookImportedAt = Date.now();
  const readbackNames = readback.worksheets.items.map((sheet) => sheet.name);
  const readbackChecks = {};
  for (const definition of manifest.sheets.filter((item) => item.kind === "csv")) {
    const stats = sheetStats[definition.name];
    const sheet = readback.worksheets.getItem(definition.name);
    const header = sheet.getRangeByIndexes(3, 0, 1, stats.headers.length).values[0];
    const lastDataRow = sheet.getRangeByIndexes(3 + stats.rawRows.length, 0, 1, stats.headers.length).values[0];
    const rowAfterData = sheet.getRangeByIndexes(4 + stats.rawRows.length, 0, 1, stats.headers.length).values[0];
    const lastRowPresent = lastDataRow.some((value) => value !== "" && value !== null && value !== undefined);
    const nextRowEmpty = rowAfterData.every((value) => value === "" || value === null || value === undefined);
    readbackChecks[definition.name] = {
      expectedRows: stats.rawRows.length,
      readbackRows: lastRowPresent && nextRowEmpty ? stats.rawRows.length : -1,
      expectedFields: stats.headers.length,
      readbackFields: header.length,
      headerEqual: JSON.stringify(header) === JSON.stringify(stats.headers),
      lastRowPresent,
      nextRowEmpty,
    };
  }
  const structuralReadbackAt = Date.now();
  const readbackErrors = await readback.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "Day8 readback formula error scan",
  });
  const formulaReadbackAt = Date.now();
  const allReadbackEqual = Object.values(readbackChecks).every(
    (item) => item.expectedRows === item.readbackRows && item.expectedFields === item.readbackFields && item.headerEqual,
  );
  const formulaErrorText = `${formulaErrors.ndjson}\n${readbackErrors.ndjson}`;
  const formulaErrorCount = (formulaErrorText.match(/#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/g) || []).length;
  const readbackCompletedAt = Date.now();
  const verification = {
    status: readbackNames.length === manifest.sheets.length && allReadbackEqual && formulaErrorCount === 0 ? "PASS" : "FAIL",
    workbookPath: outputPath,
    workbookSha256: await sha256File(outputPath),
    expectedSheetCount: manifest.sheets.length,
    readbackSheetCount: readbackNames.length,
    sheetNamesEqual: JSON.stringify(readbackNames) === JSON.stringify(manifest.sheets.map((item) => item.name)),
    overviewCheckCount: overviewStats.checkCount,
    formulaErrorCount,
    allSourceTablesReadBackEqual: allReadbackEqual,
    readbackChecks,
    previewCount: manifest.sheets.length,
    overviewInspection: overviewInspect.ndjson,
    performanceSeconds: {
      buildSheets: (sheetsBuiltAt - buildStarted) / 1000,
      renderPreviews: (previewCompletedAt - previewStartedAt) / 1000,
      inspectWorkbook: (workbookInspectedAt - sheetsBuiltAt) / 1000,
      exportWorkbook: (workbookExportedAt - workbookInspectedAt) / 1000,
      generationTotal: (workbookExportedAt - buildStarted) / 1000,
      importAndReadback: (readbackCompletedAt - workbookExportedAt) / 1000,
      importWorkbook: (workbookImportedAt - workbookExportedAt) / 1000,
      structuralReadback: (structuralReadbackAt - workbookImportedAt) / 1000,
      formulaReadback: (formulaReadbackAt - structuralReadbackAt) / 1000,
      total: (readbackCompletedAt - buildStarted) / 1000,
    },
  };
  await fs.writeFile(verificationPath, `${JSON.stringify(verification, null, 2)}\n`, "utf8");
  console.log(JSON.stringify({ status: verification.status, workbookSha256: verification.workbookSha256, sheetCount: readbackNames.length, formulaErrorCount }));
}


if (process.argv.includes("--inspect-reference")) {
  await inspectReference();
} else if (process.argv.includes("--manifest")) {
  await buildWorkbook();
} else {
  throw new Error("Choose --inspect-reference or --manifest build mode.");
}
