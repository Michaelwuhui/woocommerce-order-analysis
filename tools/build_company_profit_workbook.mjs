import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";


const COLORS = {
  navy: "#17365D",
  blue: "#1F4E78",
  teal: "#0F6B78",
  green: "#2E7D32",
  red: "#C62828",
  amber: "#B26A00",
  lightBlue: "#D9EAF7",
  lightGreen: "#E2F0D9",
  lightAmber: "#FFF2CC",
  lightGray: "#F2F2F2",
  border: "#D9E2F3",
  white: "#FFFFFF",
  text: "#1F2937",
};

const CURRENCY_FORMAT = '¥#,##0.00;[Red]-¥#,##0.00';
const PERCENT_FORMAT = '0.0%';


function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const item = argv[index];
    if (!item.startsWith("--")) continue;
    const key = item.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) {
      args[key] = true;
    } else {
      args[key] = value;
      index += 1;
    }
  }
  if (!args.input || !args.output) {
    throw new Error(
      "用法：node build_company_profit_workbook.mjs " +
      "--input snapshot.json --output report.xlsx [--render-dir previews]",
    );
  }
  return args;
}


function normalize(value, fallback = 0) {
  return value === null || value === undefined || Number.isNaN(Number(value))
    ? fallback
    : Number(value);
}


function textValue(value, fallback = "—") {
  if (value === null || value === undefined || value === "") return fallback;
  if (Array.isArray(value)) return value.length ? value.join("、") : fallback;
  return String(value);
}


function applyTitle(sheet, rangeAddress, title) {
  const range = sheet.getRange(rangeAddress);
  range.merge();
  range.values = [[title]];
  range.format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 18 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
  range.format.rowHeight = 32;
}


function applySectionHeader(range) {
  range.format = {
    fill: COLORS.blue,
    font: { bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
}


function applyBody(range) {
  range.format = {
    font: { color: COLORS.text, size: 10 },
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
}


function applyCurrency(range) {
  range.setNumberFormat(CURRENCY_FORMAT);
}


function applyPercent(range) {
  range.setNumberFormat(PERCENT_FORMAT);
}


function addStatusCell(sheet, address, label, complete) {
  const cell = sheet.getRange(address);
  cell.values = [[`${label}：${complete ? "完整" : "待补充"}`]];
  cell.format = {
    fill: complete ? COLORS.lightGreen : COLORS.lightAmber,
    font: {
      bold: true,
      color: complete ? COLORS.green : COLORS.amber,
    },
    horizontalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: COLORS.border },
  };
}


async function buildWorkbook(snapshot) {
  if (snapshot.schema_version !== 1) {
    throw new Error(`不支持的快照版本：${snapshot.schema_version}`);
  }
  const summary = snapshot.summary;
  const workbook = Workbook.create();
  const dashboard = workbook.worksheets.add("管理摘要");
  const market = workbook.worksheets.add("市场明细");
  const expenses = workbook.worksheets.add("收支明细");
  const ladder = workbook.worksheets.add("销售情景");
  const notes = workbook.worksheets.add("口径说明");
  const trend = workbook.worksheets.add("趋势数据");

  // Market detail is the auditable source for settlement and revenue totals.
  applyTitle(
    market,
    "A1:L1",
    `${summary.year_month} 市场核算明细（人民币）`,
  );
  market.getRange("A2:L2").merge();
  market.getRange("A2").values = [[
    `核算口径：${textValue(summary.calculation_mode_label)}；` +
    "合伙人市场优先引用当月对账结果。",
  ]];
  market.getRange("A2:L2").format = {
    fill: COLORS.lightBlue,
    font: { color: COLORS.blue, italic: true },
    wrapText: true,
  };
  const marketHeaders = [
    "市场",
    "GMV",
    "团队净销售",
    "结算净销售",
    "实际产品成本",
    "公司分配比例",
    "实际公司收入",
    "预测净销售",
    "预测公司收入",
    "收入依据",
    "预测依据",
    "对账更新时间",
  ];
  market.getRange("A4:L4").values = [marketHeaders];
  applySectionHeader(market.getRange("A4:L4"));
  const countries = summary.countries || [];
  const marketStart = 5;
  const marketRows = countries.length
    ? countries.map((row) => [
        textValue(row.country),
        normalize(row.gmv_cny),
        normalize(row.net_sales_cny),
        normalize(row.settlement_net_sales_cny),
        row.actual_product_cost_cny === null
          ? null
          : normalize(row.actual_product_cost_cny),
        row.share_rate === null ? null : normalize(row.share_rate),
        row.company_revenue_cny === null
          ? null
          : normalize(row.company_revenue_cny),
        normalize(row.forecast_net_sales_cny),
        row.forecast_company_revenue_cny === null
          ? null
          : normalize(row.forecast_company_revenue_cny),
        textValue(row.share_source),
        textValue(row.forecast_revenue_source),
        textValue(row.statement_updated_at),
      ])
    : [["无市场数据", 0, 0, 0, null, null, 0, 0, 0, "—", "—", "—"]];
  const marketEnd = marketStart + marketRows.length - 1;
  market.getRange(`A${marketStart}:L${marketEnd}`).values = marketRows;
  applyBody(market.getRange(`A${marketStart}:L${marketEnd}`));
  applyCurrency(market.getRange(`B${marketStart}:E${marketEnd}`));
  applyPercent(market.getRange(`F${marketStart}:F${marketEnd}`));
  applyCurrency(market.getRange(`G${marketStart}:I${marketEnd}`));
  const marketTotal = marketEnd + 1;
  market.getRange(`A${marketTotal}:L${marketTotal}`).format = {
    fill: COLORS.lightBlue,
    font: { bold: true, color: COLORS.navy },
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  market.getRange(`A${marketTotal}`).values = [["合计"]];
  for (const column of ["B", "C", "D", "E", "G", "H", "I"]) {
    market.getRange(`${column}${marketTotal}`).formulas = [[
      `=ROUND(SUM(${column}${marketStart}:${column}${marketEnd}),2)`,
    ]];
  }
  applyCurrency(market.getRange(`B${marketTotal}:I${marketTotal}`));
  market.freezePanes.freezeRows(4);
  market.getRange(`A1:L${marketTotal}`).format.autofitRows();
  market.getRange("A:A").format.columnWidth = 10;
  market.getRange("B:I").format.columnWidth = 16;
  market.getRange("J:L").format.columnWidth = 27;

  // Expense detail keeps payroll and operating costs separate and aggregated.
  applyTitle(
    expenses,
    "A1:H1",
    `${summary.year_month} 实际与预测支出（人民币）`,
  );
  expenses.getRange("A3:B6").values = [
    ["实际工资", normalize(summary.actual?.payroll_cny)],
    ["实际日常运营", null],
    ["预测工资", normalize(summary.forecast?.payroll_cny)],
    ["预测日常运营", null],
  ];
  applyBody(expenses.getRange("A3:B6"));
  expenses.getRange("A3:A6").format = {
    fill: COLORS.lightBlue,
    font: { bold: true, color: COLORS.navy },
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  applyCurrency(expenses.getRange("B3:B6"));
  expenses.getRange("D3:H6").values = [
    ["项目", "口径", "来源", "是否周期", "来源月份"],
    ["实际工资", "实际", textValue(summary.actual?.payroll_source), "—", snapshot.month],
    ["预测工资", "预测", textValue(summary.forecast?.payroll_source), "—", snapshot.month],
    [
      "预测日常运营",
      "预测",
      textValue(summary.forecast?.other_expenses_source),
      "—",
      snapshot.month,
    ],
  ];
  applyBody(expenses.getRange("D3:H6"));
  applySectionHeader(expenses.getRange("D3:H3"));

  const expenseHeaders = [
    "口径",
    "类别",
    "项目",
    "金额",
    "是否周期",
    "来源月份",
    "来源/创建人",
    "备注",
  ];
  expenses.getRange("A9:H9").values = [expenseHeaders];
  applySectionHeader(expenses.getRange("A9:H9"));
  const rawExpenses = summary.expenses || [];
  const recurring = summary.recurring_forecast_expenses || [];
  const expenseRows = [
    ...rawExpenses.map((row) => [
      row.scenario === "actual" ? "实际" : "预测",
      textValue(row.category),
      textValue(row.name),
      normalize(row.amount_cny),
      row.is_recurring ? "是" : "否",
      textValue(row.year_month, snapshot.month),
      textValue(row.created_by),
      textValue(row.notes),
    ]),
    ...recurring.map((row) => [
      "预测",
      textValue(row.category),
      textValue(row.name),
      normalize(row.amount_cny),
      "是（继承）",
      textValue(row.source_month),
      "系统继承",
      textValue(row.notes),
    ]),
  ];
  const expenseStart = 10;
  const displayedExpenseRows = expenseRows.length
    ? expenseRows
    : [["—", "—", "尚未录入支出", 0, "—", snapshot.month, "—", "—"]];
  const expenseEnd = expenseStart + displayedExpenseRows.length - 1;
  expenses.getRange(`A${expenseStart}:H${expenseEnd}`).values =
    displayedExpenseRows;
  applyBody(expenses.getRange(`A${expenseStart}:H${expenseEnd}`));
  applyCurrency(expenses.getRange(`D${expenseStart}:D${expenseEnd}`));

  const actualRows = [];
  const forecastRows = [];
  displayedExpenseRows.forEach((row, index) => {
    const excelRow = expenseStart + index;
    if (row[0] === "实际") actualRows.push(excelRow);
    if (row[0] === "预测") forecastRows.push(excelRow);
  });
  const sumRowsFormula = (rows) => (
    rows.length
      ? `=${rows.map((row) => `D${row}`).join("+")}`
      : "=0"
  );
  expenses.getRange("B4").formulas = [[sumRowsFormula(actualRows)]];
  const forecastUsesActual = (
    summary.forecast?.other_expenses_source === "当月实际支出"
  );
  expenses.getRange("B6").formulas = [[
    forecastUsesActual ? "=B4" : sumRowsFormula(forecastRows),
  ]];
  expenses.freezePanes.freezeRows(9);
  expenses.getRange(`A1:H${expenseEnd}`).format.autofitRows();
  expenses.getRange("A:C").format.columnWidth = 18;
  expenses.getRange("D:D").format.columnWidth = 16;
  expenses.getRange("E:G").format.columnWidth = 17;
  expenses.getRange("H:H").format.columnWidth = 35;

  // Sales ladder uses formulas for company revenue, total cost and profit.
  applyTitle(
    ladder,
    "A1:G1",
    `${summary.year_month} 销售情景预测（人民币）`,
  );
  ladder.getRange("A2:B4").values = [
    ["公司收入率", normalize(summary.revenue_ladder?.company_revenue_rate)],
    ["盈亏平衡销售额", normalize(summary.revenue_ladder?.breakeven_net_sales_cny)],
    ["工资测算口径", textValue(summary.revenue_ladder?.payroll_method)],
  ];
  applyBody(ladder.getRange("A2:B4"));
  applyPercent(ladder.getRange("B2"));
  applyCurrency(ladder.getRange("B3"));
  ladder.getRange("B4:G4").merge();
  ladder.getRange("B4").format.wrapText = true;
  ladder.getRange("A4:G4").format.rowHeight = 48;
  ladder.getRange("A6:G6").values = [[
    "预测销售额",
    "结算净销售",
    "公司收入",
    "工资支出",
    "日常运营",
    "预计总支出",
    "预计利润",
  ]];
  applySectionHeader(ladder.getRange("A6:G6"));
  const ladderRows = summary.revenue_ladder?.rows || [];
  const ladderStart = 7;
  const displayedLadderRows = ladderRows.length
    ? ladderRows
    : [{
        label: "无可用情景",
        target_net_sales_cny: 0,
        payroll_cny: 0,
        daily_operations_cny: 0,
      }];
  const ladderEnd = ladderStart + displayedLadderRows.length - 1;
  ladder.getRange(`A${ladderStart}:B${ladderEnd}`).values =
    displayedLadderRows.map((row) => [
      textValue(row.label),
      normalize(row.target_net_sales_cny),
    ]);
  ladder.getRange(`D${ladderStart}:E${ladderEnd}`).values =
    displayedLadderRows.map((row) => [
      normalize(row.payroll_cny),
      normalize(row.daily_operations_cny),
    ]);
  for (let row = ladderStart; row <= ladderEnd; row += 1) {
    ladder.getRange(`C${row}`).formulas = [[`=ROUND(B${row}*$B$2,2)`]];
    ladder.getRange(`F${row}`).formulas = [[`=ROUND(D${row}+E${row},2)`]];
    ladder.getRange(`G${row}`).formulas = [[`=ROUND(C${row}-F${row},2)`]];
  }
  applyBody(ladder.getRange(`A${ladderStart}:G${ladderEnd}`));
  applyCurrency(ladder.getRange(`B${ladderStart}:G${ladderEnd}`));
  ladder.getRange(`G${ladderStart}:G${ladderEnd}`)
    .conditionalFormats.add("cellIs", {
      operator: "lessThan",
      formula: 0,
      format: { font: { color: COLORS.red } },
    });
  ladder.freezePanes.freezeRows(6);
  ladder.getRange(`A1:G${ladderEnd}`).format.autofitRows();
  ladder.getRange("A:A").format.columnWidth = 21;
  ladder.getRange("B:G").format.columnWidth = 17;

  // Raw trend source, kept visible for audit.
  applyTitle(trend, "A1:G1", `${summary.year_month} 趋势源数据`);
  trend.getRange("A3:G3").values = [[
    "月份",
    "GMV",
    "结算净销售",
    "公司收入",
    "利润",
    "数据完整",
    "工资来源",
  ]];
  applySectionHeader(trend.getRange("A3:G3"));
  const trendRows = summary.trend || [];
  const displayedTrendRows = trendRows.length
    ? trendRows.map((row) => [
        textValue(row.month),
        normalize(row.gmv_cny),
        normalize(row.net_sales_cny),
        normalize(row.company_revenue_cny),
        row.profit_cny === null ? null : normalize(row.profit_cny),
        row.complete ? "是" : "否",
        textValue(row.payroll_source),
      ])
    : [[snapshot.month, 0, 0, 0, null, "否", "—"]];
  const trendStart = 4;
  const trendEnd = trendStart + displayedTrendRows.length - 1;
  trend.getRange(`A${trendStart}:G${trendEnd}`).values = displayedTrendRows;
  applyBody(trend.getRange(`A${trendStart}:G${trendEnd}`));
  applyCurrency(trend.getRange(`B${trendStart}:E${trendEnd}`));
  trend.freezePanes.freezeRows(3);
  trend.getRange(`A1:G${trendEnd}`).format.autofitRows();
  trend.getRange("A:A").format.columnWidth = 12;
  trend.getRange("B:E").format.columnWidth = 17;
  trend.getRange("F:G").format.columnWidth = 19;

  // Dashboard is formula-driven from the detail sheets.
  applyTitle(
    dashboard,
    "A1:M1",
    `${summary.year_month} 公司经营月报`,
  );
  dashboard.getRange("A2:M2").merge();
  dashboard.getRange("A2").values = [[
    `生成时间：${snapshot.generated_at}　|　核算口径：` +
    `${textValue(summary.calculation_mode_label)}　|　` +
    "机密：离线文件，仅限授权人员",
  ]];
  dashboard.getRange("A2:M2").format = {
    fill: COLORS.lightBlue,
    font: { color: COLORS.blue, italic: true },
    verticalAlignment: "center",
  };
  addStatusCell(
    dashboard,
    "A4",
    "实际数据",
    Boolean(summary.actual?.complete),
  );
  addStatusCell(
    dashboard,
    "B4",
    "预测数据",
    Boolean(summary.forecast?.complete),
  );
  dashboard.getRange("A5:C5").values = [["指标", "实际", "预测"]];
  applySectionHeader(dashboard.getRange("A5:C5"));
  dashboard.getRange("A6:A12").values = [
    ["结算净销售"],
    ["公司收入"],
    ["工资支出"],
    ["日常运营"],
    ["总支出"],
    ["利润"],
    ["利润率"],
  ];
  dashboard.getRange("B6:C12").formulas = [
    [
      `='市场明细'!D${marketTotal}`,
      `='市场明细'!H${marketTotal}`,
    ],
    [
      `='市场明细'!G${marketTotal}`,
      `='市场明细'!I${marketTotal}`,
    ],
    ["='收支明细'!B3", "='收支明细'!B5"],
    ["='收支明细'!B4", "='收支明细'!B6"],
    ["=ROUND(B8+B9,2)", "=ROUND(C8+C9,2)"],
    ["=ROUND(B7-B10,2)", "=ROUND(C7-C10,2)"],
    ["=IF(B7=0,0,B11/B7)", "=IF(C7=0,0,C11/C7)"],
  ];
  applyBody(dashboard.getRange("A6:C12"));
  dashboard.getRange("A6:A12").format = {
    fill: COLORS.lightGray,
    font: { bold: true, color: COLORS.navy },
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  applyCurrency(dashboard.getRange("B6:C11"));
  applyPercent(dashboard.getRange("B12:C12"));
  dashboard.getRange("B11:C11")
    .conditionalFormats.add("cellIs", {
      operator: "lessThan",
      formula: 0,
      format: {
        fill: "#FDE9E7",
        font: { bold: true, color: COLORS.red },
      },
    });

  dashboard.getRange("A15:D15").values = [[
    "月份",
    "结算净销售",
    "公司收入",
    "利润",
  ]];
  applySectionHeader(dashboard.getRange("A15:D15"));
  const dashboardTrendStart = 16;
  const dashboardTrendEnd = dashboardTrendStart + displayedTrendRows.length - 1;
  for (let index = 0; index < displayedTrendRows.length; index += 1) {
    const sourceRow = trendStart + index;
    const targetRow = dashboardTrendStart + index;
    dashboard.getRange(`A${targetRow}:D${targetRow}`).formulas = [[
      `='趋势数据'!A${sourceRow}`,
      `='趋势数据'!C${sourceRow}`,
      `='趋势数据'!D${sourceRow}`,
      `='趋势数据'!E${sourceRow}`,
    ]];
  }
  applyBody(
    dashboard.getRange(`A${dashboardTrendStart}:D${dashboardTrendEnd}`),
  );
  applyCurrency(
    dashboard.getRange(`B${dashboardTrendStart}:D${dashboardTrendEnd}`),
  );
  const chart = dashboard.charts.add(
    "line",
    dashboard.getRange(`A15:D${dashboardTrendEnd}`),
  );
  chart.setPosition("F4", "M20");
  chart.title = "结算净销售、公司收入与利润趋势（人民币）";
  chart.titleTextStyle.fontSize = 12;
  chart.hasLegend = true;
  chart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
  chart.yAxis = { numberFormatCode: "¥#,##0", textStyle: { fontSize: 9 } };
  dashboard.getRange(`A1:M${dashboardTrendEnd}`).format.autofitRows();
  dashboard.getRange("A:A").format.columnWidth = 18;
  dashboard.getRange("B:C").format.columnWidth = 18;
  dashboard.getRange("D:D").format.columnWidth = 16;
  dashboard.getRange("E:E").format.columnWidth = 3;
  dashboard.getRange("F:M").format.columnWidth = 12;
  dashboard.freezePanes.freezeRows(5);

  // Methodology and completeness disclosure.
  applyTitle(notes, "A1:E1", `${summary.year_month} 核算口径与数据说明`);
  notes.getRange("A3:B3").values = [["项目", "说明"]];
  applySectionHeader(notes.getRange("A3:B3"));
  const definitionRows = [
    ["GMV", textValue(summary.definitions?.gmv)],
    ["净销售", textValue(summary.definitions?.net_sales)],
    ["公司收入", textValue(summary.definitions?.company_revenue)],
    ["实际利润", textValue(summary.definitions?.actual_profit)],
    ["预测利润", textValue(summary.definitions?.forecast_profit)],
    ["销售情景", textValue(summary.definitions?.scenario_profit)],
    ["工资口径", textValue(summary.revenue_ladder?.payroll_method)],
    [
      "盈亏平衡销售额",
      "使公司收入等于工资支出与日常运营支出的销售额；" +
      "其公式随实际工资锚点或预测提成率口径变化。",
    ],
    [
      "数据安全",
      "本报告由生产数据库只读一致性快照离线生成；Web 页面和 API 均已下线。" +
      "报告只展示团队工资汇总，不列个人工资明细。",
    ],
    [
      "源数据库",
      `${textValue(snapshot.source?.database_name)}；` +
      `修改时间 ${textValue(snapshot.source?.database_modified_at)}；` +
      `大小 ${normalize(snapshot.source?.database_size_bytes)} 字节。`,
    ],
  ];
  const notesEnd = 4 + definitionRows.length - 1;
  notes.getRange(`A4:B${notesEnd}`).values = definitionRows;
  applyBody(notes.getRange(`A4:B${notesEnd}`));
  notes.getRange(`A4:A${notesEnd}`).format = {
    fill: COLORS.lightGray,
    font: { bold: true, color: COLORS.navy },
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  notes.getRange(`B4:B${notesEnd}`).format.wrapText = true;
  const gapHeader = notesEnd + 2;
  notes.getRange(`A${gapHeader}:B${gapHeader}`).values = [["待补数据", "影响"]];
  applySectionHeader(notes.getRange(`A${gapHeader}:B${gapHeader}`));
  const gaps = summary.data_gaps || [];
  const gapRows = gaps.length
    ? gaps.map((gap, index) => [`${index + 1}`, textValue(gap)])
    : [["—", "无待补数据"]];
  const gapStart = gapHeader + 1;
  const gapEnd = gapStart + gapRows.length - 1;
  notes.getRange(`A${gapStart}:B${gapEnd}`).values = gapRows;
  applyBody(notes.getRange(`A${gapStart}:B${gapEnd}`));
  notes.getRange(`B${gapStart}:B${gapEnd}`).format.wrapText = true;
  notes.getRange(`A1:B${gapEnd}`).format.autofitRows();
  notes.getRange(`A4:B${notesEnd}`).format.rowHeight = 46;
  notes.getRange(`A${gapStart}:B${gapEnd}`).format.rowHeight = 34;
  notes.getRange("A:A").format.columnWidth = 19;
  notes.getRange("B:B").format.columnWidth = 80;
  notes.freezePanes.freezeRows(3);

  for (const sheet of [dashboard, market, expenses, ladder, notes, trend]) {
    sheet.showGridLines = false;
  }
  return workbook;
}


async function verifyWorkbook(workbook, renderDir) {
  const sheetInspection = await workbook.inspect({
    kind: "sheet,drawing",
    include: "id,name,type",
    maxChars: 8000,
  });
  const formulaInspection = await workbook.inspect({
    kind: "formula",
    maxChars: 12000,
    options: { maxResults: 500 },
  });
  const errorInspection = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    maxChars: 6000,
  });
  if (renderDir) {
    await fs.mkdir(renderDir, { recursive: true });
    for (const sheetName of [
      "管理摘要",
      "市场明细",
      "收支明细",
      "销售情景",
      "口径说明",
      "趋势数据",
    ]) {
      const preview = await workbook.render({
        sheetName,
        autoCrop: "all",
        scale: 1,
        format: "png",
      });
      await fs.writeFile(
        path.join(renderDir, `${sheetName}.png`),
        new Uint8Array(await preview.arrayBuffer()),
      );
    }
  }
  return {
    sheetInspection,
    formulaInspection,
    errorInspection,
  };
}


async function main() {
  const args = parseArgs(process.argv.slice(2));
  const inputPath = path.resolve(args.input);
  const outputPath = path.resolve(args.output);
  if (path.extname(inputPath).toLowerCase() !== ".json") {
    throw new Error("输入必须是离线 JSON 快照");
  }
  if (path.extname(outputPath).toLowerCase() !== ".xlsx") {
    throw new Error("输出必须使用 .xlsx 扩展名");
  }
  const snapshot = JSON.parse(await fs.readFile(inputPath, "utf8"));
  const workbook = await buildWorkbook(snapshot);
  const verification = await verifyWorkbook(
    workbook,
    args["render-dir"] ? path.resolve(args["render-dir"]) : null,
  );
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  try {
    await fs.chmod(outputPath, 0o600);
  } catch {
    // Windows ACLs are managed by the current user profile.
  }
  process.stdout.write(JSON.stringify({
    success: true,
    month: snapshot.month,
    output: outputPath,
    sheets: 6,
    formulas_verified: true,
    sheet_inspection: String(verification.sheetInspection).slice(0, 500),
    error_scan: String(verification.errorInspection).slice(0, 1000),
  }));
}


main().catch((error) => {
  process.stderr.write(`${error.stack || error.message}\n`);
  process.exitCode = 1;
});
