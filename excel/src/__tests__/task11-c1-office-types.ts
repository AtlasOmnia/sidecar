// Compile-only C1 probe. This file is intentionally never imported at runtime;
// it must type-check against the installed @types/office-js declarations.
function task11C1OfficeJsTypeProbe(
  workbook: Excel.Workbook,
  worksheet: Excel.Worksheet,
  source: Excel.Range,
  target: Excel.Range,
  table: Excel.Table,
): void {
  source.autoFill(target);
  source.values = [[1]];
  source.formulas = [["=1"]];
  target.values = [[2]];
  target.formulas = [["=2"]];
  target.numberFormat = [["0.00", "0.00"], ["0.00", "0.00"]];
  target.format.font.bold = true;
  target.format.font.italic = true;
  target.format.font.color = "#111111";
  target.format.font.name = "Aptos";
  target.format.font.size = 12;
  target.format.font.underline = "Single";
  target.format.fill.color = "#eeeeee";
  target.format.horizontalAlignment = "Center";
  target.format.verticalAlignment = "Center";
  target.format.wrapText = true;
  target.format.indentLevel = 2;
  target.format.textOrientation = 15;
  target.format.borders.getItem("EdgeTop").style = "Continuous";
  target.format.borders.getItem("EdgeTop").color = "#222222";
  target.format.borders.getItem("EdgeTop").weight = "Medium";
  target.format.autofitColumns();
  target.format.autofitRows();
  target.format.load(["horizontalAlignment", "verticalAlignment", "wrapText", "indentLevel", "textOrientation", "columnWidth", "rowHeight"]);
  target.format.font.load(["bold", "italic", "color", "size", "name", "underline"]);
  target.format.fill.load("color");
  target.format.borders.load(["items/sideIndex", "items/style", "items/color", "items/weight"]);

  table.sort.apply([{ key: 0, ascending: true }]);
  table.sort.load("fields");
  table.getDataBodyRange().load("values");
  const filter = table.columns.getItemAt(0).filter;
  filter.applyValuesFilter(["one", "two"]);
  filter.applyCustomFilter(">0", "<10", "And");
  filter.clear();

  const chart = worksheet.charts.add("ColumnClustered", source, "Auto");
  chart.name = "Hermes_Chart_00000000_0000_4000_8000_000000000000";
  chart.chartType;
  chart.load(["name", "chartType", "title/text"]);
  chart.title.text = "Title";
  const validation = target.dataValidation;
  const listRule: Excel.DataValidationRule = { list: { inCellDropDown: true, source: "one,two" } };
  validation.rule = listRule;
  const prompt: Excel.DataValidationPrompt = { showPrompt: true, title: "Pick", message: "Choose" };
  validation.prompt = prompt;

  filter.applyCustomFilter(">0");
  filter.applyCustomFilter(">0", "<10");
  filter.applyCustomFilter(">0", undefined, "And");
  const cellValue = target.conditionalFormats.add("CellValue").cellValue;
  cellValue.rule = { operator: "GreaterThan", formula1: "0" };
  cellValue.format.fill.color = "#00ff00";
  cellValue.format.font.color = "#000000";
  const containsText = target.conditionalFormats.add("ContainsText").textComparison;
  containsText.rule = { operator: "Contains", text: "ok" };
  containsText.format.fill.color = "#00ff00";
  containsText.format.font.color = "#000000";
  target.conditionalFormats.load("items/id");

  worksheet.freezePanes.freezeRows(2);
  worksheet.freezePanes.freezeColumns(3);
  const frozen = worksheet.freezePanes.getLocationOrNullObject();
  frozen.load(["isNullObject", "address", "rowCount", "columnCount"]);
  worksheet.freezePanes.unfreeze();
  worksheet.tables.add(source, true);
  worksheet.name = "Renamed";
  const copiedWorksheet = worksheet.copy("End");
  copiedWorksheet.name = "Copied";
  worksheet.copy("End");
  workbook.worksheets.add("NewSheet");
  workbook.application.calculate("Full");

  const comments = worksheet.comments;
  const count: OfficeExtension.ClientResult<number> = comments.getCount();
  void count.value;
  const comment = comments.add(target, "note");
  comment.content = "updated";
  const location = comment.getLocation();
  location.load("address");
}

void task11C1OfficeJsTypeProbe;
