(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) throw new Error((value && value.error) || "Sheets Bridge 执行失败");
    return value;
  }

  function execute(toolCalls) {
    return new Promise(function (resolve, reject) {
      Asc.scope.copilotSheetCalls = toolCalls;
      Asc.plugin.info.recalculate = true;

      Asc.plugin.callCommand(
        function () {
          try {
            var calls = Asc.scope.copilotSheetCalls || [];
            var results = [];
            var changed = 0;
            var mutating = false;
            var mutatingNames = {
              sheets_set_values: true,
              sheets_set_formula: true,
              sheets_replace_text: true,
              sheets_format_range: true,
              sheets_add_sheet: true,
              sheets_rename_sheet: true,
              sheets_delete_sheet: true,
              sheets_add_chart: true,
            };

            function getArgs(call) {
              return (call && (call.arguments || call.args)) || {};
            }

            function getSheet(name) {
              if (!name) return Api.GetActiveSheet();
              var sheets = Api.GetSheets();
              for (var index = 0; index < sheets.length; index += 1) {
                if (String(sheets[index].GetName()) === String(name)) return sheets[index];
              }
              throw new Error("找不到工作表：" + name);
            }

            function color(value) {
              var hex = String(value || "").replace("#", "");
              if (!/^[0-9a-fA-F]{6}$/.test(hex)) throw new Error("颜色必须是 #RRGGBB 格式");
              return Api.CreateColorFromRGB(parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16));
            }

            function getRange(sheet, reference) {
              if (String(reference).toLowerCase() === "selection") return Api.GetSelection();
              return sheet.GetRange(String(reference));
            }

            for (var callIndex = 0; callIndex < calls.length; callIndex += 1) {
              if (mutatingNames[calls[callIndex].name]) {
                mutating = true;
                break;
              }
            }
            if (mutating && typeof Api.CreateNewHistoryPoint === "function") Api.CreateNewHistoryPoint();

            for (var callIndex = 0; callIndex < calls.length; callIndex += 1) {
              var call = calls[callIndex];
              var args = getArgs(call);

              switch (call.name) {
                case "sheets_inspect": {
                  var sheets = Api.GetSheets();
                  var sheetData = [];
                  var maxCells = Math.min(5000, Math.max(100, Number(args.maxCells) || 1200));
                  for (var sheetIndex = 0; sheetIndex < sheets.length; sheetIndex += 1) {
                    var sheet = sheets[sheetIndex];
                    var used = sheet.GetUsedRange();
                    var values = used ? used.GetValue() : [];
                    if (Array.isArray(values) && Array.isArray(values[0])) {
                      var rowLimit = Math.max(1, Math.floor(maxCells / Math.max(1, values[0].length)));
                      values = values.slice(0, rowLimit);
                    }
                    sheetData.push({
                      name: String(sheet.GetName()),
                      usedRange: used && typeof used.GetAddress === "function" ? used.GetAddress() : null,
                      values: values,
                    });
                  }
                  var selection = Api.GetSelection();
                  results.push({
                    name: call.name,
                    activeSheet: Api.GetActiveSheet().GetName(),
                    selection: selection ? {
                      address: selection.GetAddress(true, true, "xlA1", false),
                      values: selection.GetValue(),
                    } : null,
                    sheets: sheetData,
                  });
                  break;
                }

                case "sheets_set_values": {
                  if (!args.range) throw new Error("sheets_set_values.range 不能为空");
                  if (args.values === undefined) throw new Error("sheets_set_values.values 不能为空");
                  var sheet = getSheet(args.sheet);
                  var range = getRange(sheet, args.range);
                  if (!range) throw new Error("无效单元格区域：" + args.range);
                  var accepted = range.SetValue(args.values);
                  changed += 1;
                  results.push({ name: call.name, sheet: sheet.GetName(), range: String(args.range), accepted: Boolean(accepted) });
                  break;
                }

                case "sheets_set_formula": {
                  if (!args.range || args.formula === undefined) throw new Error("sheets_set_formula 需要 range 和 formula");
                  var sheet = getSheet(args.sheet);
                  var formula = String(args.formula);
                  if (formula.charAt(0) !== "=") formula = "=" + formula;
                  var formulaRange = getRange(sheet, args.range);
                  if (!formulaRange) throw new Error("无效单元格区域：" + args.range);
                  var accepted = typeof formulaRange.SetFormula === "function"
                    ? formulaRange.SetFormula(formula)
                    : formulaRange.SetValue(formula);
                  changed += 1;
                  results.push({ name: call.name, sheet: sheet.GetName(), range: String(args.range), accepted: Boolean(accepted) });
                  break;
                }

                case "sheets_replace_text": {
                  if (!args.search) throw new Error("sheets_replace_text.search 不能为空");
                  var sheet = getSheet(args.sheet);
                  var range = args.range ? getRange(sheet, args.range) : sheet.GetUsedRange();
                  if (!range) throw new Error("工作表中没有可替换的区域");
                  var replaced = range.Replace(String(args.search), String(args.replace || ""));
                  changed += replaced ? 1 : 0;
                  results.push({ name: call.name, sheet: sheet.GetName(), replaced: Boolean(replaced) });
                  break;
                }

                case "sheets_format_range": {
                  if (!args.range) throw new Error("sheets_format_range.range 不能为空");
                  var sheet = getSheet(args.sheet);
                  var range = getRange(sheet, args.range);
                  if (!range) throw new Error("无效单元格区域：" + args.range);
                  if (args.fontSize !== undefined) range.SetFontSize(Math.max(1, Number(args.fontSize)));
                  if (args.fontName) range.SetFontName(String(args.fontName));
                  if (args.bold !== undefined) range.SetBold(Boolean(args.bold));
                  if (args.italic !== undefined) range.SetItalic(Boolean(args.italic));
                  if (args.underline !== undefined) range.SetUnderline(Boolean(args.underline));
                  if (args.fontColor) range.SetFontColor(color(args.fontColor));
                  if (args.fillColor) range.SetFillColor(color(args.fillColor));
                  if (args.horizontalAlign) range.SetAlignHorizontal(String(args.horizontalAlign));
                  if (args.verticalAlign) range.SetAlignVertical(String(args.verticalAlign));
                  if (args.numberFormat) range.SetNumberFormat(String(args.numberFormat));
                  if (args.wrap !== undefined) range.SetWrap(Boolean(args.wrap));
                  if (args.columnWidth !== undefined) range.SetColumnWidth(Number(args.columnWidth));
                  if (args.rowHeight !== undefined) range.SetRowHeight(Number(args.rowHeight));
                  changed += 1;
                  results.push({ name: call.name, sheet: sheet.GetName(), range: String(args.range) });
                  break;
                }

                case "sheets_add_sheet": {
                  if (!args.name) throw new Error("sheets_add_sheet.name 不能为空");
                  var created = Api.AddSheet(String(args.name));
                  if (!created) throw new Error("创建工作表失败，名称可能已存在");
                  changed += 1;
                  results.push({ name: call.name, sheet: created.GetName() });
                  break;
                }

                case "sheets_rename_sheet": {
                  if (!args.newName) throw new Error("sheets_rename_sheet.newName 不能为空");
                  var sheet = getSheet(args.sheet);
                  sheet.SetName(String(args.newName));
                  changed += 1;
                  results.push({ name: call.name, oldName: args.sheet || "active", newName: String(args.newName) });
                  break;
                }

                case "sheets_delete_sheet": {
                  if (Api.GetSheets().length <= 1) throw new Error("工作簿至少需要保留一个工作表");
                  var sheet = getSheet(args.sheet);
                  var deletedName = sheet.GetName();
                  var deleted = sheet.Delete();
                  if (!deleted) throw new Error("删除工作表失败");
                  changed += 1;
                  results.push({ name: call.name, deletedSheet: deletedName });
                  break;
                }

                case "sheets_add_chart": {
                  if (!args.range) throw new Error("sheets_add_chart.range 不能为空");
                  var sheet = getSheet(args.sheet);
                  var source = String(args.range);
                  if (source.indexOf("!") === -1) source = "'" + String(sheet.GetName()).replace(/'/g, "''") + "'!" + source;
                  var chart = sheet.AddChart(
                    source,
                    Boolean(args.inRows),
                    String(args.type || "bar"),
                    Math.min(48, Math.max(1, Number(args.style) || 2)),
                    Math.max(30, Number(args.widthMm) || 120) * 36000,
                    Math.max(20, Number(args.heightMm) || 70) * 36000,
                    Math.max(0, Number(args.fromColumn) || 0),
                    0,
                    Math.max(0, Number(args.fromRow) || 0),
                    0
                  );
                  if (!chart) throw new Error("创建图表失败");
                  if (args.title) chart.SetTitle(String(args.title), Number(args.titleFontSize) || 13);
                  changed += 1;
                  results.push({ name: call.name, sheet: sheet.GetName(), range: source, type: String(args.type || "bar") });
                  break;
                }

                default:
                  throw new Error("Sheets Bridge 不支持工具：" + call.name);
              }
            }

            return JSON.stringify({ ok: true, editorType: "cell", changed: changed, needsSave: mutating, results: results });
          } catch (error) {
            return JSON.stringify({ ok: false, error: error && error.message ? error.message : String(error) });
          }
        },
        false,
        true,
        function (rawResult) {
          try {
            resolve(parseResult(rawResult));
          } catch (error) {
            reject(error);
          }
        },
      );
    });
  }

  window.AICopilotBridges.cell = {
    execute: execute,
    inspect: function () {
      return execute([{ name: "sheets_inspect", arguments: { maxCells: 1200 } }]);
    },
  };
})();
