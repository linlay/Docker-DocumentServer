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
            var EMU_PER_MM = 36000;
            var EMU_PER_POINT = 12700;
            var mutatingNames = {
              sheets_set_values: true,
              sheets_set_formula: true,
              sheets_replace_text: true,
              sheets_format_range: true,
              sheets_add_sheet: true,
              sheets_rename_sheet: true,
              sheets_delete_sheet: true,
              sheets_add_chart: true,
              sheets_update_chart: true,
              sheets_delete_chart: true,
            };

            function getArgs(call) {
              return (call && (call.arguments || call.args)) || {};
            }

            function hasOwn(value, key) {
              return Boolean(value) && Object.prototype.hasOwnProperty.call(value, key);
            }

            function isObject(value) {
              return Boolean(value) && typeof value === "object" && !Array.isArray(value);
            }

            function asFinite(value, fallback) {
              var number = Number(value);
              return isFinite(number) ? number : fallback;
            }

            function clamp(value, minimum, maximum) {
              return Math.min(maximum, Math.max(minimum, value));
            }

            function normalizeChartType(value, fallback) {
              var type = String(value || fallback || "bar");
              var aliases = {
                line: "lineNormal",
                lineMarker: "lineNormalMarker",
                stackedBar: "barStacked",
                stackedBarPercent: "barStackedPercent",
                stackedLine: "lineStacked",
                stackedLinePercent: "lineStackedPercent",
                column: "bar",
              };
              return aliases[type] || type;
            }

            function mmToEmu(value) {
              return Math.round(asFinite(value, 0) * EMU_PER_MM);
            }

            function emuToMm(value) {
              return typeof value === "number" && isFinite(value)
                ? Math.round(value / EMU_PER_MM * 1000) / 1000
                : null;
            }

            function safeCall(value, method) {
              if (!value || typeof value[method] !== "function") return null;
              try {
                return value[method].apply(value, Array.prototype.slice.call(arguments, 2));
              } catch (error) {
                return null;
              }
            }

            function serialized(value) {
              if (!value || typeof value.ToJSON !== "function") return null;
              try {
                var raw = value.ToJSON();
                if (typeof raw !== "string") return raw;
                try {
                  return JSON.parse(raw);
                } catch (error) {
                  return raw;
                }
              } catch (error) {
                return null;
              }
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
              var red = parseInt(hex.slice(0, 2), 16);
              var green = parseInt(hex.slice(2, 4), 16);
              var blue = parseInt(hex.slice(4, 6), 16);
              if (typeof Api.CreateRGBColor === "function") return Api.CreateRGBColor(red, green, blue);
              if (typeof Api.CreateColorFromRGB === "function") return Api.CreateColorFromRGB(red, green, blue);
              return Api.RGB(red, green, blue);
            }

            function createGradientStops(stops) {
              if (!Array.isArray(stops) || stops.length < 2 || stops.length > 16) {
                throw new Error("渐变色 stops 必须包含 2 到 16 个颜色节点");
              }
              if (typeof Api.CreateGradientStop !== "function") throw new Error("当前 ONLYOFFICE 版本不支持渐变节点");
              return stops.map(function (stop) {
                if (!isObject(stop) || !stop.color) throw new Error("每个渐变节点都需要 color");
                return Api.CreateGradientStop(
                  color(stop.color),
                  Math.round(clamp(asFinite(stop.position, 0), 0, 100) * 1000)
                );
              });
            }

            function createFill(spec) {
              if (typeof spec === "string") return Api.CreateSolidFill(color(spec));
              if (!isObject(spec)) throw new Error("fill 必须是颜色字符串或填充对象");
              if (spec.raw !== undefined) {
                if (typeof Api.FromJSON !== "function") throw new Error("当前 ONLYOFFICE 版本不支持 fill.raw");
                var fillJson = typeof spec.raw === "string" ? spec.raw : JSON.stringify(spec.raw);
                var restored = Api.FromJSON(fillJson);
                if (!restored) throw new Error("fill.raw 无法恢复为填充对象");
                return restored;
              }
              var type = String(spec.type || "solid");
              if (type === "none" || type === "nofill") return Api.CreateNoFill();
              if (type === "solid") return Api.CreateSolidFill(color(spec.color));
              if (type === "linearGradient") {
                if (typeof Api.CreateLinearGradientFill !== "function") throw new Error("当前 ONLYOFFICE 版本不支持线性渐变");
                return Api.CreateLinearGradientFill(
                  createGradientStops(spec.stops),
                  Math.round(asFinite(spec.angleDeg, 0) * 60000)
                );
              }
              if (type === "radialGradient") {
                if (typeof Api.CreateRadialGradientFill !== "function") throw new Error("当前 ONLYOFFICE 版本不支持径向渐变");
                return Api.CreateRadialGradientFill(createGradientStops(spec.stops));
              }
              if (type === "pattern") {
                if (typeof Api.CreatePatternFill !== "function") throw new Error("当前 ONLYOFFICE 版本不支持图案填充");
                return Api.CreatePatternFill(
                  String(spec.pattern || "pct20"),
                  color(spec.backgroundColor),
                  color(spec.foregroundColor)
                );
              }
              throw new Error("不支持的填充类型：" + type);
            }

            function createStroke(spec) {
              if (spec === false || (isObject(spec) && (spec.type === "none" || spec.enabled === false))) {
                return Api.CreateStroke(0, Api.CreateNoFill());
              }
              if (!isObject(spec)) throw new Error("line 必须是线条对象");
              if (spec.raw !== undefined) {
                if (typeof Api.FromJSON !== "function") throw new Error("当前 ONLYOFFICE 版本不支持 line.raw");
                var lineJson = typeof spec.raw === "string" ? spec.raw : JSON.stringify(spec.raw);
                var restored = Api.FromJSON(lineJson);
                if (!restored) throw new Error("line.raw 无法恢复为线条对象");
                return restored;
              }
              var fill = spec.fill !== undefined
                ? createFill(spec.fill)
                : (spec.color ? Api.CreateSolidFill(color(spec.color)) : Api.CreateNoFill());
              return Api.CreateStroke(Math.max(0, asFinite(spec.widthPt, 1)) * EMU_PER_POINT, fill);
            }

            function getRange(sheet, reference) {
              if (String(reference).toLowerCase() === "selection") return Api.GetSelection();
              return sheet.GetRange(String(reference));
            }

            function qualifyRange(sheet, reference) {
              if (String(reference).toLowerCase() === "selection") {
                var selection = Api.GetSelection();
                if (!selection || typeof selection.GetAddress !== "function") throw new Error("当前选区不可用");
                reference = selection.GetAddress(true, true, "xlA1", false);
              }
              var source = String(reference);
              if (source.indexOf("!") === -1) {
                source = "'" + String(sheet.GetName()).replace(/'/g, "''") + "'!" + source;
              }
              return source;
            }

            function sheetCharts(sheet) {
              return typeof sheet.GetAllCharts === "function" ? sheet.GetAllCharts() : (sheet.charts || []);
            }

            function chartIndexOnSheet(sheet, chart) {
              var charts = sheetCharts(sheet);
              var directIndex = charts.indexOf(chart);
              if (directIndex >= 0) return directIndex;
              var targetId = safeCall(chart, "GetInternalId");
              if (targetId !== null && targetId !== undefined) {
                for (var idIndex = 0; idIndex < charts.length; idIndex += 1) {
                  if (String(safeCall(charts[idIndex], "GetInternalId")) === String(targetId)) return idIndex;
                }
              }
              return Math.max(0, charts.length - 1);
            }

            function describeChart(chart, index, includeRaw) {
              var series = safeCall(chart, "GetAllSeries");
              return {
                chartIndex: index,
                name: safeCall(chart, "GetName"),
                type: safeCall(chart, "GetChartType"),
                title: safeCall(chart, "GetTitle"),
                widthMm: emuToMm(safeCall(chart, "GetWidth")),
                heightMm: emuToMm(safeCall(chart, "GetHeight")),
                rotationDeg: safeCall(chart, "GetRotation"),
                series: Array.isArray(series) ? series.map(function (item, seriesIndex) {
                  return {
                    index: seriesIndex,
                    chartType: safeCall(item, "GetChartType"),
                    raw: includeRaw ? serialized(item) : undefined,
                  };
                }) : [],
                raw: includeRaw ? serialized(chart) : undefined,
              };
            }

            function resolveChart(args) {
              var sheet = getSheet(args.sheet);
              var charts = sheetCharts(sheet);
              for (var index = 0; index < charts.length; index += 1) {
                if (hasOwn(args, "chartIndex") && index !== Number(args.chartIndex)) continue;
                if (args.name && String(safeCall(charts[index], "GetName")) !== String(args.name)) continue;
                if (hasOwn(args, "chartIndex") || args.name) return { sheet: sheet, chart: charts[index], index: index };
              }
              throw new Error("找不到目标图表；请提供有效的 chartIndex 或 name");
            }

            function applyChartAxis(chart, axis, horizontal) {
              if (!isObject(axis)) return;
              var titleMethod = horizontal ? "SetHorAxisTitle" : "SetVerAxisTitle";
              var labelsMethod = horizontal ? "SetHorAxisLabelsFontSize" : "SetVertAxisLabelsFontSize";
              var orientationMethod = horizontal ? "SetHorAxisOrientation" : "SetVerAxisOrientation";
              var majorTickMethod = horizontal ? "SetHorAxisMajorTickMark" : "SetVertAxisMajorTickMark";
              var minorTickMethod = horizontal ? "SetHorAxisMinorTickMark" : "SetVertAxisMinorTickMark";
              var labelPositionMethod = horizontal ? "SetHorAxisTickLabelPosition" : "SetVertAxisTickLabelPosition";
              if (hasOwn(axis, "title") && typeof chart[titleMethod] === "function") {
                chart[titleMethod](String(axis.title), asFinite(axis.titleFontSize, 11));
              }
              if (hasOwn(axis, "labelsFontSize") && typeof chart[labelsMethod] === "function") {
                chart[labelsMethod](asFinite(axis.labelsFontSize, 10));
              }
              if (hasOwn(axis, "normalOrder") && typeof chart[orientationMethod] === "function") {
                chart[orientationMethod](Boolean(axis.normalOrder));
              }
              if (hasOwn(axis, "majorTickMark") && typeof chart[majorTickMethod] === "function") {
                chart[majorTickMethod](String(axis.majorTickMark));
              }
              if (hasOwn(axis, "minorTickMark") && typeof chart[minorTickMethod] === "function") {
                chart[minorTickMethod](String(axis.minorTickMark));
              }
              if (hasOwn(axis, "tickLabelPosition") && typeof chart[labelPositionMethod] === "function") {
                chart[labelPositionMethod](String(axis.tickLabelPosition));
              }
              if (hasOwn(axis, "numberFormat") && typeof chart.SetAxieNumFormat === "function") {
                chart.SetAxieNumFormat(String(axis.numberFormat), String(axis.position || (horizontal ? "bottom" : "left")));
              }
            }

            function applySeriesUpdates(chart, updates) {
              if (!Array.isArray(updates)) return;
              for (var updateIndex = 0; updateIndex < updates.length; updateIndex += 1) {
                var update = updates[updateIndex] || {};
                var seriesIndex = Number(update.index);
                if (!Number.isInteger(seriesIndex) || seriesIndex < 0) throw new Error("seriesUpdates.index 必须是从 0 开始的整数");
                if (hasOwn(update, "type") && typeof chart.GetSeries === "function") {
                  var typedSeries = chart.GetSeries(seriesIndex);
                  if (!typedSeries || typeof typedSeries.ChangeChartType !== "function") {
                    throw new Error("当前 ONLYOFFICE 版本不支持修改系列图表类型");
                  }
                  typedSeries.ChangeChartType(normalizeChartType(update.type, "bar"));
                }
                if (hasOwn(update, "name") && typeof chart.SetSeriaName === "function") {
                  chart.SetSeriaName(String(update.name), seriesIndex);
                }
                if (hasOwn(update, "valuesRange") && typeof chart.SetSeriaValues === "function") {
                  chart.SetSeriaValues(String(update.valuesRange), seriesIndex);
                }
                if (hasOwn(update, "xValuesRange") && typeof chart.SetSeriaXValues === "function") {
                  chart.SetSeriaXValues(String(update.xValuesRange), seriesIndex);
                }
                if (hasOwn(update, "numberFormat") && typeof chart.SetSeriaNumFormat === "function") {
                  chart.SetSeriaNumFormat(String(update.numberFormat), seriesIndex);
                }
                if (hasOwn(update, "fill") && typeof chart.SetSeriesFill === "function") {
                  chart.SetSeriesFill(createFill(update.fill), seriesIndex, Boolean(update.allSeries));
                }
                if (hasOwn(update, "line") && typeof chart.SetSeriesOutLine === "function") {
                  chart.SetSeriesOutLine(createStroke(update.line), seriesIndex, Boolean(update.allSeries));
                }
                if (Array.isArray(update.points)) {
                  for (var pointIndex = 0; pointIndex < update.points.length; pointIndex += 1) {
                    var point = update.points[pointIndex] || {};
                    var targetPoint = Number(point.index);
                    if (!Number.isInteger(targetPoint) || targetPoint < 0) throw new Error("points.index 必须是从 0 开始的整数");
                    if (hasOwn(point, "fill") && typeof chart.SetDataPointFill === "function") {
                      chart.SetDataPointFill(createFill(point.fill), seriesIndex, targetPoint, Boolean(point.allSeries));
                    }
                    if (hasOwn(point, "line") && typeof chart.SetDataPointOutLine === "function") {
                      chart.SetDataPointOutLine(createStroke(point.line), seriesIndex, targetPoint, Boolean(point.allSeries));
                    }
                    if (hasOwn(point, "markerFill") && typeof chart.SetMarkerFill === "function") {
                      chart.SetMarkerFill(createFill(point.markerFill), seriesIndex, targetPoint, Boolean(point.allMarkers));
                    }
                    if (hasOwn(point, "markerLine") && typeof chart.SetMarkerOutLine === "function") {
                      chart.SetMarkerOutLine(createStroke(point.markerLine), seriesIndex, targetPoint, Boolean(point.allMarkers));
                    }
                    if (hasOwn(point, "numberFormat") && typeof chart.SetDataPointNumFormat === "function") {
                      chart.SetDataPointNumFormat(String(point.numberFormat), seriesIndex, targetPoint);
                    }
                    if (isObject(point.dataLabels) && typeof chart.SetShowPointDataLabel === "function") {
                      chart.SetShowPointDataLabel(
                        seriesIndex,
                        targetPoint,
                        Boolean(point.dataLabels.showSeriesName),
                        Boolean(point.dataLabels.showCategoryName),
                        Boolean(point.dataLabels.showValue),
                        Boolean(point.dataLabels.showPercent)
                      );
                    }
                  }
                }
              }
            }

            function setChartGridline(chart, method, value) {
              if (!isObject(value) || !hasOwn(value, "line")) return;
              if (typeof chart[method] === "function") chart[method](createStroke(value.line));
            }

            function applyChartOptions(chart, args, skipPosition) {
              if (!skipPosition && (hasOwn(args, "fromColumn") || hasOwn(args, "fromRow") || hasOwn(args, "columnOffsetMm") || hasOwn(args, "rowOffsetMm"))) {
                if (typeof chart.SetPosition === "function") {
                  chart.SetPosition(
                    Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                    mmToEmu(args.columnOffsetMm || 0),
                    Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                    mmToEmu(args.rowOffsetMm || 0)
                  );
                }
              }
              if (hasOwn(args, "widthMm") || hasOwn(args, "heightMm")) {
                var width = Math.max(20, asFinite(args.widthMm, emuToMm(safeCall(chart, "GetWidth")) || 120));
                var height = Math.max(20, asFinite(args.heightMm, emuToMm(safeCall(chart, "GetHeight")) || 70));
                if (typeof chart.SetSize === "function") chart.SetSize(mmToEmu(width), mmToEmu(height));
              }
              if (hasOwn(args, "rotationDeg") && typeof chart.SetRotation === "function") {
                chart.SetRotation(asFinite(args.rotationDeg, 0));
              }
              if (hasOwn(args, "name") && typeof chart.SetName === "function") chart.SetName(String(args.name));
              if (hasOwn(args, "style") && typeof chart.ApplyChartStyle === "function") {
                chart.ApplyChartStyle(clamp(Math.round(asFinite(args.style, 2)), 1, 48));
              }
              if (hasOwn(args, "title") && typeof chart.SetTitle === "function") {
                chart.SetTitle(String(args.title), asFinite(args.titleFontSize, 13));
              }
              if (hasOwn(args, "fill") && typeof chart.Fill === "function") chart.Fill(createFill(args.fill));
              if (hasOwn(args, "line") && typeof chart.SetOutLine === "function") chart.SetOutLine(createStroke(args.line));
              if (hasOwn(args, "plotAreaFill") && typeof chart.SetPlotAreaFill === "function") chart.SetPlotAreaFill(createFill(args.plotAreaFill));
              if (hasOwn(args, "plotAreaLine") && typeof chart.SetPlotAreaOutLine === "function") chart.SetPlotAreaOutLine(createStroke(args.plotAreaLine));
              if (hasOwn(args, "titleFill") && typeof chart.SetTitleFill === "function") chart.SetTitleFill(createFill(args.titleFill));
              if (hasOwn(args, "titleLine") && typeof chart.SetTitleOutLine === "function") chart.SetTitleOutLine(createStroke(args.titleLine));
              if (isObject(args.legend)) {
                if (hasOwn(args.legend, "position") && typeof chart.SetLegendPos === "function") chart.SetLegendPos(String(args.legend.position));
                if (hasOwn(args.legend, "fontSize") && typeof chart.SetLegendFontSize === "function") chart.SetLegendFontSize(asFinite(args.legend.fontSize, 10));
                if (hasOwn(args.legend, "fill") && typeof chart.SetLegendFill === "function") chart.SetLegendFill(createFill(args.legend.fill));
                if (hasOwn(args.legend, "line") && typeof chart.SetLegendOutLine === "function") chart.SetLegendOutLine(createStroke(args.legend.line));
              }
              applyChartAxis(chart, args.horizontalAxis, true);
              applyChartAxis(chart, args.verticalAxis, false);
              if (isObject(args.dataLabels) && typeof chart.SetShowDataLabels === "function") {
                chart.SetShowDataLabels(
                  Boolean(args.dataLabels.showSeriesName),
                  Boolean(args.dataLabels.showCategoryName),
                  Boolean(args.dataLabels.showValue),
                  Boolean(args.dataLabels.showPercent)
                );
              }
              if (isObject(args.gridlines)) {
                setChartGridline(chart, "SetMajorHorizontalGridlines", args.gridlines.majorHorizontal);
                setChartGridline(chart, "SetMinorHorizontalGridlines", args.gridlines.minorHorizontal);
                setChartGridline(chart, "SetMajorVerticalGridlines", args.gridlines.majorVertical);
                setChartGridline(chart, "SetMinorVerticalGridlines", args.gridlines.minorVertical);
              }
              if (hasOwn(args, "categoryRange") && typeof chart.SetCatFormula === "function") {
                chart.SetCatFormula(String(args.categoryRange));
              }
              if (Array.isArray(args.addSeries) && typeof chart.AddSeria === "function") {
                for (var addIndex = 0; addIndex < args.addSeries.length; addIndex += 1) {
                  var added = args.addSeries[addIndex] || {};
                  if (!added.name || !added.valuesRange) throw new Error("addSeries 需要 name 和 valuesRange");
                  chart.AddSeria(String(added.name), String(added.valuesRange), added.xValuesRange ? String(added.xValuesRange) : undefined);
                }
              }
              applySeriesUpdates(chart, args.seriesUpdates);
              if (Array.isArray(args.removeSeries) && typeof chart.RemoveSeria === "function") {
                var ordered = args.removeSeries.slice().map(Number).sort(function (left, right) { return right - left; });
                for (var removeIndex = 0; removeIndex < ordered.length; removeIndex += 1) chart.RemoveSeria(ordered[removeIndex]);
              }
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
                      charts: sheetCharts(sheet).length,
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
                  var valuesSheet = getSheet(args.sheet);
                  var range = getRange(valuesSheet, args.range);
                  if (!range) throw new Error("无效单元格区域：" + args.range);
                  var accepted = range.SetValue(args.values);
                  changed += 1;
                  results.push({ name: call.name, sheet: valuesSheet.GetName(), range: String(args.range), accepted: Boolean(accepted) });
                  break;
                }

                case "sheets_set_formula": {
                  if (!args.range || args.formula === undefined) throw new Error("sheets_set_formula 需要 range 和 formula");
                  var formulaSheet = getSheet(args.sheet);
                  var formula = String(args.formula);
                  if (formula.charAt(0) !== "=") formula = "=" + formula;
                  var formulaRange = getRange(formulaSheet, args.range);
                  if (!formulaRange) throw new Error("无效单元格区域：" + args.range);
                  var formulaAccepted = typeof formulaRange.SetFormula === "function"
                    ? formulaRange.SetFormula(formula)
                    : formulaRange.SetValue(formula);
                  changed += 1;
                  results.push({ name: call.name, sheet: formulaSheet.GetName(), range: String(args.range), accepted: Boolean(formulaAccepted) });
                  break;
                }

                case "sheets_replace_text": {
                  if (!args.search) throw new Error("sheets_replace_text.search 不能为空");
                  var replaceSheet = getSheet(args.sheet);
                  var replaceRange = args.range ? getRange(replaceSheet, args.range) : replaceSheet.GetUsedRange();
                  if (!replaceRange) throw new Error("工作表中没有可替换的区域");
                  var replaced = replaceRange.Replace(String(args.search), String(args.replace || ""));
                  changed += replaced ? 1 : 0;
                  results.push({ name: call.name, sheet: replaceSheet.GetName(), replaced: Boolean(replaced) });
                  break;
                }

                case "sheets_format_range": {
                  if (!args.range) throw new Error("sheets_format_range.range 不能为空");
                  var formatSheet = getSheet(args.sheet);
                  var formatRange = getRange(formatSheet, args.range);
                  if (!formatRange) throw new Error("无效单元格区域：" + args.range);
                  if (args.fontSize !== undefined) formatRange.SetFontSize(Math.max(1, Number(args.fontSize)));
                  if (args.fontName) formatRange.SetFontName(String(args.fontName));
                  if (args.bold !== undefined) formatRange.SetBold(Boolean(args.bold));
                  if (args.italic !== undefined) formatRange.SetItalic(Boolean(args.italic));
                  if (args.underline !== undefined) formatRange.SetUnderline(Boolean(args.underline));
                  if (args.fontColor) formatRange.SetFontColor(color(args.fontColor));
                  if (args.fillColor) formatRange.SetFillColor(color(args.fillColor));
                  if (args.horizontalAlign) formatRange.SetAlignHorizontal(String(args.horizontalAlign));
                  if (args.verticalAlign) formatRange.SetAlignVertical(String(args.verticalAlign));
                  if (args.numberFormat) formatRange.SetNumberFormat(String(args.numberFormat));
                  if (args.wrap !== undefined) formatRange.SetWrap(Boolean(args.wrap));
                  if (args.columnWidth !== undefined) formatRange.SetColumnWidth(Number(args.columnWidth));
                  if (args.rowHeight !== undefined) formatRange.SetRowHeight(Number(args.rowHeight));
                  changed += 1;
                  results.push({ name: call.name, sheet: formatSheet.GetName(), range: String(args.range) });
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
                  var renameSheet = getSheet(args.sheet);
                  renameSheet.SetName(String(args.newName));
                  changed += 1;
                  results.push({ name: call.name, oldName: args.sheet || "active", newName: String(args.newName) });
                  break;
                }

                case "sheets_delete_sheet": {
                  if (Api.GetSheets().length <= 1) throw new Error("工作簿至少需要保留一个工作表");
                  var deleteSheet = getSheet(args.sheet);
                  var deletedName = deleteSheet.GetName();
                  var sheetDeleted = deleteSheet.Delete();
                  if (!sheetDeleted) throw new Error("删除工作表失败");
                  changed += 1;
                  results.push({ name: call.name, deletedSheet: deletedName });
                  break;
                }

                case "sheets_inspect_charts": {
                  var inspectSheets = args.sheet ? [getSheet(args.sheet)] : Api.GetSheets();
                  var maxCharts = clamp(Math.round(asFinite(args.maxCharts, 100)), 1, 500);
                  var inspectedSheets = [];
                  var chartCount = 0;
                  for (var inspectSheetIndex = 0; inspectSheetIndex < inspectSheets.length && chartCount < maxCharts; inspectSheetIndex += 1) {
                    var inspectedSheet = inspectSheets[inspectSheetIndex];
                    var inspectedCharts = sheetCharts(inspectedSheet);
                    var chartItems = [];
                    for (var inspectChartIndex = 0; inspectChartIndex < inspectedCharts.length && chartCount < maxCharts; inspectChartIndex += 1) {
                      chartItems.push(describeChart(inspectedCharts[inspectChartIndex], inspectChartIndex, args.includeRaw !== false));
                      chartCount += 1;
                    }
                    inspectedSheets.push({ sheet: inspectedSheet.GetName(), charts: chartItems });
                  }
                  results.push({
                    name: call.name,
                    sheets: inspectedSheets,
                    chartCount: chartCount,
                    truncated: chartCount >= maxCharts,
                  });
                  break;
                }

                case "sheets_add_chart": {
                  if (!args.range) throw new Error("sheets_add_chart.range 不能为空");
                  var chartSheet = getSheet(args.sheet);
                  var source = qualifyRange(chartSheet, args.range);
                  var chart = chartSheet.AddChart(
                    source,
                    Boolean(args.inRows),
                    normalizeChartType(args.type, "bar"),
                    clamp(Math.round(asFinite(args.style, 2)), 1, 48),
                    Math.max(30, asFinite(args.widthMm, 120)) * EMU_PER_MM,
                    Math.max(20, asFinite(args.heightMm, 70)) * EMU_PER_MM,
                    Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                    mmToEmu(args.columnOffsetMm || 0),
                    Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                    mmToEmu(args.rowOffsetMm || 0)
                  );
                  if (!chart) throw new Error("创建图表失败");
                  applyChartOptions(chart, args, true);
                  changed += 1;
                  results.push({
                    name: call.name,
                    sheet: chartSheet.GetName(),
                    range: source,
                    type: normalizeChartType(args.type, "bar"),
                    chart: describeChart(chart, chartIndexOnSheet(chartSheet, chart), Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "sheets_update_chart": {
                  var resolvedChart = resolveChart(args);
                  applyChartOptions(resolvedChart.chart, args, false);
                  changed += 1;
                  results.push({
                    name: call.name,
                    sheet: resolvedChart.sheet.GetName(),
                    chart: describeChart(resolvedChart.chart, resolvedChart.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "sheets_delete_chart": {
                  var chartToDelete = resolveChart(args);
                  if (typeof chartToDelete.chart.Delete !== "function") throw new Error("当前 ONLYOFFICE 版本不支持删除图表");
                  var chartDeleted = chartToDelete.chart.Delete();
                  if (chartDeleted === false) throw new Error("删除图表失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    sheet: chartToDelete.sheet.GetName(),
                    deletedChartIndex: chartToDelete.index,
                  });
                  break;
                }

                default:
                  throw new Error("Sheets Bridge 不支持工具：" + call.name);
              }
            }

            return JSON.stringify({
              ok: true,
              editorType: "cell",
              changed: changed,
              needsSave: mutating,
              results: results,
            });
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
