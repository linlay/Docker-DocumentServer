(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) throw new Error((value && value.error) || "Sheets Bridge 执行失败");
    return value;
  }

  function executeMacroTool(toolCalls) {
    return new Promise(function (resolve, reject) {
      if (toolCalls.length !== 1) {
        reject(new Error("宏工具必须单独调用，不能与普通工作表命令混批"));
        return;
      }
      const call = toolCalls[0] || {};
      const args = call.arguments || call.args || {};
      let method;
      let params;
      let needsSave = false;
      if (call.name === "sheets_inspect_macros") {
        method = args.kind === "vba" ? "GetVBAMacros" : "GetMacros";
        params = null;
      } else if (call.name === "sheets_set_macros") {
        if (!args.content || typeof args.content !== "object" || Array.isArray(args.content)) {
          reject(new Error("sheets_set_macros.content 必须是宏配置对象"));
          return;
        }
        method = "SetMacros";
        params = [JSON.stringify(args.content)];
        needsSave = true;
      } else {
        reject(new Error("未知宏工具：" + call.name));
        return;
      }
      if (!window.Asc || !Asc.plugin || typeof Asc.plugin.executeMethod !== "function") {
        reject(new Error("当前 ONLYOFFICE 版本不支持宏插件方法"));
        return;
      }
      Asc.plugin.executeMethod(method, params, function (data) {
        try {
          let content = data;
          if (method === "GetMacros" && typeof data === "string") {
            try {
              content = JSON.parse(data);
            } catch (error) {
              content = { raw: data };
            }
          }
          resolve({
            ok: true,
            editorType: "cell",
            changed: needsSave ? 1 : 0,
            needsSave: needsSave,
            results: [{ name: call.name, kind: method === "GetVBAMacros" ? "vba" : "onlyoffice", content: content }],
          });
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  function execute(toolCalls) {
    if (toolCalls.some(function (call) {
      return call && (call.name === "sheets_inspect_macros" || call.name === "sheets_set_macros");
    })) {
      return executeMacroTool(toolCalls);
    }
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
              sheets_set_array_formula: true,
              sheets_manage_sheet: true,
              sheets_manage_range: true,
              sheets_set_rich_text: true,
              sheets_manage_names: true,
              sheets_recalculate: true,
              sheets_sort: true,
              sheets_filter: true,
              sheets_manage_table: true,
              sheets_manage_conditional_format: true,
              sheets_manage_validation: true,
              sheets_manage_pivot: true,
              sheets_manage_drawing: true,
              sheets_manage_hyperlink: true,
              sheets_manage_comments: true,
              sheets_manage_freeze_panes: true,
              sheets_manage_properties: true,
              sheets_manage_protected_ranges: true,
              sheets_manage_page_layout: true,
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

            function unitAlias(args, explicitName, legacyName) {
              var hasExplicit = hasOwn(args, explicitName);
              var hasLegacy = hasOwn(args, legacyName);
              if (hasExplicit && hasLegacy) {
                throw new Error(explicitName + " 与旧字段 " + legacyName + " 不能同时提供");
              }
              if (hasExplicit) return args[explicitName];
              return hasLegacy ? args[legacyName] : undefined;
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

            function requireMethod(value, method, feature) {
              if (!value || typeof value[method] !== "function") {
                throw new Error("当前 ONLYOFFICE 版本不支持" + (feature || method));
              }
              return value[method];
            }

            function arrayValue(value) {
              if (Array.isArray(value)) return value;
              if (!value) return [];
              if (typeof value.GetCount === "function" && typeof value.GetItem === "function") {
                var items = [];
                for (var itemIndex = 0; itemIndex < value.GetCount(); itemIndex += 1) {
                  items.push(value.GetItem(itemIndex));
                }
                return items;
              }
              return [];
            }

            function isoValue(value) {
              if (value instanceof Date) return value.toISOString();
              return value;
            }

            function normalizeCellInput(value) {
              if (Array.isArray(value)) return value.map(normalizeCellInput);
              if (!isObject(value) || !value.type || !hasOwn(value, "value")) return value;
              var taggedType = String(value.type);
              if (taggedType === "date" || taggedType === "datetime") {
                var taggedDate = new Date(value.value);
                if (isNaN(taggedDate.getTime())) throw new Error("无效日期值：" + value.value);
                return taggedDate;
              }
              if (taggedType === "time") {
                var timeMatch = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?$/.exec(String(value.value));
                if (!timeMatch) throw new Error("time 值必须是 HH:MM 或 HH:MM:SS");
                var timeSeconds = Number(timeMatch[1]) * 3600 + Number(timeMatch[2]) * 60 + Number(timeMatch[3] || 0);
                if (timeSeconds < 0 || timeSeconds >= 86400) throw new Error("time 值超出一天范围");
                return timeSeconds / 86400;
              }
              if (taggedType === "percent") {
                var percentText = String(value.value);
                var percentNumber = Number(percentText.replace(/%$/, ""));
                if (!isFinite(percentNumber)) throw new Error("无效百分比值：" + value.value);
                return /%$/.test(percentText) ? percentNumber / 100 : percentNumber;
              }
              if (taggedType === "boolean") {
                if (value.value === true || value.value === false) return value.value;
                if (String(value.value).toLowerCase() === "true") return true;
                if (String(value.value).toLowerCase() === "false") return false;
                throw new Error("无效布尔值：" + value.value);
              }
              throw new Error("不支持的单元格值类型：" + taggedType);
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

            function describeRange(range, includeValues) {
              if (!range) return null;
              return {
                address: safeCall(range, "GetAddress", true, true, "xlA1", false),
                values: includeValues === false ? undefined : safeCall(range, "GetValue"),
                values2: includeValues === false ? undefined : safeCall(range, "GetValue2"),
                text: includeValues === false ? undefined : safeCall(range, "GetText"),
                formula: safeCall(range, "GetFormula"),
                arrayFormula: safeCall(range, "GetFormulaArray"),
                numberFormat: safeCall(range, "GetNumberFormat"),
                rowHeight: safeCall(range, "GetRowHeight"),
                columnWidth: safeCall(range, "GetColumnWidth"),
                hidden: safeCall(range, "GetHidden"),
                wrap: safeCall(range, "GetWrapText"),
                orientation: safeCall(range, "GetOrientation"),
                rows: safeCall(range, "GetRowsCount"),
                columns: safeCall(range, "GetColumnsCount"),
              };
            }

            function getListObjects(sheet) {
              return arrayValue(
                requireMethod(sheet, "GetListObjects", "结构化表格枚举").call(sheet)
              );
            }

            function describeTable(table, index) {
              var listColumns = safeCall(table, "GetListColumns");
              var listRows = safeCall(table, "GetListRows");
              return {
                tableIndex: index,
                name: safeCall(table, "GetName"),
                displayName: safeCall(table, "GetDisplayName"),
                range: describeRange(safeCall(table, "GetRange"), false),
                sourceType: safeCall(table, "GetSourceType"),
                style: safeCall(table, "GetTableStyle"),
                showHeaders: safeCall(table, "GetShowHeaders"),
                showTotals: safeCall(table, "GetShowTotals"),
                rowStripes: safeCall(table, "GetShowTableStyleRowStripes"),
                columnStripes: safeCall(table, "GetShowTableStyleColumnStripes"),
                firstColumn: safeCall(table, "GetShowTableStyleFirstColumn"),
                lastColumn: safeCall(table, "GetShowTableStyleLastColumn"),
                showAutoFilter: safeCall(table, "GetShowAutoFilter"),
                showAutoFilterDropDown: safeCall(table, "GetShowAutoFilterDropDown"),
                summary: safeCall(table, "GetSummary"),
                alternativeText: safeCall(table, "GetAlternativeText"),
                columnCount: Array.isArray(listColumns) ? listColumns.length : null,
                rowCount: Array.isArray(listRows) ? listRows.length : null,
              };
            }

            function resolveTable(sheet, args) {
              var tables = getListObjects(sheet);
              if (hasOwn(args, "tableIndex")) {
                var requestedIndex = Number(args.tableIndex);
                if (requestedIndex >= 0 && requestedIndex < tables.length) {
                  return { table: tables[requestedIndex], index: requestedIndex };
                }
                throw new Error("tableIndex 超出范围：" + args.tableIndex);
              }
              var requestedName = args.tableName || args.name;
              if (requestedName) {
                for (var tableIndex = 0; tableIndex < tables.length; tableIndex += 1) {
                  var currentName = safeCall(tables[tableIndex], "GetName");
                  var displayName = safeCall(tables[tableIndex], "GetDisplayName");
                  if (String(currentName) === String(requestedName) || String(displayName) === String(requestedName)) {
                    return { table: tables[tableIndex], index: tableIndex };
                  }
                }
                throw new Error("找不到结构化表格：" + requestedName);
              }
              throw new Error("需要 tableIndex 或 tableName");
            }

            function hasTableChanges(args) {
              return args.newName !== undefined
                || args.style !== undefined
                || args.showTotals !== undefined
                || args.showHeaders !== undefined
                || args.rowStripes !== undefined
                || args.columnStripes !== undefined
                || args.firstColumn !== undefined
                || args.lastColumn !== undefined
                || args.showAutoFilter !== undefined
                || args.showAutoFilterDropDown !== undefined
                || args.summary !== undefined
                || args.alternativeText !== undefined;
            }

            function applyTableSettings(table, args, createMode) {
              var desiredName = createMode ? args.name : args.newName;
              if (desiredName !== undefined) {
                var renamedTable = requireMethod(table, "SetName", "结构化表格名称").call(table, String(desiredName));
                if (renamedTable === false) throw new Error("设置结构化表格名称失败");
              }
              if (args.style !== undefined) requireMethod(table, "SetTableStyle", "结构化表格样式").call(table, String(args.style));
              if (args.showTotals !== undefined) requireMethod(table, "SetShowTotals", "结构化表格汇总行").call(table, Boolean(args.showTotals));
              if (args.showHeaders !== undefined) requireMethod(table, "SetShowHeaders", "结构化表格标题行").call(table, Boolean(args.showHeaders));
              if (args.rowStripes !== undefined) {
                requireMethod(table, "SetShowTableStyleRowStripes", "结构化表格行条纹").call(table, Boolean(args.rowStripes));
              }
              if (args.columnStripes !== undefined) {
                requireMethod(table, "SetShowTableStyleColumnStripes", "结构化表格列条纹").call(table, Boolean(args.columnStripes));
              }
              if (args.firstColumn !== undefined) {
                requireMethod(table, "SetShowTableStyleFirstColumn", "结构化表格首列样式").call(table, Boolean(args.firstColumn));
              }
              if (args.lastColumn !== undefined) {
                requireMethod(table, "SetShowTableStyleLastColumn", "结构化表格末列样式").call(table, Boolean(args.lastColumn));
              }
              if (args.showAutoFilter !== undefined) {
                requireMethod(table, "SetShowAutoFilter", "结构化表格自动筛选").call(table, Boolean(args.showAutoFilter));
              }
              if (args.showAutoFilterDropDown !== undefined) {
                requireMethod(table, "SetShowAutoFilterDropDown", "结构化表格筛选按钮").call(
                  table,
                  Boolean(args.showAutoFilterDropDown)
                );
              }
              if (args.summary !== undefined) requireMethod(table, "SetSummary", "结构化表格摘要").call(table, String(args.summary));
              if (args.alternativeText !== undefined) {
                requireMethod(table, "SetAlternativeText", "结构化表格替代文本").call(table, String(args.alternativeText));
              }
            }

            function setDrawingPosition(drawing, args) {
              if (
                hasOwn(args, "fromColumn")
                || hasOwn(args, "fromRow")
                || hasOwn(args, "columnOffsetMm")
                || hasOwn(args, "rowOffsetMm")
              ) {
                requireMethod(drawing, "SetPosition", "对象定位").call(
                  drawing,
                  Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                  mmToEmu(args.columnOffsetMm || 0),
                  Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                  mmToEmu(args.rowOffsetMm || 0)
                );
              }
              if (hasOwn(args, "widthMm") || hasOwn(args, "heightMm")) {
                requireMethod(drawing, "SetSize", "对象尺寸").call(
                  drawing,
                  mmToEmu(Math.max(1, asFinite(args.widthMm, emuToMm(safeCall(drawing, "GetWidth")) || 60))),
                  mmToEmu(Math.max(1, asFinite(args.heightMm, emuToMm(safeCall(drawing, "GetHeight")) || 35)))
                );
              }
              if (hasOwn(args, "rotationDeg")) requireMethod(drawing, "SetRotation", "对象旋转").call(drawing, asFinite(args.rotationDeg, 0));
              if (hasOwn(args, "flipH")) requireMethod(drawing, "SetFlipH", "对象水平翻转").call(drawing, Boolean(args.flipH));
              if (hasOwn(args, "flipV")) requireMethod(drawing, "SetFlipV", "对象垂直翻转").call(drawing, Boolean(args.flipV));
              if (hasOwn(args, "name")) requireMethod(drawing, "SetName", "对象命名").call(drawing, String(args.name));
              if (hasOwn(args, "fill")) requireMethod(drawing, "Fill", "对象填充").call(drawing, createFill(args.fill));
              if (hasOwn(args, "line")) requireMethod(drawing, "SetOutLine", "对象线条").call(drawing, createStroke(args.line));
            }

            function describeDrawing(drawing, index, includeRaw) {
              return {
                drawingIndex: index,
                classType: safeCall(drawing, "GetClassType"),
                name: safeCall(drawing, "GetName"),
                widthMm: emuToMm(safeCall(drawing, "GetWidth")),
                heightMm: emuToMm(safeCall(drawing, "GetHeight")),
                rotationDeg: safeCall(drawing, "GetRotation"),
                flipH: safeCall(drawing, "GetFlipH"),
                flipV: safeCall(drawing, "GetFlipV"),
                raw: includeRaw ? serialized(drawing) : undefined,
              };
            }

            function resolveDrawing(args) {
              var sheet = getSheet(args.sheet);
              var drawings = typeof sheet.GetAllDrawings === "function" ? sheet.GetAllDrawings() : [];
              for (var drawingIndex = 0; drawingIndex < drawings.length; drawingIndex += 1) {
                if (hasOwn(args, "drawingIndex") && drawingIndex !== Number(args.drawingIndex)) continue;
                if (args.name && String(safeCall(drawings[drawingIndex], "GetName")) !== String(args.name)) continue;
                if (hasOwn(args, "drawingIndex") || args.name) {
                  return { sheet: sheet, drawing: drawings[drawingIndex], index: drawingIndex };
                }
              }
              throw new Error("找不到目标对象；请提供有效的 drawingIndex 或 name");
            }

            function getComment(args) {
              var commentSheet = getSheet(args.sheet);
              if (args.range) {
                var rangeComment = getRange(commentSheet, args.range);
                var directComment = safeCall(rangeComment, "GetComment");
                if (directComment) return { sheet: commentSheet, comment: directComment };
              }
              var comments = typeof commentSheet.GetComments === "function" ? commentSheet.GetComments() : [];
              for (var commentIndex = 0; commentIndex < comments.length; commentIndex += 1) {
                if (args.commentId && String(safeCall(comments[commentIndex], "GetId")) === String(args.commentId)) {
                  return { sheet: commentSheet, comment: comments[commentIndex] };
                }
              }
              throw new Error("找不到目标批注；请提供有效的 range 或 commentId");
            }

            function describeComment(comment) {
              var replies = [];
              var repliesCount = asFinite(safeCall(comment, "GetRepliesCount"), 0);
              for (var replyIndex = 0; replyIndex < repliesCount; replyIndex += 1) {
                var reply = safeCall(comment, "GetReply", replyIndex);
                replies.push({
                  text: safeCall(reply, "GetText"),
                  author: safeCall(reply, "GetAuthorName"),
                  userId: safeCall(reply, "GetUserId"),
                  time: isoValue(safeCall(reply, "GetTime")),
                });
              }
              return {
                id: safeCall(comment, "GetId"),
                text: safeCall(comment, "GetText"),
                quoteText: safeCall(comment, "GetQuoteText"),
                author: safeCall(comment, "GetAuthorName"),
                userId: safeCall(comment, "GetUserId"),
                solved: safeCall(comment, "IsSolved"),
                time: isoValue(safeCall(comment, "GetTime")),
                replies: replies,
              };
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

                case "sheets_inspect_range": {
                  if (!args.range) throw new Error("sheets_inspect_range.range 不能为空");
                  var inspectRangeSheet = getSheet(args.sheet);
                  var inspectedRange = getRange(inspectRangeSheet, args.range);
                  if (!inspectedRange) throw new Error("无效单元格区域：" + args.range);
                  results.push({
                    name: call.name,
                    sheet: inspectRangeSheet.GetName(),
                    range: describeRange(inspectedRange, args.includeValues !== false),
                  });
                  break;
                }

                case "sheets_set_values": {
                  if (!args.range) throw new Error("sheets_set_values.range 不能为空");
                  if (args.values === undefined) throw new Error("sheets_set_values.values 不能为空");
                  var valuesSheet = getSheet(args.sheet);
                  var range = getRange(valuesSheet, args.range);
                  if (!range) throw new Error("无效单元格区域：" + args.range);
                  var accepted = range.SetValue(normalizeCellInput(args.values));
                  changed += 1;
                  results.push({ name: call.name, sheet: valuesSheet.GetName(), range: String(args.range), accepted: Boolean(accepted) });
                  break;
                }

                case "sheets_set_array_formula": {
                  if (!args.range || args.formula === undefined) throw new Error("sheets_set_array_formula 需要 range 和 formula");
                  var arrayFormulaSheet = getSheet(args.sheet);
                  var arrayFormulaRange = getRange(arrayFormulaSheet, args.range);
                  if (!arrayFormulaRange) throw new Error("无效单元格区域：" + args.range);
                  var arrayFormula = String(args.formula);
                  if (arrayFormula.charAt(0) !== "=") arrayFormula = "=" + arrayFormula;
                  requireMethod(arrayFormulaRange, "SetFormulaArray", "数组公式").call(arrayFormulaRange, arrayFormula);
                  changed += 1;
                  results.push({
                    name: call.name,
                    sheet: arrayFormulaSheet.GetName(),
                    range: String(args.range),
                    formula: arrayFormula,
                  });
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
                  if (args.strikeout !== undefined) requireMethod(formatRange, "SetStrikeout", "删除线").call(formatRange, Boolean(args.strikeout));
                  if (args.fontColor) formatRange.SetFontColor(color(args.fontColor));
                  if (args.fillColor) formatRange.SetFillColor(color(args.fillColor));
                  if (args.horizontalAlign) formatRange.SetAlignHorizontal(String(args.horizontalAlign));
                  if (args.verticalAlign) formatRange.SetAlignVertical(String(args.verticalAlign));
                  if (args.numberFormat) formatRange.SetNumberFormat(String(args.numberFormat));
                  if (args.wrap !== undefined) formatRange.SetWrap(Boolean(args.wrap));
                  if (args.orientation !== undefined) requireMethod(formatRange, "SetOrientation", "文本旋转").call(formatRange, args.orientation);
                  var columnWidthChars = unitAlias(args, "columnWidthChars", "columnWidth");
                  var rowHeightPt = unitAlias(args, "rowHeightPt", "rowHeight");
                  if (columnWidthChars !== undefined) formatRange.SetColumnWidth(Number(columnWidthChars));
                  if (rowHeightPt !== undefined) formatRange.SetRowHeight(Number(rowHeightPt));
                  if (Array.isArray(args.borders)) {
                    for (var borderIndex = 0; borderIndex < args.borders.length; borderIndex += 1) {
                      var border = args.borders[borderIndex] || {};
                      if (!border.side || !border.style || !border.color) throw new Error("每个边框需要 side、style 和 color");
                      requireMethod(formatRange, "SetBorders", "单元格边框").call(
                        formatRange,
                        String(border.side),
                        String(border.style),
                        color(border.color)
                      );
                    }
                  }
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

                case "sheets_manage_sheet": {
                  if (!args.action) throw new Error("sheets_manage_sheet.action 不能为空");
                  var managedSheet = getSheet(args.sheet);
                  var sheetAction = String(args.action);
                  if (sheetAction === "setVisibility") {
                    if (!hasOwn(args, "visible")) throw new Error("setVisibility 需要 visible");
                    if (!args.visible) {
                      var visibleCount = 0;
                      var visibilitySheets = Api.GetSheets();
                      for (var visibleIndex = 0; visibleIndex < visibilitySheets.length; visibleIndex += 1) {
                        if (safeCall(visibilitySheets[visibleIndex], "GetVisible") !== false) visibleCount += 1;
                      }
                      if (visibleCount <= 1 && safeCall(managedSheet, "GetVisible") !== false) {
                        throw new Error("工作簿至少需要保留一个可见工作表");
                      }
                    }
                    requireMethod(managedSheet, "SetVisible", "工作表显示/隐藏").call(managedSheet, Boolean(args.visible));
                  } else if (sheetAction === "setActive") {
                    requireMethod(managedSheet, "SetActive", "激活工作表").call(managedSheet);
                  } else if (sheetAction === "moveBefore") {
                    if (!args.beforeSheet) throw new Error("moveBefore 需要 beforeSheet");
                    requireMethod(managedSheet, "Move", "调整工作表顺序").call(managedSheet, getSheet(args.beforeSheet));
                  } else if (sheetAction === "copy") {
                    if (!args.newName) throw new Error("copy 需要 newName");
                    var copiedSheet = Api.AddSheet(String(args.newName));
                    if (!copiedSheet) throw new Error("复制工作表失败，目标名称可能已存在");
                    var copiedUsed = managedSheet.GetUsedRange();
                    if (copiedUsed) {
                      var copiedAddress = safeCall(copiedUsed, "GetAddress", false, false, "xlA1", false) || "A1";
                      requireMethod(copiedUsed, "Copy", "复制工作表单元格").call(copiedUsed, copiedSheet.GetRange(copiedAddress));
                    }
                    var copiedDrawings = typeof managedSheet.GetAllDrawings === "function" ? managedSheet.GetAllDrawings() : [];
                    for (var copiedDrawingIndex = 0; copiedDrawingIndex < copiedDrawings.length; copiedDrawingIndex += 1) {
                      var detachedDrawing = safeCall(copiedDrawings[copiedDrawingIndex], "Copy");
                      if (detachedDrawing && typeof copiedSheet.AddDrawing === "function") copiedSheet.AddDrawing(detachedDrawing, 0, 0, 0, 0);
                    }
                    managedSheet = copiedSheet;
                  } else {
                    throw new Error("不支持的工作表操作：" + sheetAction);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: sheetAction,
                    sheet: managedSheet.GetName(),
                    visible: safeCall(managedSheet, "GetVisible"),
                    index: safeCall(managedSheet, "GetIndex"),
                  });
                  break;
                }

                case "sheets_manage_range": {
                  if (!args.action || !args.range) throw new Error("sheets_manage_range 需要 action 和 range");
                  var managedRangeSheet = getSheet(args.sheet);
                  var managedRange = getRange(managedRangeSheet, args.range);
                  if (!managedRange) throw new Error("无效单元格区域：" + args.range);
                  var rangeAction = String(args.action);
                  if (rangeAction === "merge") requireMethod(managedRange, "Merge", "合并单元格").call(managedRange, Boolean(args.across));
                  else if (rangeAction === "unmerge") requireMethod(managedRange, "UnMerge", "取消合并单元格").call(managedRange);
                  else if (rangeAction === "insert") requireMethod(managedRange, "Insert", "插入单元格/行列").call(managedRange, String(args.shift || "down"));
                  else if (rangeAction === "delete") requireMethod(managedRange, "Delete", "删除单元格/行列").call(managedRange, String(args.shift || "up"));
                  else if (rangeAction === "copy" || rangeAction === "cut") {
                    if (!args.destination) throw new Error(rangeAction + " 需要 destination");
                    var destinationSheet = getSheet(args.destinationSheet || args.sheet);
                    var destinationRange = getRange(destinationSheet, args.destination);
                    requireMethod(managedRange, rangeAction === "copy" ? "Copy" : "Cut", rangeAction === "copy" ? "复制区域" : "移动区域")
                      .call(managedRange, destinationRange);
                  } else if (rangeAction === "clear") requireMethod(managedRange, "Clear", "清除区域").call(managedRange);
                  else if (rangeAction === "clearContents") requireMethod(managedRange, "ClearContents", "清除内容").call(managedRange);
                  else if (rangeAction === "clearFormats") requireMethod(managedRange, "ClearFormats", "清除格式").call(managedRange);
                  else if (rangeAction === "clearHyperlinks") requireMethod(managedRange, "ClearHyperlinks", "清除超链接").call(managedRange);
                  else if (rangeAction === "autoFit") requireMethod(managedRange, "AutoFit", "自动调整行高列宽").call(
                    managedRange,
                    hasOwn(args, "rows") ? Boolean(args.rows) : true,
                    hasOwn(args, "columns") ? Boolean(args.columns) : true
                  );
                  else if (rangeAction === "setHidden") {
                    if (!hasOwn(args, "hidden")) throw new Error("setHidden 需要 hidden");
                    requireMethod(managedRange, "SetHidden", "隐藏行列").call(managedRange, Boolean(args.hidden));
                  } else if (rangeAction === "fillDown") requireMethod(managedRange, "FillDown", "向下填充").call(managedRange);
                  else if (rangeAction === "fillUp") requireMethod(managedRange, "FillUp", "向上填充").call(managedRange);
                  else if (rangeAction === "fillLeft") requireMethod(managedRange, "FillLeft", "向左填充").call(managedRange);
                  else if (rangeAction === "fillRight") requireMethod(managedRange, "FillRight", "向右填充").call(managedRange);
                  else if (rangeAction === "select") requireMethod(managedRange, "Select", "选择区域").call(managedRange);
                  else throw new Error("不支持的区域操作：" + rangeAction);
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: rangeAction,
                    sheet: managedRangeSheet.GetName(),
                    range: describeRange(managedRange, false),
                  });
                  break;
                }

                case "sheets_set_rich_text": {
                  if (!args.range || !Array.isArray(args.runs) || !args.runs.length) {
                    throw new Error("sheets_set_rich_text 需要 range 和非空 runs");
                  }
                  var richTextSheet = getSheet(args.sheet);
                  var richTextRange = getRange(richTextSheet, args.range);
                  if (!richTextRange) throw new Error("无效单元格区域：" + args.range);
                  if (args.text !== undefined) richTextRange.SetValue(String(args.text));
                  for (var richRunIndex = 0; richRunIndex < args.runs.length; richRunIndex += 1) {
                    var richRun = args.runs[richRunIndex] || {};
                    var characters = requireMethod(richTextRange, "GetCharacters", "单元格内局部富文本")
                      .call(richTextRange, Math.max(0, Math.round(asFinite(richRun.start, 0))), Math.max(0, Math.round(asFinite(richRun.length, 0))));
                    if (richRun.text !== undefined) requireMethod(characters, "SetText", "富文本内容").call(characters, String(richRun.text));
                    var richFont = requireMethod(characters, "GetFont", "富文本字体").call(characters);
                    if (hasOwn(richRun, "bold")) requireMethod(richFont, "SetBold", "富文本粗体").call(richFont, Boolean(richRun.bold));
                    if (hasOwn(richRun, "italic")) requireMethod(richFont, "SetItalic", "富文本斜体").call(richFont, Boolean(richRun.italic));
                    if (hasOwn(richRun, "underline")) requireMethod(richFont, "SetUnderline", "富文本下划线").call(richFont, richRun.underline === true ? "xlUnderlineStyleSingle" : richRun.underline);
                    if (hasOwn(richRun, "strikeout")) requireMethod(richFont, "SetStrikethrough", "富文本删除线").call(richFont, Boolean(richRun.strikeout));
                    if (richRun.fontName) requireMethod(richFont, "SetName", "富文本字体").call(richFont, String(richRun.fontName));
                    if (richRun.fontSize !== undefined) requireMethod(richFont, "SetSize", "富文本字号").call(richFont, Number(richRun.fontSize));
                    if (richRun.fontColor) requireMethod(richFont, "SetColor", "富文本颜色").call(richFont, color(richRun.fontColor));
                  }
                  changed += 1;
                  results.push({ name: call.name, sheet: richTextSheet.GetName(), range: String(args.range), runs: args.runs.length });
                  break;
                }

                case "sheets_inspect_names": {
                  var names = typeof Api.GetDefNames === "function" ? Api.GetDefNames() : [];
                  results.push({
                    name: call.name,
                    names: arrayValue(names).map(function (definedName) {
                      return {
                        name: safeCall(definedName, "GetName"),
                        refersTo: safeCall(definedName, "GetRefersTo"),
                        range: describeRange(safeCall(definedName, "GetRefersToRange"), false),
                      };
                    }),
                  });
                  break;
                }

                case "sheets_manage_names": {
                  if (!args.action || !args.name) throw new Error("sheets_manage_names 需要 action 和 name");
                  var nameAction = String(args.action);
                  var managedName;
                  if (nameAction === "add") {
                    if (!args.refersTo) throw new Error("add 需要 refersTo");
                    managedName = requireMethod(Api, "AddDefName", "命名区域").call(Api, String(args.name), String(args.refersTo));
                    if (!managedName) managedName = safeCall(Api, "GetDefName", String(args.name));
                  } else {
                    managedName = requireMethod(Api, "GetDefName", "命名区域").call(Api, String(args.name));
                    if (!managedName) throw new Error("找不到命名区域：" + args.name);
                    if (nameAction === "update") {
                      if (args.newName) requireMethod(managedName, "SetName", "重命名命名区域").call(managedName, String(args.newName));
                      if (args.refersTo) requireMethod(managedName, "SetRefersTo", "更新命名区域").call(managedName, String(args.refersTo));
                    } else if (nameAction === "delete") {
                      requireMethod(managedName, "Delete", "删除命名区域").call(managedName);
                    } else {
                      throw new Error("不支持的命名区域操作：" + nameAction);
                    }
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: nameAction,
                    definedName: nameAction === "delete" ? String(args.name) : {
                      name: safeCall(managedName, "GetName") || args.newName || args.name,
                      refersTo: safeCall(managedName, "GetRefersTo") || args.refersTo,
                    },
                  });
                  break;
                }

                case "sheets_recalculate": {
                  var recalculateMode = String(args.mode || "formulas");
                  if (recalculateMode === "formulas" || recalculateMode === "all") {
                    requireMethod(Api, "RecalculateAllFormulas", "公式重新计算").call(Api);
                  }
                  if (recalculateMode === "pivots" || recalculateMode === "all") {
                    requireMethod(Api, "RefreshAllPivots", "刷新数据透视表").call(Api);
                  }
                  if (recalculateMode !== "formulas" && recalculateMode !== "pivots" && recalculateMode !== "all") {
                    throw new Error("mode 必须是 formulas、pivots 或 all");
                  }
                  changed += 1;
                  results.push({ name: call.name, mode: recalculateMode });
                  break;
                }

                case "sheets_sort": {
                  if (!args.range || !Array.isArray(args.keys) || !args.keys.length || args.keys.length > 3) {
                    throw new Error("sheets_sort 需要 range 和 1 到 3 个 keys");
                  }
                  var sortSheet = getSheet(args.sheet);
                  var sortRange = getRange(sortSheet, args.range);
                  var sortArguments = [];
                  for (var sortKeyIndex = 0; sortKeyIndex < 3; sortKeyIndex += 1) {
                    var sortKey = args.keys[sortKeyIndex];
                    sortArguments.push(sortKey ? String(sortKey.range) : undefined);
                    sortArguments.push(sortKey ? String(sortKey.order || "xlAscending") : undefined);
                  }
                  sortArguments.push(String(args.header || "xlGuess"));
                  sortArguments.push(String(args.orientation || "xlSortColumns"));
                  requireMethod(sortRange, "SetSort", "排序").apply(sortRange, sortArguments);
                  changed += 1;
                  results.push({ name: call.name, sheet: sortSheet.GetName(), range: String(args.range), keys: args.keys });
                  break;
                }

                case "sheets_filter": {
                  if (!args.action) throw new Error("sheets_filter.action 不能为空");
                  var filterSheet = getSheet(args.sheet);
                  var filterAction = String(args.action);
                  if (filterAction === "set") {
                    if (!args.range) throw new Error("set 需要 range");
                    var filterRange = getRange(filterSheet, args.range);
                    requireMethod(filterRange, "SetAutoFilter", "自动筛选").call(
                      filterRange,
                      hasOwn(args, "field") ? Number(args.field) : null,
                      args.criteria1,
                      args.operator ? String(args.operator) : undefined,
                      args.criteria2,
                      hasOwn(args, "visibleDropDown") ? Boolean(args.visibleDropDown) : undefined
                    );
                  } else if (filterAction === "showAll") {
                    var autoFilter = typeof filterSheet.GetAutoFilter === "function" ? filterSheet.GetAutoFilter() : filterSheet.AutoFilter;
                    requireMethod(autoFilter, "ShowAllData", "显示全部筛选数据").call(autoFilter);
                  } else if (filterAction === "reapply") {
                    var reappliedFilter = typeof filterSheet.GetAutoFilter === "function" ? filterSheet.GetAutoFilter() : filterSheet.AutoFilter;
                    requireMethod(reappliedFilter, "ApplyFilter", "重新应用筛选").call(reappliedFilter);
                  } else {
                    throw new Error("不支持的筛选操作：" + filterAction);
                  }
                  changed += 1;
                  var inspectedFilter = typeof filterSheet.GetAutoFilter === "function" ? filterSheet.GetAutoFilter() : filterSheet.AutoFilter;
                  results.push({
                    name: call.name,
                    action: filterAction,
                    sheet: filterSheet.GetName(),
                    range: describeRange(safeCall(inspectedFilter, "GetRange"), false),
                    filterMode: safeCall(inspectedFilter, "GetFilterMode"),
                  });
                  break;
                }

                case "sheets_inspect_tables": {
                  var tableInspectionSheets = args.sheet ? [getSheet(args.sheet)] : Api.GetSheets();
                  var remainingTables = Math.max(1, Math.min(1000, Number(args.maxTables || 200)));
                  var inspectedTableSheets = [];
                  for (var inspectedSheetIndex = 0; inspectedSheetIndex < tableInspectionSheets.length; inspectedSheetIndex += 1) {
                    var inspectedTableSheet = tableInspectionSheets[inspectedSheetIndex];
                    var sheetTables = getListObjects(inspectedTableSheet);
                    var describedTables = [];
                    for (
                      var inspectedTableIndex = 0;
                      inspectedTableIndex < sheetTables.length && remainingTables > 0;
                      inspectedTableIndex += 1
                    ) {
                      var inspectedTable = describeTable(sheetTables[inspectedTableIndex], inspectedTableIndex);
                      if (hasOwn(args, "tableIndex") && inspectedTableIndex !== Number(args.tableIndex)) continue;
                      if (
                        args.name
                        && String(inspectedTable.name) !== String(args.name)
                        && String(inspectedTable.displayName) !== String(args.name)
                      ) continue;
                      describedTables.push(inspectedTable);
                      remainingTables -= 1;
                    }
                    inspectedTableSheets.push({
                      sheet: inspectedTableSheet.GetName(),
                      tableCount: sheetTables.length,
                      tables: describedTables,
                    });
                  }
                  results.push({ name: call.name, sheets: inspectedTableSheets });
                  break;
                }

                case "sheets_manage_table": {
                  if (!args.action) throw new Error("sheets_manage_table.action 不能为空");
                  var tableSheet = getSheet(args.sheet);
                  var tableAction = String(args.action);
                  var table;
                  var managedTableIndex = null;
                  if (tableAction === "create") {
                    if (!args.range) throw new Error("create 需要 range");
                    if (typeof tableSheet.AddListObject === "function") {
                      table = tableSheet.AddListObject(String(args.sourceType || "xlSrcRange"), String(args.range));
                      if (!table) throw new Error("创建格式化表格失败");
                    } else {
                      if (args.name || hasTableChanges(args)) {
                        throw new Error("当前 ONLYOFFICE 版本只能创建基础格式化表格，不支持结构化表格属性");
                      }
                      table = requireMethod(tableSheet, "FormatAsTable", "Excel 格式化表格").call(tableSheet, String(args.range));
                      if (table === false) throw new Error("格式化表格失败");
                    }
                    if (table && typeof table === "object") applyTableSettings(table, args, true);
                  } else if (tableAction === "format") {
                    if (!args.range) throw new Error("format 需要 range");
                    table = requireMethod(tableSheet, "FormatAsTable", "Excel 格式化表格").call(tableSheet, String(args.range));
                    if (table === false) throw new Error("格式化表格失败");
                  } else if (
                    tableAction === "update"
                    || tableAction === "resize"
                    || tableAction === "delete"
                    || tableAction === "unlist"
                  ) {
                    var resolvedTable = resolveTable(tableSheet, args);
                    table = resolvedTable.table;
                    managedTableIndex = resolvedTable.index;
                    if (tableAction === "update") {
                      if (!hasTableChanges(args)) throw new Error("update 至少需要一个表格属性");
                      applyTableSettings(table, args, false);
                    } else if (tableAction === "resize") {
                      if (!args.range) throw new Error("resize 需要 range");
                      requireMethod(table, "Resize", "调整结构化表格区域").call(table, String(args.range));
                    } else if (tableAction === "delete") {
                      requireMethod(table, "Delete", "删除结构化表格").call(table);
                    } else {
                      requireMethod(table, "Unlist", "结构化表格转普通区域").call(table);
                    }
                  } else {
                    throw new Error("不支持的结构化表格操作：" + tableAction);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: tableAction,
                    sheet: tableSheet.GetName(),
                    table: tableAction === "delete" || tableAction === "unlist"
                      ? null
                      : (
                        table && typeof table === "object"
                          ? describeTable(table, managedTableIndex)
                          : {
                            name: args.name || null,
                            range: describeRange(getRange(tableSheet, args.range), false),
                            style: args.style || null,
                          }
                      ),
                  });
                  break;
                }

                case "sheets_manage_conditional_format": {
                  if (!args.action || !args.range) throw new Error("sheets_manage_conditional_format 需要 action 和 range");
                  var conditionalSheet = getSheet(args.sheet);
                  var conditionalRange = getRange(conditionalSheet, args.range);
                  var conditions = requireMethod(conditionalRange, "GetFormatConditions", "条件格式").call(conditionalRange);
                  var conditionalAction = String(args.action);
                  var condition = null;
                  if (conditionalAction === "deleteAll") {
                    requireMethod(conditions, "Delete", "删除条件格式").call(conditions);
                  } else if (conditionalAction === "add") {
                    var conditionType = String(args.type || "cellValue");
                    if (conditionType === "cellValue" || conditionType === "expression") {
                      condition = requireMethod(conditions, "Add", "条件格式规则").call(
                        conditions,
                        conditionType === "expression" ? "xlExpression" : "xlCellValue",
                        String(args.operator || "xlGreater"),
                        args.formula1,
                        args.formula2
                      );
                    } else if (conditionType === "uniqueValues" || conditionType === "duplicateValues") {
                      condition = requireMethod(conditions, "AddUniqueValues", "重复值条件格式").call(conditions);
                      if (condition && typeof condition.SetDupeUnique === "function") {
                        condition.SetDupeUnique(conditionType === "duplicateValues" ? "xlDuplicate" : "xlUnique");
                      }
                    } else if (conditionType === "colorScale") {
                      condition = requireMethod(conditions, "AddColorScale", "色阶条件格式").call(conditions, clamp(Math.round(asFinite(args.scale, 3)), 2, 3));
                    } else if (conditionType === "dataBar") {
                      condition = requireMethod(conditions, "AddDatabar", "数据条条件格式").call(conditions);
                    } else if (conditionType === "iconSet") {
                      condition = requireMethod(conditions, "AddIconSetCondition", "图标集条件格式").call(conditions);
                    } else if (conditionType === "top10") {
                      condition = requireMethod(conditions, "AddTop10", "前十项条件格式").call(conditions);
                    } else if (conditionType === "aboveAverage") {
                      condition = requireMethod(conditions, "AddAboveAverage", "高于平均值条件格式").call(conditions);
                    } else {
                      throw new Error("不支持的条件格式类型：" + conditionType);
                    }
                    if (condition) {
                      if (args.fillColor && typeof condition.SetFillColor === "function") condition.SetFillColor(color(args.fillColor));
                      var conditionFont = safeCall(condition, "GetFont");
                      if (conditionFont) {
                        if (args.fontColor) requireMethod(conditionFont, "SetColor", "条件格式字体颜色").call(conditionFont, color(args.fontColor));
                        if (hasOwn(args, "bold")) requireMethod(conditionFont, "SetBold", "条件格式粗体").call(conditionFont, Boolean(args.bold));
                      }
                    }
                  } else {
                    throw new Error("条件格式 action 必须是 add 或 deleteAll");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: conditionalAction,
                    sheet: conditionalSheet.GetName(),
                    range: String(args.range),
                    count: safeCall(conditions, "GetCount"),
                  });
                  break;
                }

                case "sheets_manage_validation": {
                  if (!args.action || !args.range) throw new Error("sheets_manage_validation 需要 action 和 range");
                  var validationSheet = getSheet(args.sheet);
                  var validationRange = getRange(validationSheet, args.range);
                  var validation = requireMethod(validationRange, "GetValidation", "数据验证").call(validationRange);
                  var validationAction = String(args.action);
                  if (validationAction === "delete") {
                    requireMethod(validation, "Delete", "删除数据验证").call(validation);
                  } else if (validationAction === "add" || validationAction === "modify") {
                    var validationMethod = validationAction === "add" ? "Add" : "Modify";
                    requireMethod(validation, validationMethod, "数据验证规则").call(
                      validation,
                      String(args.type || "xlValidateList"),
                      String(args.alertStyle || "xlValidAlertStop"),
                      String(args.operator || "xlBetween"),
                      args.formula1,
                      args.formula2
                    );
                    if (hasOwn(args, "ignoreBlank") && typeof validation.SetIgnoreBlank === "function") validation.SetIgnoreBlank(Boolean(args.ignoreBlank));
                    if (hasOwn(args, "showInput")) requireMethod(validation, "SetShowInput", "数据验证输入提示").call(validation, Boolean(args.showInput));
                    if (hasOwn(args, "showError")) requireMethod(validation, "SetShowError", "数据验证错误提示").call(validation, Boolean(args.showError));
                    if (args.inputTitle !== undefined) requireMethod(validation, "SetInputTitle", "数据验证输入标题").call(validation, String(args.inputTitle));
                    if (args.inputMessage !== undefined) requireMethod(validation, "SetInputMessage", "数据验证输入消息").call(validation, String(args.inputMessage));
                    if (args.errorTitle !== undefined) requireMethod(validation, "SetErrorTitle", "数据验证错误标题").call(validation, String(args.errorTitle));
                    if (args.errorMessage !== undefined) requireMethod(validation, "SetErrorMessage", "数据验证错误消息").call(validation, String(args.errorMessage));
                  } else {
                    throw new Error("数据验证 action 必须是 add、modify 或 delete");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: validationAction,
                    sheet: validationSheet.GetName(),
                    range: String(args.range),
                  });
                  break;
                }

                case "sheets_inspect_pivots": {
                  var inspectedPivots = typeof Api.GetAllPivotTables === "function" ? Api.GetAllPivotTables() : [];
                  results.push({
                    name: call.name,
                    pivots: arrayValue(inspectedPivots).map(function (pivot) {
                      return {
                        name: safeCall(pivot, "GetName"),
                        title: safeCall(pivot, "GetTitle"),
                        description: safeCall(pivot, "GetDescription"),
                        source: safeCall(pivot, "GetSource"),
                        style: safeCall(pivot, "GetStyleName"),
                        tableRange: describeRange(safeCall(pivot, "GetTableRange2"), false),
                        rowFields: arrayValue(safeCall(pivot, "GetRowFields")).map(function (field) { return safeCall(field, "GetName"); }),
                        columnFields: arrayValue(safeCall(pivot, "GetColumnFields")).map(function (field) { return safeCall(field, "GetName"); }),
                        dataFields: arrayValue(safeCall(pivot, "GetDataFields")).map(function (field) { return safeCall(field, "GetName"); }),
                      };
                    }),
                  });
                  break;
                }

                case "sheets_manage_pivot": {
                  if (!args.action) throw new Error("sheets_manage_pivot.action 不能为空");
                  var pivotAction = String(args.action);
                  var managedPivot;
                  if (pivotAction === "createNewSheet" || pivotAction === "createExisting") {
                    if (!args.sourceRange) throw new Error(pivotAction + " 需要 sourceRange");
                    var pivotSourceSheet = getSheet(args.sourceSheet || args.sheet);
                    var pivotSource = typeof Api.GetRange === "function"
                      ? Api.GetRange(qualifyRange(pivotSourceSheet, args.sourceRange))
                      : getRange(pivotSourceSheet, args.sourceRange);
                    if (pivotAction === "createNewSheet") {
                      managedPivot = requireMethod(Api, "InsertPivotNewWorksheet", "新建工作表数据透视表").call(Api, pivotSource);
                    } else {
                      if (!args.destinationRange) throw new Error("createExisting 需要 destinationRange");
                      var pivotDestinationSheet = getSheet(args.destinationSheet || args.sheet);
                      managedPivot = requireMethod(Api, "InsertPivotExistingWorksheet", "现有工作表数据透视表").call(
                        Api,
                        pivotSource,
                        getRange(pivotDestinationSheet, args.destinationRange)
                      );
                    }
                  } else if (pivotAction === "refreshAll") {
                    requireMethod(Api, "RefreshAllPivots", "刷新全部数据透视表").call(Api);
                  } else {
                    if (!args.name) throw new Error(pivotAction + " 需要 name");
                    managedPivot = requireMethod(Api, "GetPivotByName", "按名称获取数据透视表").call(Api, String(args.name));
                    if (!managedPivot) throw new Error("找不到数据透视表：" + args.name);
                    if (pivotAction === "refresh") requireMethod(managedPivot, "RefreshTable", "刷新数据透视表").call(managedPivot);
                    else if (pivotAction === "clear") requireMethod(managedPivot, "ClearTable", "清空数据透视表").call(managedPivot);
                    else if (pivotAction !== "update") throw new Error("不支持的数据透视表操作：" + pivotAction);
                  }
                  if (managedPivot) {
                    if (args.name && (pivotAction === "createNewSheet" || pivotAction === "createExisting") && typeof managedPivot.SetName === "function") {
                      managedPivot.SetName(String(args.name));
                    }
                    if (args.fields && typeof managedPivot.AddFields === "function") managedPivot.AddFields(args.fields);
                    if (Array.isArray(args.dataFields) && typeof managedPivot.AddDataField === "function") {
                      for (var dataFieldIndex = 0; dataFieldIndex < args.dataFields.length; dataFieldIndex += 1) {
                        var dataField = args.dataFields[dataFieldIndex];
                        var addedDataField = managedPivot.AddDataField(typeof dataField === "string" ? dataField : dataField.name);
                        if (addedDataField && isObject(dataField) && dataField.function && typeof addedDataField.SetFunction === "function") {
                          addedDataField.SetFunction(String(dataField.function));
                        }
                      }
                    }
                    if (args.style && typeof managedPivot.SetStyleName === "function") managedPivot.SetStyleName(String(args.style));
                    if (args.title !== undefined && typeof managedPivot.SetTitle === "function") managedPivot.SetTitle(String(args.title));
                    if (args.description !== undefined && typeof managedPivot.SetDescription === "function") managedPivot.SetDescription(String(args.description));
                    if (hasOwn(args, "rowGrand") && typeof managedPivot.SetRowGrand === "function") managedPivot.SetRowGrand(Boolean(args.rowGrand));
                    if (hasOwn(args, "columnGrand") && typeof managedPivot.SetColumnGrand === "function") managedPivot.SetColumnGrand(Boolean(args.columnGrand));
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: pivotAction,
                    pivot: managedPivot ? {
                      name: safeCall(managedPivot, "GetName"),
                      source: safeCall(managedPivot, "GetSource"),
                      style: safeCall(managedPivot, "GetStyleName"),
                    } : null,
                  });
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

                case "sheets_inspect_drawings": {
                  var drawingSheets = args.sheet ? [getSheet(args.sheet)] : Api.GetSheets();
                  var maxDrawings = clamp(Math.round(asFinite(args.maxDrawings, 200)), 1, 1000);
                  var drawingSheetResults = [];
                  var inspectedDrawingCount = 0;
                  for (var drawingSheetIndex = 0; drawingSheetIndex < drawingSheets.length && inspectedDrawingCount < maxDrawings; drawingSheetIndex += 1) {
                    var inspectedDrawingSheet = drawingSheets[drawingSheetIndex];
                    var sheetDrawings = typeof inspectedDrawingSheet.GetAllDrawings === "function" ? inspectedDrawingSheet.GetAllDrawings() : [];
                    var drawingItems = [];
                    for (var inspectedDrawingIndex = 0; inspectedDrawingIndex < sheetDrawings.length && inspectedDrawingCount < maxDrawings; inspectedDrawingIndex += 1) {
                      drawingItems.push(describeDrawing(sheetDrawings[inspectedDrawingIndex], inspectedDrawingIndex, Boolean(args.includeRaw)));
                      inspectedDrawingCount += 1;
                    }
                    drawingSheetResults.push({ sheet: inspectedDrawingSheet.GetName(), drawings: drawingItems });
                  }
                  results.push({
                    name: call.name,
                    sheets: drawingSheetResults,
                    drawingCount: inspectedDrawingCount,
                    truncated: inspectedDrawingCount >= maxDrawings,
                  });
                  break;
                }

                case "sheets_manage_drawing": {
                  if (!args.action) throw new Error("sheets_manage_drawing.action 不能为空");
                  var drawingAction = String(args.action);
                  var drawingSheet = getSheet(args.sheet);
                  var drawing;
                  var managedDrawingIndex;
                  if (drawingAction === "addImage") {
                    if (!args._image || !args._image.url) throw new Error("addImage 需要经过宿主页安全导入的 source");
                    drawing = requireMethod(drawingSheet, "AddImage", "插入图片").call(
                      drawingSheet,
                      String(args._image.url),
                      mmToEmu(Math.max(1, asFinite(args.widthMm, 60))),
                      mmToEmu(Math.max(1, asFinite(args.heightMm, 35))),
                      Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                      mmToEmu(args.columnOffsetMm || 0),
                      Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                      mmToEmu(args.rowOffsetMm || 0)
                    );
                  } else if (drawingAction === "addShape" || drawingAction === "addTextBox") {
                    drawing = requireMethod(drawingSheet, "AddShape", "插入形状/文本框").call(
                      drawingSheet,
                      String(args.shapeType || "rect"),
                      mmToEmu(Math.max(1, asFinite(args.widthMm, 60))),
                      mmToEmu(Math.max(1, asFinite(args.heightMm, 35))),
                      createFill(args.fill || { type: "solid", color: "#FFFFFF" }),
                      createStroke(args.line || { widthPt: 1, color: "#000000" }),
                      Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                      mmToEmu(args.columnOffsetMm || 0),
                      Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                      mmToEmu(args.rowOffsetMm || 0)
                    );
                    if (args.text !== undefined) {
                      var shapeContent = safeCall(drawing, "GetDocContent") || safeCall(drawing, "GetContent");
                      var shapeParagraph = safeCall(shapeContent, "GetElement", 0);
                      if (shapeParagraph && typeof shapeParagraph.AddText === "function") shapeParagraph.AddText(String(args.text));
                      else throw new Error("当前 ONLYOFFICE 版本不支持设置形状文本");
                    }
                  } else if (drawingAction === "addOleObject") {
                    if (!args._image || !args._image.url || !args.data || !args.appId) {
                      throw new Error("addOleObject 需要 source、data 和 appId");
                    }
                    drawing = requireMethod(drawingSheet, "AddOleObject", "嵌入对象").call(
                      drawingSheet,
                      String(args._image.url),
                      mmToEmu(Math.max(1, asFinite(args.widthMm, 60))),
                      mmToEmu(Math.max(1, asFinite(args.heightMm, 35))),
                      String(args.data),
                      String(args.appId),
                      Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                      mmToEmu(args.columnOffsetMm || 0),
                      Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                      mmToEmu(args.rowOffsetMm || 0)
                    );
                  } else if (drawingAction === "update" || drawingAction === "delete" || drawingAction === "copy") {
                    var resolvedDrawing = resolveDrawing(args);
                    drawingSheet = resolvedDrawing.sheet;
                    drawing = resolvedDrawing.drawing;
                    managedDrawingIndex = resolvedDrawing.index;
                    if (drawingAction === "delete") {
                      requireMethod(drawing, "Delete", "删除对象").call(drawing);
                    } else if (drawingAction === "copy") {
                      var detachedCopy = requireMethod(drawing, "Copy", "复制对象").call(drawing);
                      var copyDrawingSheet = getSheet(args.destinationSheet || args.sheet);
                      drawing = requireMethod(copyDrawingSheet, "AddDrawing", "粘贴复制对象").call(
                        copyDrawingSheet,
                        detachedCopy,
                        Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                        mmToEmu(args.columnOffsetMm || 0),
                        Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                        mmToEmu(args.rowOffsetMm || 0)
                      );
                      drawingSheet = copyDrawingSheet;
                    }
                  } else {
                    throw new Error("不支持的对象操作：" + drawingAction);
                  }
                  if (drawing && drawingAction !== "delete") {
                    setDrawingPosition(drawing, args);
                    var currentDrawings = typeof drawingSheet.GetAllDrawings === "function" ? drawingSheet.GetAllDrawings() : [];
                    managedDrawingIndex = currentDrawings.indexOf(drawing);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: drawingAction,
                    sheet: drawingSheet.GetName(),
                    drawing: drawingAction === "delete" ? null : describeDrawing(drawing, managedDrawingIndex, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "sheets_manage_hyperlink": {
                  if (!args.action || !args.range) throw new Error("sheets_manage_hyperlink 需要 action 和 range");
                  var hyperlinkSheet = getSheet(args.sheet);
                  var hyperlinkAction = String(args.action);
                  if (hyperlinkAction === "set") {
                    if (!args.url && !args.location) throw new Error("set 需要 url 或 location");
                    requireMethod(hyperlinkSheet, "SetHyperlink", "超链接").call(
                      hyperlinkSheet,
                      String(args.range),
                      args.url ? String(args.url) : null,
                      args.location ? String(args.location) : null,
                      args.displayText !== undefined ? String(args.displayText) : null,
                      args.tooltip !== undefined ? String(args.tooltip) : null
                    );
                  } else if (hyperlinkAction === "delete") {
                    requireMethod(getRange(hyperlinkSheet, args.range), "ClearHyperlinks", "删除超链接").call(getRange(hyperlinkSheet, args.range));
                  } else {
                    throw new Error("超链接 action 必须是 set 或 delete");
                  }
                  changed += 1;
                  results.push({ name: call.name, action: hyperlinkAction, sheet: hyperlinkSheet.GetName(), range: String(args.range) });
                  break;
                }

                case "sheets_inspect_comments": {
                  var commentSheets = args.sheet ? [getSheet(args.sheet)] : Api.GetSheets();
                  var inspectedCommentSheets = [];
                  for (var commentSheetIndex = 0; commentSheetIndex < commentSheets.length; commentSheetIndex += 1) {
                    var inspectedCommentSheet = commentSheets[commentSheetIndex];
                    var sheetComments = typeof inspectedCommentSheet.GetComments === "function" ? inspectedCommentSheet.GetComments() : [];
                    inspectedCommentSheets.push({
                      sheet: inspectedCommentSheet.GetName(),
                      comments: arrayValue(sheetComments).map(describeComment),
                    });
                  }
                  results.push({ name: call.name, sheets: inspectedCommentSheets });
                  break;
                }

                case "sheets_manage_comments": {
                  if (!args.action) throw new Error("sheets_manage_comments.action 不能为空");
                  var commentAction = String(args.action);
                  var managedCommentSheet = getSheet(args.sheet);
                  var managedComment;
                  if (commentAction === "add") {
                    if (!args.range || args.text === undefined) throw new Error("add 需要 range 和 text");
                    managedComment = requireMethod(getRange(managedCommentSheet, args.range), "AddComment", "添加批注").call(
                      getRange(managedCommentSheet, args.range),
                      String(args.text)
                    );
                  } else {
                    var resolvedComment = getComment(args);
                    managedCommentSheet = resolvedComment.sheet;
                    managedComment = resolvedComment.comment;
                    if (commentAction === "update") {
                      if (args.text !== undefined) requireMethod(managedComment, "SetText", "更新批注").call(managedComment, String(args.text));
                      if (args.author !== undefined) requireMethod(managedComment, "SetAuthorName", "批注作者").call(managedComment, String(args.author));
                      if (args.userId !== undefined) requireMethod(managedComment, "SetUserId", "批注用户").call(managedComment, String(args.userId));
                    } else if (commentAction === "delete") requireMethod(managedComment, "Delete", "删除批注").call(managedComment);
                    else if (commentAction === "setSolved") requireMethod(managedComment, "SetSolved", "解决批注").call(managedComment, Boolean(args.solved));
                    else if (commentAction === "addReply") {
                      if (args.text === undefined) throw new Error("addReply 需要 text");
                      requireMethod(managedComment, "AddReply", "批注回复").call(
                        managedComment,
                        String(args.text),
                        args.author !== undefined ? String(args.author) : "",
                        args.userId !== undefined ? String(args.userId) : ""
                      );
                    } else if (commentAction === "removeReplies") requireMethod(managedComment, "RemoveReplies", "删除批注回复").call(
                      managedComment,
                      Number(args.start || 0),
                      Number(args.count || 1),
                      Boolean(args.removeAll)
                    );
                    else throw new Error("不支持的批注操作：" + commentAction);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: commentAction,
                    sheet: managedCommentSheet.GetName(),
                    comment: commentAction === "delete" ? null : describeComment(managedComment),
                  });
                  break;
                }

                case "sheets_inspect_freeze_panes": {
                  var freezeSheet = getSheet(args.sheet);
                  var freezePanes = requireMethod(freezeSheet, "GetFreezePanes", "冻结窗格").call(freezeSheet);
                  results.push({
                    name: call.name,
                    sheet: freezeSheet.GetName(),
                    location: describeRange(safeCall(freezePanes, "GetLocation"), false),
                  });
                  break;
                }

                case "sheets_manage_freeze_panes": {
                  if (!args.action) throw new Error("sheets_manage_freeze_panes.action 不能为空");
                  var managedFreezeSheet = getSheet(args.sheet);
                  var managedFreeze = requireMethod(managedFreezeSheet, "GetFreezePanes", "冻结窗格").call(managedFreezeSheet);
                  var freezeAction = String(args.action);
                  if (freezeAction === "unfreeze") requireMethod(managedFreeze, "Unfreeze", "取消冻结窗格").call(managedFreeze);
                  else if (freezeAction === "freezeAt") {
                    if (!args.range) throw new Error("freezeAt 需要 range");
                    requireMethod(managedFreeze, "FreezeAt", "按区域冻结窗格").call(managedFreeze, getRange(managedFreezeSheet, args.range));
                  } else if (freezeAction === "freezeRows") {
                    requireMethod(managedFreeze, "FreezeRows", "冻结行").call(managedFreeze, Math.max(0, Math.round(asFinite(args.count, 1))));
                  } else if (freezeAction === "freezeColumns") {
                    requireMethod(managedFreeze, "FreezeColumns", "冻结列").call(managedFreeze, Math.max(0, Math.round(asFinite(args.count, 1))));
                  } else throw new Error("不支持的冻结窗格操作：" + freezeAction);
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: freezeAction,
                    sheet: managedFreezeSheet.GetName(),
                    location: describeRange(safeCall(managedFreeze, "GetLocation"), false),
                  });
                  break;
                }

                case "sheets_inspect_properties": {
                  var core = requireMethod(Api, "GetCore", "工作簿属性").call(Api);
                  var coreFields = [
                    "Category", "ContentStatus", "Created", "Creator", "Description", "Identifier", "Keywords",
                    "Language", "LastModifiedBy", "LastPrinted", "Modified", "Revision", "Subject", "Title", "Version",
                  ];
                  var coreData = {};
                  for (var coreFieldIndex = 0; coreFieldIndex < coreFields.length; coreFieldIndex += 1) {
                    coreData[coreFields[coreFieldIndex].charAt(0).toLowerCase() + coreFields[coreFieldIndex].slice(1)] =
                      isoValue(safeCall(core, "Get" + coreFields[coreFieldIndex]));
                  }
                  var requestedCustomProperties = {};
                  var customProperties = typeof Api.GetCustomProperties === "function" ? Api.GetCustomProperties() : null;
                  if (customProperties && Array.isArray(args.customNames)) {
                    for (var customNameIndex = 0; customNameIndex < args.customNames.length; customNameIndex += 1) {
                      requestedCustomProperties[String(args.customNames[customNameIndex])] =
                        isoValue(safeCall(customProperties, "Get", String(args.customNames[customNameIndex])));
                    }
                  }
                  results.push({ name: call.name, core: coreData, custom: requestedCustomProperties });
                  break;
                }

                case "sheets_manage_properties": {
                  if (!args.core && !args.custom) throw new Error("sheets_manage_properties 需要 core 或 custom");
                  var managedCore = requireMethod(Api, "GetCore", "工作簿属性").call(Api);
                  if (isObject(args.core)) {
                    var propertyMap = {
                      category: "Category", contentStatus: "ContentStatus", created: "Created", creator: "Creator",
                      description: "Description", identifier: "Identifier", keywords: "Keywords", language: "Language",
                      lastModifiedBy: "LastModifiedBy", lastPrinted: "LastPrinted", modified: "Modified", revision: "Revision",
                      subject: "Subject", title: "Title", version: "Version",
                    };
                    Object.keys(args.core).forEach(function (propertyKey) {
                      if (!propertyMap[propertyKey]) throw new Error("不支持的核心属性：" + propertyKey);
                      var propertyValue = args.core[propertyKey];
                      if (propertyKey === "created" || propertyKey === "lastPrinted" || propertyKey === "modified") {
                        propertyValue = new Date(propertyValue);
                        if (isNaN(propertyValue.getTime())) throw new Error(propertyKey + " 不是有效日期");
                      }
                      requireMethod(managedCore, "Set" + propertyMap[propertyKey], "设置工作簿属性").call(managedCore, propertyValue);
                    });
                  }
                  if (isObject(args.custom)) {
                    var managedCustom = requireMethod(Api, "GetCustomProperties", "自定义工作簿属性").call(Api);
                    Object.keys(args.custom).forEach(function (customKey) {
                      requireMethod(managedCustom, "Add", "设置自定义工作簿属性").call(managedCustom, customKey, args.custom[customKey]);
                    });
                  }
                  changed += 1;
                  results.push({ name: call.name, coreKeys: Object.keys(args.core || {}), customKeys: Object.keys(args.custom || {}) });
                  break;
                }

                case "sheets_inspect_protected_ranges": {
                  var protectionSheets = args.sheet ? [getSheet(args.sheet)] : Api.GetSheets();
                  var protectedSheetResults = [];
                  for (var protectedSheetIndex = 0; protectedSheetIndex < protectionSheets.length; protectedSheetIndex += 1) {
                    var inspectedProtectedSheet = protectionSheets[protectedSheetIndex];
                    var protectedRanges = typeof inspectedProtectedSheet.GetAllProtectedRanges === "function"
                      ? inspectedProtectedSheet.GetAllProtectedRanges()
                      : [];
                    protectedSheetResults.push({
                      sheet: inspectedProtectedSheet.GetName(),
                      protectedRanges: arrayValue(protectedRanges).map(function (protectedRange) {
                        return {
                          title: safeCall(protectedRange, "GetTitle"),
                          range: safeCall(protectedRange, "GetRange"),
                          anyoneType: safeCall(protectedRange, "GetAnyoneType"),
                          users: arrayValue(safeCall(protectedRange, "GetAllUsers")).map(function (user) {
                            return {
                              id: safeCall(user, "GetId"),
                              name: safeCall(user, "GetName"),
                              type: safeCall(user, "GetType"),
                            };
                          }),
                        };
                      }),
                    });
                  }
                  results.push({ name: call.name, sheets: protectedSheetResults });
                  break;
                }

                case "sheets_manage_protected_ranges": {
                  if (!args.action || !args.title) throw new Error("sheets_manage_protected_ranges 需要 action 和 title");
                  var protectedSheet = getSheet(args.sheet);
                  var protectedAction = String(args.action);
                  var protectedRange;
                  if (protectedAction === "add") {
                    if (!args.range) throw new Error("add 需要 range");
                    protectedRange = requireMethod(protectedSheet, "AddProtectedRange", "受保护区域").call(
                      protectedSheet,
                      String(args.title),
                      qualifyRange(protectedSheet, args.range)
                    );
                    if (!protectedRange && typeof protectedSheet.GetProtectedRange === "function") {
                      protectedRange = protectedSheet.GetProtectedRange(String(args.title));
                    }
                  } else {
                    protectedRange = requireMethod(protectedSheet, "GetProtectedRange", "受保护区域").call(protectedSheet, String(args.title));
                    if (!protectedRange) throw new Error("找不到受保护区域：" + args.title);
                    if (protectedAction === "update") {
                      if (args.newTitle) requireMethod(protectedRange, "SetTitle", "重命名受保护区域").call(protectedRange, String(args.newTitle));
                      if (args.range) requireMethod(protectedRange, "SetRange", "更新受保护区域").call(protectedRange, qualifyRange(protectedSheet, args.range));
                      if (args.anyoneType) requireMethod(protectedRange, "SetAnyoneType", "受保护区域默认权限").call(protectedRange, String(args.anyoneType));
                    } else if (protectedAction === "addUser") {
                      if (!args.userId || !args.userName || !args.permission) throw new Error("addUser 需要 userId、userName 和 permission");
                      requireMethod(protectedRange, "AddUser", "受保护区域用户权限").call(
                        protectedRange,
                        String(args.userId),
                        String(args.userName),
                        String(args.permission)
                      );
                    } else if (protectedAction === "deleteUser") {
                      if (!args.userId) throw new Error("deleteUser 需要 userId");
                      requireMethod(protectedRange, "DeleteUser", "删除受保护区域用户").call(protectedRange, String(args.userId));
                    } else throw new Error("不支持的受保护区域操作：" + protectedAction);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: protectedAction,
                    sheet: protectedSheet.GetName(),
                    title: safeCall(protectedRange, "GetTitle") || args.newTitle || args.title,
                  });
                  break;
                }

                case "sheets_inspect_page_layout": {
                  var pageSheet = getSheet(args.sheet);
                  var inspectedMarginsPt = {
                    top: safeCall(pageSheet, "GetTopMargin"),
                    right: safeCall(pageSheet, "GetRightMargin"),
                    bottom: safeCall(pageSheet, "GetBottomMargin"),
                    left: safeCall(pageSheet, "GetLeftMargin"),
                  };
                  results.push({
                    name: call.name,
                    sheet: pageSheet.GetName(),
                    orientation: safeCall(pageSheet, "GetPageOrientation"),
                    marginsPt: inspectedMarginsPt,
                    margins: inspectedMarginsPt,
                    printGridlines: safeCall(pageSheet, "GetPrintGridlines"),
                    printHeadings: safeCall(pageSheet, "GetPrintHeadings"),
                  });
                  break;
                }

                case "sheets_manage_page_layout": {
                  var managedMarginsPt = unitAlias(args, "marginsPt", "margins");
                  if (
                    args.orientation === undefined
                    && !isObject(managedMarginsPt)
                    && !hasOwn(args, "printGridlines")
                    && !hasOwn(args, "printHeadings")
                  ) {
                    throw new Error("sheets_manage_page_layout 至少需要一个页面布局属性");
                  }
                  var managedPageSheet = getSheet(args.sheet);
                  if (args.orientation !== undefined) requireMethod(managedPageSheet, "SetPageOrientation", "页面方向").call(managedPageSheet, String(args.orientation));
                  if (isObject(managedMarginsPt)) {
                    if (hasOwn(managedMarginsPt, "top")) requireMethod(managedPageSheet, "SetTopMargin", "上页边距").call(managedPageSheet, Number(managedMarginsPt.top));
                    if (hasOwn(managedMarginsPt, "right")) requireMethod(managedPageSheet, "SetRightMargin", "右页边距").call(managedPageSheet, Number(managedMarginsPt.right));
                    if (hasOwn(managedMarginsPt, "bottom")) requireMethod(managedPageSheet, "SetBottomMargin", "下页边距").call(managedPageSheet, Number(managedMarginsPt.bottom));
                    if (hasOwn(managedMarginsPt, "left")) requireMethod(managedPageSheet, "SetLeftMargin", "左页边距").call(managedPageSheet, Number(managedMarginsPt.left));
                  }
                  if (hasOwn(args, "printGridlines")) requireMethod(managedPageSheet, "SetPrintGridlines", "打印网格线").call(managedPageSheet, Boolean(args.printGridlines));
                  if (hasOwn(args, "printHeadings")) requireMethod(managedPageSheet, "SetPrintHeadings", "打印标题").call(managedPageSheet, Boolean(args.printHeadings));
                  changed += 1;
                  var persistedMarginsPt = {
                    top: safeCall(managedPageSheet, "GetTopMargin"),
                    right: safeCall(managedPageSheet, "GetRightMargin"),
                    bottom: safeCall(managedPageSheet, "GetBottomMargin"),
                    left: safeCall(managedPageSheet, "GetLeftMargin"),
                  };
                  results.push({
                    name: call.name,
                    sheet: managedPageSheet.GetName(),
                    orientation: safeCall(managedPageSheet, "GetPageOrientation"),
                    marginsPt: persistedMarginsPt,
                    margins: persistedMarginsPt,
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
