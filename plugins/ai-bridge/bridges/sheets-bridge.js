(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) {
      const rawError = value && value.error;
      const error = new Error(
        rawError && typeof rawError === "object"
          ? rawError.message || "Sheets Bridge 执行失败"
          : (value && value.message) || rawError || "Sheets Bridge 执行失败"
      );
      error.code = value && value.code
        ? value.code
        : (rawError && rawError.code ? rawError.code : "EXECUTION_FAILED");
      const rawDetails = value && value.details && typeof value.details === "object"
        ? value.details
        : (
          rawError && rawError.details && typeof rawError.details === "object"
            ? rawError.details
            : {}
        );
      error.details = { ...rawDetails };
      if (!error.details.phase) error.details.phase = "sheets-command";
      throw error;
    }
    return value;
  }

  const VIEW_VERIFY_INTERVAL_MS = 100;
  const VIEW_VERIFY_TIMEOUT_MS = 5000;

  function staticA1RangeShape(value) {
    if (typeof value !== "string" || value.toLowerCase() === "selection") return null;
    const localAddress = value.split("!").pop().replace(/\$/g, "");
    const matched = /^([A-Za-z]{1,3})([1-9][0-9]*)(?::([A-Za-z]{1,3})([1-9][0-9]*))?$/.exec(localAddress);
    if (!matched) return null;
    function columnNumber(label) {
      let number = 0;
      for (let index = 0; index < label.length; index += 1) {
        number = number * 26 + label.toUpperCase().charCodeAt(index) - 64;
      }
      return number;
    }
    const firstColumn = columnNumber(matched[1]);
    const firstRow = Number(matched[2]);
    const lastColumn = columnNumber(matched[3] || matched[1]);
    const lastRow = Number(matched[4] || matched[2]);
    if (
      firstColumn > 16384
      || lastColumn > 16384
      || firstRow > 1048576
      || lastRow > 1048576
    ) return null;
    return {
      rows: Math.abs(lastRow - firstRow) + 1,
      columns: Math.abs(lastColumn - firstColumn) + 1,
    };
  }

  function canonicalStaticFilterOperator(value) {
    const aliases = {
      and: "xlAnd",
      or: "xlOr",
      filterValues: "xlFilterValues",
      values: "xlFilterValues",
      top10Items: "xlTop10Items",
      bottom10Items: "xlBottom10Items",
      top10Percent: "xlTop10Percent",
      bottom10Percent: "xlBottom10Percent",
      filterCellColor: "xlFilterCellColor",
      filterFontColor: "xlFilterFontColor",
      filterIcon: "xlFilterIcon",
      dynamic: "xlFilterDynamic",
    };
    return aliases[value] || value;
  }

  function staticFilterScalar(value) {
    return typeof value === "string"
      || typeof value === "boolean"
      || (typeof value === "number" && Number.isFinite(value));
  }

  function staticFilterCriterion(value, allowArray) {
    return staticFilterScalar(value) || Boolean(
      allowArray
      && Array.isArray(value)
      && value.length
      && value.every(staticFilterScalar)
    );
  }

  function collectFilterValidationErrors(toolCalls) {
    const validationErrors = [];
    const calls = Array.isArray(toolCalls) ? toolCalls : [];
    function semanticError(toolCallIndex, path, message) {
      validationErrors.push({
        toolCallIndex,
        tool: "sheets_filter",
        path: `arguments.${path}`,
        keyword: "semantic",
        message,
      });
    }
    calls.forEach(function (call, toolCallIndex) {
      if (!call || call.name !== "sheets_filter") return;
      const args = call.arguments || call.args || {};
      const action = args.action;
      if (action === "set") {
        if (typeof args.range === "string" && args.range.toLowerCase() !== "selection") {
          const shape = staticA1RangeShape(args.range);
          if (!shape) {
            semanticError(toolCallIndex, "range", "range must be a finite A1 cell range or selection");
          } else {
            if (shape.rows < 2) {
              semanticError(toolCallIndex, "range", "filter range must contain a header row and at least one data row");
            }
            if (Number.isInteger(args.field) && args.field > shape.columns) {
              semanticError(toolCallIndex, "field", `field must not exceed the filter range width (${shape.columns})`);
            }
          }
        }
        const criteria1Present = Object.prototype.hasOwnProperty.call(args, "criteria1");
        const criteria2Present = Object.prototype.hasOwnProperty.call(args, "criteria2");
        const operator = canonicalStaticFilterOperator(args.operator);
        if (criteria1Present && !staticFilterCriterion(args.criteria1, true)) {
          semanticError(toolCallIndex, "criteria1", "criteria1 must be a scalar or a non-empty array of scalar values");
        }
        if (criteria2Present && !staticFilterCriterion(args.criteria2, false)) {
          semanticError(toolCallIndex, "criteria2", "criteria2 must be a scalar value");
        }
        if (args.operator !== undefined && !criteria1Present) {
          semanticError(toolCallIndex, "criteria1", "criteria1 is required when operator is provided");
        }
        if (criteria2Present && operator !== "xlAnd" && operator !== "xlOr") {
          semanticError(toolCallIndex, "criteria2", "criteria2 is only allowed with xlAnd or xlOr");
        }
        if (criteria2Present && !criteria1Present) {
          semanticError(toolCallIndex, "criteria1", "criteria1 is required when criteria2 is provided");
        }
        if (operator === "xlFilterValues" && !Array.isArray(args.criteria1)) {
          semanticError(toolCallIndex, "criteria1", "xlFilterValues requires a non-empty criteria1 array");
        }
        if (Array.isArray(args.criteria1) && operator !== "xlFilterValues") {
          semanticError(toolCallIndex, "operator", "array criteria1 is only allowed with xlFilterValues");
        }
        if ([
          "xlTop10Items",
          "xlBottom10Items",
          "xlTop10Percent",
          "xlBottom10Percent",
        ].indexOf(operator) >= 0) {
          const ranking = Number(args.criteria1);
          if (!Number.isFinite(ranking) || ranking <= 0) {
            semanticError(toolCallIndex, "criteria1", "top/bottom filters require a positive numeric criteria1");
          }
        }
        if (["xlFilterCellColor", "xlFilterFontColor", "xlFilterIcon"].indexOf(operator) >= 0) {
          semanticError(toolCallIndex, "operator", "color and icon filters are not safe through the JSON bridge");
        }
      } else if (action === "showAll" || action === "reapply") {
        const setOnlyFields = [
          "range",
          "field",
          "criteria1",
          "operator",
          "criteria2",
          "visibleDropDown",
        ];
        const invalidField = setOnlyFields.find(function (field) {
          return Object.prototype.hasOwnProperty.call(args, field);
        });
        if (invalidField) {
          semanticError(toolCallIndex, invalidField, `${invalidField} is only allowed when action is set`);
        }
      }
    });
    return validationErrors;
  }

  function preflight(toolCalls) {
    const calls = Array.isArray(toolCalls) ? toolCalls : [];
    return {
      toolCalls: calls,
      argumentNormalizations: [],
      validationErrors: collectFilterValidationErrors(calls),
    };
  }

  function probeCapabilities() {
    return new Promise(function (resolve) {
      Asc.plugin.callCommand(
        function () {
          try {
            var sheet = Api.GetActiveSheet();
            var probeRange = sheet && typeof sheet.GetRange === "function"
              ? sheet.GetRange("A1")
              : null;
            var charts = sheet && typeof sheet.GetAllCharts === "function"
              ? sheet.GetAllCharts()
              : [];
            var firstChart = Array.isArray(charts) && charts.length ? charts[0] : null;
            var chartDelete = false;
            if (Array.isArray(charts)) {
              for (var chartIndex = 0; chartIndex < charts.length; chartIndex += 1) {
                if (charts[chartIndex] && typeof charts[chartIndex].Delete === "function") {
                  chartDelete = true;
                  break;
                }
              }
            }
            var validation = probeRange && typeof probeRange.GetValidation === "function"
              ? probeRange.GetValidation()
              : null;
            var freezePanes = sheet && typeof sheet.GetFreezePanes === "function"
              ? sheet.GetFreezePanes()
              : null;
            var comments = sheet && typeof sheet.GetComments === "function"
              ? sheet.GetComments()
              : [];
            var firstComment = Array.isArray(comments) && comments.length ? comments[0] : null;
            var version = null;
            try {
              if (
                typeof Asc !== "undefined"
                && Asc.editor
                && typeof Asc.editor.asc_getVersion === "function"
              ) {
                version = String(Asc.editor.asc_getVersion() || "") || null;
              }
            } catch (versionError) {
              version = null;
            }
            return JSON.stringify({
              ok: true,
              runtime: {
                product: "ONLYOFFICE",
                version: version,
                edition: "unknown",
              },
              features: {
                sheets: {
                  nativeTables: {
                    create: Boolean(sheet && typeof sheet.AddListObject === "function"),
                    inspect: Boolean(sheet && typeof sheet.GetListObjects === "function"),
                  },
                  rangeStyleTables: {
                    create: Boolean(
                      probeRange
                      && typeof probeRange.SetBold === "function"
                      && typeof probeRange.SetFillColor === "function"
                      && typeof probeRange.SetBorders === "function"
                    ),
                  },
                  conditionalFormatting: {
                    create: Boolean(
                      probeRange
                      && typeof probeRange.GetFormatConditions === "function"
                    ),
                  },
                  rangeFill: {
                    fillDown: Boolean(probeRange && typeof probeRange.FillDown === "function"),
                    fillUp: Boolean(probeRange && typeof probeRange.FillUp === "function"),
                    fillLeft: Boolean(probeRange && typeof probeRange.FillLeft === "function"),
                    fillRight: Boolean(probeRange && typeof probeRange.FillRight === "function"),
                  },
                  arrayFormula: {
                    set: Boolean(probeRange && typeof probeRange.SetFormulaArray === "function"),
                  },
                  validation: {
                    manage: Boolean(
                      validation
                      && typeof validation.Add === "function"
                      && typeof validation.Modify === "function"
                      && typeof validation.Delete === "function"
                    ),
                  },
                  comments: {
                    create: Boolean(probeRange && typeof probeRange.AddComment === "function"),
                    inspect: Boolean(sheet && typeof sheet.GetComments === "function"),
                    update: Boolean(firstComment && typeof firstComment.SetText === "function"),
                    delete: Boolean(firstComment && typeof firstComment.Delete === "function"),
                  },
                  freezePanes: {
                    inspect: Boolean(sheet && typeof sheet.GetFreezePanes === "function"),
                    manage: Boolean(
                      freezePanes
                      && typeof freezePanes.Unfreeze === "function"
                      && typeof freezePanes.FreezeAt === "function"
                      && typeof freezePanes.FreezeRows === "function"
                      && typeof freezePanes.FreezeColumns === "function"
                    ),
                  },
                  charts: {
                    create: Boolean(sheet && typeof sheet.AddChart === "function"),
                    inspect: Boolean(sheet && typeof sheet.GetAllCharts === "function"),
                    update: Boolean(firstChart && typeof firstChart.SetTitle === "function"),
                    // Adding series while a chart is being created is kept
                    // conservative because affected Community builds expose
                    // AddSeria but fail inside createDuplicate.
                    addSeriesOnCreate: false,
                    delete: chartDelete,
                  },
                },
              },
            });
          } catch (error) {
            return JSON.stringify({
              ok: false,
              error: error && error.message ? error.message : String(error),
            });
          }
        },
        false,
        true,
        function (rawResult) {
          try {
            var result = typeof rawResult === "string"
              ? JSON.parse(rawResult || "{}")
              : rawResult;
            if (!result || result.ok !== true) throw new Error(result && result.error || "能力探测失败");
            resolve({ runtime: result.runtime, features: result.features });
          } catch (error) {
            resolve({
              runtime: { product: "ONLYOFFICE", version: null, edition: "unknown" },
              features: {
                sheets: {
                  nativeTables: { create: false, inspect: false },
                  rangeStyleTables: { create: false },
                  conditionalFormatting: { create: false },
                  rangeFill: { fillDown: false, fillUp: false, fillLeft: false, fillRight: false },
                  arrayFormula: { set: false },
                  validation: { manage: false },
                  comments: { create: false, inspect: false, update: false, delete: false },
                  freezePanes: { inspect: false, manage: false },
                  charts: { create: false, inspect: false, update: false, addSeriesOnCreate: false, delete: false },
                },
              },
            });
          }
        },
      );
    });
  }

  function a1ColumnNumber(column) {
    let number = 0;
    const normalized = String(column || "").toUpperCase();
    for (let index = 0; index < normalized.length; index += 1) {
      number = number * 26 + normalized.charCodeAt(index) - 64;
    }
    return number;
  }

  function a1LastCell(address) {
    const localAddress = String(address || "")
      .split("!")
      .pop()
      .replace(/\$/g, "");
    const parts = localAddress.split(":");
    const matched = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(parts[parts.length - 1]);
    if (!matched) return null;
    return {
      column: a1ColumnNumber(matched[1]),
      row: Number(matched[2]),
    };
  }

  function viewChecksFor(toolCalls, executionResult) {
    const resultItems = Array.isArray(executionResult.results) ? executionResult.results : [];
    const checks = [];
    toolCalls.forEach(function (call, resultIndex) {
      const args = (call && (call.arguments || call.args)) || {};
      const resultItem = resultItems[resultIndex] || {};
      if (call && call.name === "sheets_manage_freeze_panes") {
        const action = String(args.action || "");
        let expected;
        if (action === "unfreeze") {
          expected = { frozenRows: 0, frozenColumns: 0 };
        } else if (action === "freezeRows") {
          expected = { frozenRows: Number(args.count), frozenColumns: 0 };
        } else if (action === "freezeColumns") {
          expected = { frozenRows: 0, frozenColumns: Number(args.count) };
        } else {
          const lastCell = a1LastCell(args.range);
          if (!lastCell) return;
          expected = {
            frozenRows: lastCell.row,
            frozenColumns: lastCell.column,
          };
        }
        checks.push({
          kind: "freeze",
          name: call.name,
          resultIndex: resultIndex,
          sheet: resultItem.sheet || args.sheet || "",
          expected: expected,
        });
      } else if (
        call
        && call.name === "sheets_manage_page_layout"
        && (
          Object.prototype.hasOwnProperty.call(args, "displayGridlines")
          || Object.prototype.hasOwnProperty.call(args, "displayHeadings")
        )
      ) {
        const expected = {};
        if (Object.prototype.hasOwnProperty.call(args, "displayGridlines")) {
          expected.displayGridlines = Boolean(args.displayGridlines);
        }
        if (Object.prototype.hasOwnProperty.call(args, "displayHeadings")) {
          expected.displayHeadings = Boolean(args.displayHeadings);
        }
        checks.push({
          kind: "pageLayout",
          name: call.name,
          resultIndex: resultIndex,
          sheet: resultItem.sheet || args.sheet || "",
          expected: expected,
        });
      }
    });
    toolCalls.forEach(function (call, resultIndex) {
      if (
        !call
        || (
          call.name !== "sheets_inspect_freeze_panes"
          && call.name !== "sheets_inspect_page_layout"
        )
      ) {
        return;
      }
      const resultItem = resultItems[resultIndex] || {};
      const kind = call.name === "sheets_inspect_freeze_panes" ? "freeze" : "pageLayout";
      const sheetName = resultItem.sheet || "";
      const priorCheck = checks.slice().reverse().find(function (check) {
        return check.resultIndex < resultIndex
          && check.kind === kind
          && check.sheet === sheetName;
      });
      if (!priorCheck) return;
      checks.push({
        kind: kind,
        name: call.name,
        resultIndex: resultIndex,
        sheet: sheetName,
        expected: priorCheck.expected,
      });
    });
    checks.sort(function (left, right) {
      return left.resultIndex - right.resultIndex;
    });
    return checks;
  }

  function viewStateMatches(check, observed) {
    if (!observed) return false;
    return Object.keys(check.expected).every(function (key) {
      return observed[key] === check.expected[key];
    });
  }

  function applyVerifiedViewStates(executionResult, checks, states) {
    checks.forEach(function (check, index) {
      const resultItem = executionResult.results[check.resultIndex];
      const state = states[index];
      if (check.kind === "freeze") {
        if (state.locationAddress) {
          resultItem.location = { address: state.locationAddress };
          if (state.locationRows !== undefined) {
            resultItem.location.rows = state.locationRows;
          }
          if (state.locationColumns !== undefined) {
            resultItem.location.columns = state.locationColumns;
          }
        } else {
          resultItem.location = null;
        }
        resultItem.frozenRows = state.frozenRows;
        resultItem.frozenColumns = state.frozenColumns;
        resultItem.topLeftCell = state.topLeftCell;
      } else {
        resultItem.displayGridlines = state.displayGridlines;
        resultItem.displayHeadings = state.displayHeadings;
      }
      resultItem.verified = true;
    });
    return executionResult;
  }

  function verificationError(checks, states) {
    const failedIndex = checks.findIndex(function (check, index) {
      return !viewStateMatches(check, states[index]);
    });
    const failedCheck = checks[Math.max(0, failedIndex)] || {};
    const error = new Error("ONLYOFFICE 工作表视图状态未在保存前生效");
    error.code = "SHEETS_VIEW_STATE_NOT_APPLIED";
    error.details = {
      phase: "sheets-view-verification",
      tool: failedCheck.name,
      toolCallIndex: failedCheck.resultIndex,
      completedToolCalls: checks.length,
      partialMutationPossible: true,
      expected: failedCheck.expected,
      observed: states[Math.max(0, failedIndex)] || null,
    };
    return error;
  }

  function verificationReadError(checks, payload) {
    const failedCheck = checks[0] || {};
    const payloadDetails = payload
      && payload.details
      && typeof payload.details === "object"
      ? payload.details
      : {};
    const error = new Error(
      payload && payload.message
        ? payload.message
        : "无法读取 ONLYOFFICE 工作表视图状态",
    );
    error.code = payload && payload.code
      ? payload.code
      : "SHEETS_VIEW_STATE_READ_FAILED";
    error.details = Object.assign({}, payloadDetails, {
      phase: "sheets-view-verification",
      tool: failedCheck.name,
      toolCallIndex: failedCheck.resultIndex,
      completedToolCalls: checks.length,
      partialMutationPossible: true,
      expected: failedCheck.expected,
      observed: payloadDetails.location || null,
    });
    return error;
  }

  function markSheetViewModified(checks) {
    return new Promise(function (resolve, reject) {
      Asc.plugin.callCommand(
        function () {
          try {
            var editor = typeof Asc !== "undefined" ? Asc.editor : null;
            var markModified = editor && (
              typeof editor.SetDocumentModified === "function"
                ? editor.SetDocumentModified
                : editor.onUpdateDocumentModified
            );
            if (typeof markModified === "function") {
              markModified.call(editor, true);
              return JSON.stringify({ ok: true });
            }
            var activeSheet = typeof Api !== "undefined" && typeof Api.GetActiveSheet === "function"
              ? Api.GetActiveSheet()
              : null;
            var workbook = activeSheet
              && activeSheet.worksheet
              && activeSheet.worksheet.workbook;
            var workbookApi = workbook && workbook.oApi;
            if (workbookApi && typeof workbookApi.SetDocumentModified === "function") {
              workbookApi.SetDocumentModified(true);
              return JSON.stringify({ ok: true });
            }
            if (workbook && workbook.handlers && typeof workbook.handlers.trigger === "function") {
              workbook.handlers.trigger("setDocumentModified", true);
              return JSON.stringify({ ok: true });
            }
            throw new Error("当前 ONLYOFFICE 版本无法标记工作簿视图变更");
          } catch (error) {
            return JSON.stringify({
              ok: false,
              message: error && error.message ? error.message : String(error),
            });
          }
        },
        false,
        true,
        function (rawResult) {
          let payload;
          try {
            payload = typeof rawResult === "string"
              ? JSON.parse(rawResult || "{}")
              : rawResult;
          } catch (error) {
            reject(error);
            return;
          }
          if (!payload || !payload.ok) {
            const error = verificationError(checks, []);
            error.message = payload && payload.message
              ? payload.message
              : "无法标记工作簿视图变更";
            reject(error);
            return;
          }
          resolve();
        },
      );
    });
  }

  function verifyViewChanges(toolCalls, executionResult) {
    const checks = viewChecksFor(toolCalls, executionResult);
    if (!checks.length) return Promise.resolve(executionResult);

    const timeoutMs = Number(window.AICopilotSheetsViewVerificationTimeoutMs)
      || VIEW_VERIFY_TIMEOUT_MS;
    const intervalMs = Number(window.AICopilotSheetsViewVerificationIntervalMs)
      || VIEW_VERIFY_INTERVAL_MS;
    const startedAt = Date.now();
    Asc.scope.copilotSheetViewChecks = checks;

    return new Promise(function (resolve, reject) {
      function readStates() {
        Asc.plugin.callCommand(
          function () {
            try {
              var checks = Asc.scope.copilotSheetViewChecks || [];

              function findSheet(name) {
                if (!name) return Api.GetActiveSheet();
                var sheets = Api.GetSheets();
                for (var sheetIndex = 0; sheetIndex < sheets.length; sheetIndex += 1) {
                  if (String(sheets[sheetIndex].GetName()) === String(name)) {
                    return sheets[sheetIndex];
                  }
                }
                throw new Error("找不到工作表：" + name);
              }

              function columnNumber(column) {
                var number = 0;
                var normalized = String(column || "").toUpperCase();
                for (var columnIndex = 0; columnIndex < normalized.length; columnIndex += 1) {
                  number = number * 26 + normalized.charCodeAt(columnIndex) - 64;
                }
                return number;
              }

              function columnName(number) {
                var remaining = Math.max(1, Number(number) || 1);
                var result = "";
                while (remaining > 0) {
                  var remainder = (remaining - 1) % 26;
                  result = String.fromCharCode(65 + remainder) + result;
                  remaining = Math.floor((remaining - 1) / 26);
                }
                return result;
              }

              function freezeDimensions(address, rows, columns) {
                var maxRows = 1048576;
                var maxColumns = 16384;
                var localAddress = String(address || "")
                  .split("!")
                  .pop()
                  .replace(/\$/g, "")
                  .trim();
                var matched = /^([1-9][0-9]*):([1-9][0-9]*)$/.exec(localAddress);
                var frozenRows;
                var frozenColumns;
                if (matched) {
                  frozenRows = Number(matched[2]);
                  frozenColumns = 0;
                } else {
                  matched = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(localAddress);
                  if (matched) {
                    frozenRows = 0;
                    frozenColumns = columnNumber(matched[2]);
                  } else {
                    var parts = localAddress.split(":");
                    matched = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(parts[parts.length - 1]);
                    if (matched) {
                      var lastColumn = columnNumber(matched[1]);
                      var lastRow = Number(matched[2]);
                      frozenRows = lastColumn >= maxColumns ? lastRow : (
                        lastRow >= maxRows ? 0 : lastRow
                      );
                      frozenColumns = lastRow >= maxRows ? lastColumn : (
                        lastColumn >= maxColumns ? 0 : lastColumn
                      );
                    }
                  }
                }
                if (frozenRows === undefined || frozenColumns === undefined) {
                  var rowCount = Number(rows);
                  var columnCount = Number(columns);
                  var validRows = isFinite(rowCount) && rowCount >= 1 && Math.floor(rowCount) === rowCount;
                  var validColumns = isFinite(columnCount) && columnCount >= 1 && Math.floor(columnCount) === columnCount;
                  if (validRows && validColumns) {
                    if (columnCount >= maxColumns && rowCount < maxRows) {
                      frozenRows = rowCount;
                      frozenColumns = 0;
                    } else if (rowCount >= maxRows && columnCount < maxColumns) {
                      frozenRows = 0;
                      frozenColumns = columnCount;
                    } else if (rowCount < maxRows && columnCount < maxColumns) {
                      frozenRows = rowCount;
                      frozenColumns = columnCount;
                    }
                  }
                }
                if (
                  frozenRows === undefined
                  || frozenColumns === undefined
                  || frozenRows < 0
                  || frozenColumns < 0
                  || frozenRows >= maxRows
                  || frozenColumns >= maxColumns
                ) {
                  return null;
                }
                return {
                  frozenRows: frozenRows,
                  frozenColumns: frozenColumns,
                };
              }

              function freezeState(sheet) {
                var panes = sheet.GetFreezePanes();
                var location = panes && typeof panes.GetLocation === "function"
                  ? panes.GetLocation()
                  : null;
                if (!location) {
                  return {
                    locationAddress: null,
                    frozenRows: 0,
                    frozenColumns: 0,
                    topLeftCell: null,
                  };
                }
                var address = location.GetAddress(true, true, "xlA1", false);
                var rows = typeof location.GetRowsCount === "function"
                  ? location.GetRowsCount()
                  : null;
                var columns = typeof location.GetColumnsCount === "function"
                  ? location.GetColumnsCount()
                  : null;
                var dimensions = freezeDimensions(address, rows, columns);
                if (!dimensions) {
                  var parseError = new Error("无法解析 ONLYOFFICE 返回的冻结区域");
                  parseError.code = "SHEETS_API_UNSUPPORTED";
                  parseError.details = {
                    location: { address: address, rows: rows, columns: columns },
                  };
                  throw parseError;
                }
                return {
                  locationAddress: address,
                  locationRows: rows,
                  locationColumns: columns,
                  frozenRows: dimensions.frozenRows,
                  frozenColumns: dimensions.frozenColumns,
                  topLeftCell: columnName(dimensions.frozenColumns + 1)
                    + String(dimensions.frozenRows + 1),
                };
              }

              function displayState(sheet) {
                var model = sheet && sheet.worksheet;
                var settings = model && typeof model.getSheetViewSettings === "function"
                  ? model.getSheetViewSettings(true)
                  : null;
                if (!settings) throw new Error("当前 ONLYOFFICE 版本无法读取工作表屏幕视图");
                return {
                  displayGridlines: settings.showGridLines !== false,
                  displayHeadings: settings.showRowColHeaders !== false,
                };
              }

              var states = checks.map(function (check) {
                var sheet = findSheet(check.sheet);
                return check.kind === "freeze" ? freezeState(sheet) : displayState(sheet);
              });
              return JSON.stringify({ ok: true, states: states });
            } catch (error) {
              return JSON.stringify({
                ok: false,
                code: error && error.code ? error.code : "SHEETS_VIEW_STATE_READ_FAILED",
                message: error && error.message ? error.message : String(error),
                details: error && error.details && typeof error.details === "object"
                  ? error.details
                  : {},
              });
            }
          },
          false,
          true,
          function (rawResult) {
            let payload;
            try {
              payload = typeof rawResult === "string"
                ? JSON.parse(rawResult || "{}")
                : rawResult;
            } catch (error) {
              reject(error);
              return;
            }
            if (!payload || !payload.ok) {
              reject(verificationReadError(checks, payload || {}));
              return;
            }
            const states = Array.isArray(payload.states) ? payload.states : [];
            if (checks.every(function (check, index) {
              return viewStateMatches(check, states[index]);
            })) {
              markSheetViewModified(checks).then(function () {
                resolve(applyVerifiedViewStates(executionResult, checks, states));
              }, reject);
              return;
            }
            if (Date.now() - startedAt >= timeoutMs) {
              reject(verificationError(checks, states));
              return;
            }
            window.setTimeout(readStates, intervalMs);
          },
        );
      }

      readStates();
    });
  }

  function executeMacroTool(toolCalls) {
    return new Promise(function (resolve, reject) {
      function rejectBeforeMutation(message, code) {
        const error = new Error(message);
        error.code = code || "INVALID_TOOL_ARGUMENTS";
        error.details = { partialMutationPossible: false };
        reject(error);
      }
      if (toolCalls.length !== 1) {
        rejectBeforeMutation("宏工具必须单独调用，不能与普通工作表命令混批");
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
          rejectBeforeMutation("sheets_set_macros.content 必须是宏配置对象");
          return;
        }
        method = "SetMacros";
        params = [JSON.stringify(args.content)];
        needsSave = true;
      } else {
        rejectBeforeMutation("未知宏工具：" + call.name);
        return;
      }
      if (!window.Asc || !Asc.plugin || typeof Asc.plugin.executeMethod !== "function") {
        rejectBeforeMutation(
          "当前 ONLYOFFICE 版本不支持宏插件方法",
          "SHEETS_API_UNSUPPORTED"
        );
        return;
      }
      try {
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
              results: [{ name: call.name, kind: method === "GetVBAMacros" ? "vba" : "office", content: content }],
            });
          } catch (error) {
            if (!error.details || typeof error.details !== "object") error.details = {};
            error.details.partialMutationPossible = needsSave;
            reject(error);
          }
        });
      } catch (error) {
        if (!error.details || typeof error.details !== "object") error.details = {};
        error.details.partialMutationPossible = needsSave;
        reject(error);
      }
    });
  }

  function execute(toolCalls) {
    if (toolCalls.some(function (call) {
      return call && (call.name === "sheets_inspect_macros" || call.name === "sheets_set_macros");
    })) {
      return executeMacroTool(toolCalls).catch(function (error) {
        const macroIndex = toolCalls.findIndex(function (call) {
          return call && (
            call.name === "sheets_inspect_macros"
            || call.name === "sheets_set_macros"
          );
        });
        const failedIndex = macroIndex >= 0 ? macroIndex : 0;
        const failedCall = toolCalls[failedIndex] || {};
        const existingDetails = error
          && error.details
          && typeof error.details === "object"
          ? error.details
          : {};
        error.code = error.code || "EXECUTION_FAILED";
        error.details = {
          ...existingDetails,
          phase: existingDetails.phase || "sheets-command",
          tool: failedCall.name ? String(failedCall.name) : undefined,
          toolCallIndex: failedIndex,
          completedToolCalls: 0,
          partialMutationPossible: Boolean(existingDetails.partialMutationPossible),
        };
        throw error;
      });
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

            var CHART_TYPE_ALIASES = {
              line: "lineNormal",
              lineMarker: "lineNormalMarker",
              stackedBar: "barStacked",
              stackedBarPercent: "barStackedPercent",
              stackedLine: "lineStacked",
              stackedLinePercent: "lineStackedPercent",
              column: "bar",
            };

            var CHART_TYPE_CANONICALS = [
              "bar", "barStacked", "barStackedPercent", "bar3D",
              "barStacked3D", "barStackedPercent3D", "barStackedPercent3DPerspective",
              "horizontalBar", "horizontalBarStacked", "horizontalBarStackedPercent",
              "horizontalBar3D", "horizontalBarStacked3D", "horizontalBarStackedPercent3D",
              "lineNormal", "lineStacked", "lineStackedPercent", "lineNormalMarker",
              "lineStackedMarker", "lineStackedPerMarker", "line3D",
              "pie", "pie3D", "doughnut",
              "scatter", "scatterLine", "scatterLineMarker", "scatterSmooth", "scatterSmoothMarker",
              "stock", "area", "areaStacked", "areaStackedPercent",
              "comboCustom", "comboBarLine", "comboBarLineSecondary",
              "radar", "radarMarker", "radarFilled",
            ];

            var CHART_TYPE_INPUTS = CHART_TYPE_CANONICALS.concat(
              Object.keys(CHART_TYPE_ALIASES)
            );

            function chartTypeSuggestions(value) {
              var token = String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
              if (!token) return [];
              var candidates = CHART_TYPE_INPUTS.map(function (candidate, index) {
                var candidateToken = candidate.toLowerCase().replace(/[^a-z0-9]/g, "");
                if (
                  !candidateToken
                  || (
                    token.indexOf(candidateToken) !== 0
                    && candidateToken.indexOf(token) !== 0
                  )
                ) {
                  return null;
                }
                return {
                  candidate: candidate,
                  direction: token.indexOf(candidateToken) === 0 ? 0 : 1,
                  distance: Math.abs(token.length - candidateToken.length),
                  index: index,
                };
              }).filter(Boolean);
              candidates.sort(function (left, right) {
                return left.direction - right.direction
                  || left.distance - right.distance
                  || left.index - right.index;
              });
              var suggestions = candidates.slice(0, 3).map(function (item) {
                return item.candidate;
              });
              for (var suggestionIndex = 0; suggestionIndex < suggestions.length; suggestionIndex += 1) {
                var canonical = CHART_TYPE_ALIASES[suggestions[suggestionIndex]];
                if (canonical && suggestions.indexOf(canonical) < 0) suggestions.push(canonical);
              }
              return suggestions.slice(0, 4);
            }

            function resolveChartType(value, fallback, field) {
              var received = value === undefined || value === null || value === ""
                ? String(fallback || "bar")
                : String(value);
              var normalized = CHART_TYPE_ALIASES[received] || received;
              if (CHART_TYPE_CANONICALS.indexOf(normalized) < 0) {
                throw sheetError(
                  "INVALID_TOOL_ARGUMENTS",
                  "不支持的图表类型：" + received,
                  {
                    phase: "sheets-chart-validation",
                    field: field || "type",
                    receivedType: received,
                    supportedTypes: CHART_TYPE_INPUTS.slice(),
                    aliases: Object.assign({}, CHART_TYPE_ALIASES),
                    suggestedTypes: chartTypeSuggestions(received),
                    partialMutationPossible: false,
                  }
                );
              }
              return {
                received: received,
                normalized: normalized,
              };
            }

            function normalizeChartType(value, fallback, field) {
              return resolveChartType(value, fallback, field).normalized;
            }

            function chartCreationError(error, context) {
              var causeMessage = error && error.message
                ? String(error.message)
                : String(error || "ONLYOFFICE 未返回图表对象");
              return sheetError(
                "SHEETS_CHART_CREATE_FAILED",
                "ONLYOFFICE 创建图表失败：" + causeMessage,
                {
                  phase: "sheets-chart-create",
                  apiMethod: "ApiWorksheet.AddChart",
                  sheet: context.sheet,
                  range: context.range,
                  receivedType: context.receivedType,
                  normalizedType: context.normalizedType,
                  causeMessage: causeMessage,
                  partialMutationPossible: true,
                }
              );
            }

            function enumKey(value) {
              return String(value).toLowerCase().replace(/[^a-z0-9]/g, "");
            }

            function normalizeEnum(value, fallback, aliases, canonicalValues, label) {
              var candidate = value === undefined || value === null || value === "" ? fallback : value;
              var lookup = {};
              var canonical = Array.isArray(canonicalValues) ? canonicalValues : [];
              for (var canonicalIndex = 0; canonicalIndex < canonical.length; canonicalIndex += 1) {
                lookup[enumKey(canonical[canonicalIndex])] = canonical[canonicalIndex];
              }
              var aliasNames = Object.keys(aliases || {});
              for (var aliasIndex = 0; aliasIndex < aliasNames.length; aliasIndex += 1) {
                lookup[enumKey(aliasNames[aliasIndex])] = aliases[aliasNames[aliasIndex]];
              }
              var normalized = lookup[enumKey(candidate)];
              if (!normalized) {
                throw new Error(
                  (label || "枚举值")
                  + " 不支持："
                  + candidate
                  + "；可用值："
                  + canonical.concat(aliasNames).join("、")
                );
              }
              return normalized;
            }

            function normalizeComparisonOperator(value, fallback, label) {
              return normalizeEnum(value, fallback || "xlGreater", {
                between: "xlBetween",
                notBetween: "xlNotBetween",
                equal: "xlEqual",
                notEqual: "xlNotEqual",
                greaterThan: "xlGreater",
                lessThan: "xlLess",
                greaterThanOrEqual: "xlGreaterEqual",
                lessThanOrEqual: "xlLessEqual",
              }, [
                "xlBetween",
                "xlNotBetween",
                "xlEqual",
                "xlNotEqual",
                "xlGreater",
                "xlLess",
                "xlGreaterEqual",
                "xlLessEqual",
              ], label || "比较运算符");
            }

            function normalizeFilterOperator(value) {
              if (value === undefined || value === null || value === "") return undefined;
              return normalizeEnum(value, undefined, {
                and: "xlAnd",
                or: "xlOr",
                filterValues: "xlFilterValues",
                values: "xlFilterValues",
                top10Items: "xlTop10Items",
                bottom10Items: "xlBottom10Items",
                top10Percent: "xlTop10Percent",
                bottom10Percent: "xlBottom10Percent",
                filterCellColor: "xlFilterCellColor",
                filterFontColor: "xlFilterFontColor",
                filterIcon: "xlFilterIcon",
                dynamic: "xlFilterDynamic",
              }, [
                "xlAnd",
                "xlOr",
                "xlFilterValues",
                "xlTop10Items",
                "xlBottom10Items",
                "xlTop10Percent",
                "xlBottom10Percent",
                "xlFilterCellColor",
                "xlFilterFontColor",
                "xlFilterIcon",
                "xlFilterDynamic",
              ], "筛选运算符");
            }

            function normalizeValidationType(value) {
              return normalizeEnum(value, "xlValidateList", {
                inputOnly: "xlValidateInputOnly",
                wholeNumber: "xlValidateWholeNumber",
                decimal: "xlValidateDecimal",
                list: "xlValidateList",
                date: "xlValidateDate",
                time: "xlValidateTime",
                textLength: "xlValidateTextLength",
                custom: "xlValidateCustom",
              }, [
                "xlValidateInputOnly",
                "xlValidateWholeNumber",
                "xlValidateDecimal",
                "xlValidateList",
                "xlValidateDate",
                "xlValidateTime",
                "xlValidateTextLength",
                "xlValidateCustom",
              ], "数据验证类型");
            }

            function normalizeValidationAlertStyle(value) {
              return normalizeEnum(value, "xlValidAlertStop", {
                stop: "xlValidAlertStop",
                warning: "xlValidAlertWarning",
                information: "xlValidAlertInformation",
                info: "xlValidAlertInformation",
              }, [
                "xlValidAlertStop",
                "xlValidAlertWarning",
                "xlValidAlertInformation",
              ], "数据验证警告样式");
            }

            function normalizeSortHeader(value) {
              if (value === true) return "xlYes";
              if (value === false) return "xlNo";
              return normalizeEnum(value, "xlGuess", {
                yes: "xlYes",
                no: "xlNo",
                guess: "xlGuess",
              }, ["xlYes", "xlNo", "xlGuess"], "排序表头");
            }

            function normalizeSortOrientation(value) {
              return normalizeEnum(value, "xlSortColumns", {
                rows: "xlSortRows",
                columns: "xlSortColumns",
              }, ["xlSortRows", "xlSortColumns"], "排序方向");
            }

            function normalizePageOrientation(value) {
              return normalizeEnum(value, "xlPortrait", {
                portrait: "xlPortrait",
                landscape: "xlLandscape",
              }, ["xlPortrait", "xlLandscape"], "页面方向");
            }

            function normalizeFormulaInput(value) {
              if (Array.isArray(value)) return value.map(normalizeFormulaInput);
              var formulaText = String(value);
              return formulaText.charAt(0) === "=" ? formulaText : "=" + formulaText;
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

            function mutationCall(value, method, feature) {
              var result = requireMethod(value, method, feature).apply(
                value,
                Array.prototype.slice.call(arguments, 3)
              );
              if (result === false) {
                throw new Error("ONLYOFFICE 拒绝" + (feature || method));
              }
              return result;
            }

            function sheetError(code, message, details) {
              var error = new Error(message);
              error.code = code || "EXECUTION_FAILED";
              error.details = details || {};
              return error;
            }

            function objectMutationCall(value, method, feature) {
              var result = mutationCall.apply(null, arguments);
              if (result === null || result === undefined) {
                throw sheetError(
                  "EXECUTION_FAILED",
                  "ONLYOFFICE 未创建" + (feature || method)
                );
              }
              return result;
            }

            function isSheetCellValue(value) {
              return value === null
                || typeof value === "string"
                || typeof value === "boolean"
                || (typeof value === "number" && isFinite(value));
            }

            function isFormulaValue(value) {
              return typeof value === "string";
            }

            function describeShape(shape) {
              return shape.rows + "x" + shape.columns;
            }

            function targetRangeShape(range, label) {
              var rows = safeCall(range, "GetRowsCount");
              var columns = safeCall(range, "GetColumnsCount");
              if (
                typeof rows !== "number"
                || typeof columns !== "number"
                || !isFinite(rows)
                || !isFinite(columns)
                || rows < 1
                || columns < 1
              ) {
                throw sheetError(
                  "SHEETS_API_UNSUPPORTED",
                  "当前 ONLYOFFICE 版本无法读取" + label + "目标区域尺寸",
                  { partialMutationPossible: false }
                );
              }
              return {
                rows: Math.floor(rows),
                columns: Math.floor(columns),
              };
            }

            function inputShape(value, label, cellValidator) {
              if (!Array.isArray(value)) {
                if (!cellValidator(value)) {
                  throw sheetError(
                    "INVALID_TOOL_ARGUMENTS",
                    label + " 标量类型不受支持",
                    { partialMutationPossible: false }
                  );
                }
                return { rows: 1, columns: 1, scalar: true };
              }
              if (!value.length) {
                throw sheetError(
                  "INVALID_TOOL_ARGUMENTS",
                  label + " 矩阵不能为空",
                  { partialMutationPossible: false }
                );
              }
              var columns = null;
              for (var rowIndex = 0; rowIndex < value.length; rowIndex += 1) {
                var row = value[rowIndex];
                if (!Array.isArray(row) || !row.length) {
                  throw sheetError(
                    "INVALID_TOOL_ARGUMENTS",
                    label + " 必须是非空二维矩阵",
                    { partialMutationPossible: false }
                  );
                }
                if (columns === null) columns = row.length;
                if (row.length !== columns) {
                  throw sheetError(
                    "INVALID_TOOL_ARGUMENTS",
                    label + " 矩阵每一行的列数必须一致",
                    { partialMutationPossible: false }
                  );
                }
                for (var columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
                  if (!cellValidator(row[columnIndex])) {
                    throw sheetError(
                      "INVALID_TOOL_ARGUMENTS",
                      label + " 矩阵包含不受支持的单元格类型",
                      { partialMutationPossible: false }
                    );
                  }
                }
              }
              return {
                rows: value.length,
                columns: columns,
                scalar: false,
              };
            }

            function validateSizedInput(range, value, label, cellValidator) {
              var target = targetRangeShape(range, label);
              var input = inputShape(value, label, cellValidator);
              if (input.rows !== target.rows || input.columns !== target.columns) {
                throw sheetError(
                  "INVALID_TOOL_ARGUMENTS",
                  label
                    + " 尺寸 "
                    + describeShape(input)
                    + " 与目标区域 "
                    + describeShape(target)
                    + " 不一致",
                  { partialMutationPossible: false }
                );
              }
              return {
                input: { rows: input.rows, columns: input.columns },
                target: target,
              };
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
              if (typeof Api.CreateColorFromRGB === "function") return Api.CreateColorFromRGB(red, green, blue);
              if (typeof Api.Color === "function") return Api.Color(red, green, blue);
              if (typeof Api.RGB === "function") return Api.RGB(red, green, blue);
              if (typeof Api.CreateRGBColor === "function") return Api.CreateRGBColor(red, green, blue);
              throw new Error("当前 ONLYOFFICE 版本无法创建单元格颜色");
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

            function describeColor(value) {
              if (value === null || value === undefined) return null;
              var json = serialized(value);
              if (json !== null) return json;
              var rgb = safeCall(value, "GetRGB");
              if (rgb !== null && rgb !== undefined) {
                var rgbJson = serialized(rgb);
                return rgbJson !== null ? rgbJson : rgb;
              }
              if (
                typeof value.r === "number"
                && typeof value.g === "number"
                && typeof value.b === "number"
              ) {
                return { r: value.r, g: value.g, b: value.b };
              }
              return String(value);
            }

            function describeBorders(range) {
              var borders = safeCall(range, "GetBorders");
              if (borders === null || borders === undefined) return null;
              var json = serialized(borders);
              return json !== null ? json : borders;
            }

            function colorChannel(value, names) {
              for (var nameIndex = 0; nameIndex < names.length; nameIndex += 1) {
                var channel = value && value[names[nameIndex]];
                if (typeof channel === "number" && isFinite(channel)) {
                  return Math.max(0, Math.min(255, Math.round(channel)));
                }
              }
              return null;
            }

            function hexChannel(value) {
              var text = Math.max(0, Math.min(255, Number(value) || 0)).toString(16).toUpperCase();
              return text.length < 2 ? "0" + text : text;
            }

            function normalizeVerificationColor(value) {
              if (typeof value === "string") {
                var text = value.trim();
                var shortHex = /^#([0-9a-f]{3})$/i.exec(text);
                if (shortHex) {
                  return "#" + shortHex[1].split("").map(function (digit) {
                    return digit + digit;
                  }).join("").toUpperCase();
                }
                var fullHex = /^#([0-9a-f]{6})$/i.exec(text);
                if (fullHex) return "#" + fullHex[1].toUpperCase();
                var rgbText = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(text);
                if (rgbText) {
                  return "#" + hexChannel(rgbText[1]) + hexChannel(rgbText[2]) + hexChannel(rgbText[3]);
                }
                return null;
              }
              if (!value || typeof value !== "object") return null;
              var red = colorChannel(value, ["r", "R", "red", "Red"]);
              var green = colorChannel(value, ["g", "G", "green", "Green"]);
              var blue = colorChannel(value, ["b", "B", "blue", "Blue"]);
              if (red !== null && green !== null && blue !== null) {
                return "#" + hexChannel(red) + hexChannel(green) + hexChannel(blue);
              }
              if (value.color !== undefined) return normalizeVerificationColor(value.color);
              return null;
            }

            function normalizeVerificationBorders(value) {
              if (!Array.isArray(value)) return null;
              var normalized = [];
              for (var borderIndex = 0; borderIndex < value.length; borderIndex += 1) {
                var border = value[borderIndex];
                if (!border || typeof border !== "object") return null;
                var borderColor = normalizeVerificationColor(border.color);
                if (!border.side || !border.style || !borderColor) return null;
                normalized.push({
                  side: String(border.side).toLowerCase(),
                  style: String(border.style).toLowerCase(),
                  color: borderColor,
                });
              }
              return normalized;
            }

            function verificationComparable(field, value) {
              if (value === null || value === undefined) {
                return { available: false, reason: "readback-null" };
              }
              if (field === "fontColor" || field === "fillColor") {
                var normalizedColor = normalizeVerificationColor(value);
                return normalizedColor
                  ? { available: true, value: normalizedColor }
                  : { available: false, reason: "unsupported-color-readback" };
              }
              if (field === "borders") {
                var normalizedBorders = normalizeVerificationBorders(value);
                return normalizedBorders
                  ? { available: true, value: normalizedBorders }
                  : { available: false, reason: "unsupported-borders-readback" };
              }
              if (
                field === "horizontalAlign"
                || field === "verticalAlign"
                || field === "type"
                || field === "operator"
              ) {
                return {
                  available: true,
                  value: String(value).toLowerCase().replace(/[^a-z0-9]/g, ""),
                };
              }
              if (typeof value === "number" && isFinite(value)) {
                return { available: true, value: Math.round(value * 1000000) / 1000000 };
              }
              if (
                typeof value === "string"
                || typeof value === "boolean"
              ) {
                return { available: true, value: value };
              }
              var json = serialized(value);
              return json === null
                ? { available: false, reason: "unsupported-readback-shape" }
                : { available: true, value: json };
            }

            function verificationValuesEqual(left, right) {
              if (left === right) return true;
              return JSON.stringify(left) === JSON.stringify(right);
            }

            function verifyReadback(requested, observed) {
              var fields = {};
              var counts = { verified: 0, failed: 0, unavailable: 0 };
              Object.keys(requested || {}).forEach(function (field) {
                var expected = verificationComparable(field, requested[field]);
                var actual = verificationComparable(
                  field,
                  observed && hasOwn(observed, field) ? observed[field] : null
                );
                var status = "verified";
                var reason = null;
                if (!actual.available) {
                  status = "unavailable";
                  reason = actual.reason;
                } else if (!expected.available || !verificationValuesEqual(expected.value, actual.value)) {
                  status = "failed";
                  reason = expected.available ? "readback-mismatch" : expected.reason;
                }
                counts[status] += 1;
                fields[field] = {
                  requested: requested[field],
                  observed: observed && hasOwn(observed, field) ? observed[field] : null,
                  status: status,
                  reason: reason || undefined,
                };
              });
              return {
                status: counts.failed > 0
                  ? "failed"
                  : (counts.unavailable > 0 ? "unavailable" : "verified"),
                counts: counts,
                fields: fields,
              };
            }

            function requestedFormat(args, columnWidthChars, rowHeightPt) {
              var requested = {};
              [
                "fontSize", "fontName", "bold", "italic", "underline", "strikeout",
                "fontColor", "fillColor", "horizontalAlign", "verticalAlign",
                "numberFormat", "wrap", "orientation",
              ].forEach(function (field) {
                if (hasOwn(args, field)) requested[field] = args[field];
              });
              if (columnWidthChars !== undefined) requested.columnWidthChars = columnWidthChars;
              if (rowHeightPt !== undefined) requested.rowHeightPt = rowHeightPt;
              if (Array.isArray(args.borders) && args.borders.length > 0) {
                requested.borders = args.borders;
              }
              return requested;
            }

            function describeFormat(range) {
              var font = safeCall(range, "GetFont");
              return {
                fontName: safeCall(range, "GetFontName") || safeCall(font, "GetName"),
                fontSize: safeCall(range, "GetFontSize") || safeCall(font, "GetSize"),
                bold: safeCall(range, "GetBold"),
                italic: safeCall(range, "GetItalic"),
                underline: safeCall(range, "GetUnderline"),
                strikeout: safeCall(range, "GetStrikeout"),
                fontColor: describeColor(safeCall(range, "GetFontColor") || safeCall(font, "GetColor")),
                fillColor: describeColor(safeCall(range, "GetFillColor")),
                horizontalAlign: safeCall(range, "GetAlignHorizontal"),
                verticalAlign: safeCall(range, "GetAlignVertical"),
                numberFormat: safeCall(range, "GetNumberFormat"),
                wrap: safeCall(range, "GetWrapText"),
                orientation: safeCall(range, "GetOrientation"),
                rowHeightPt: safeCall(range, "GetRowHeight"),
                columnWidthChars: safeCall(range, "GetColumnWidth"),
                borders: describeBorders(range),
              };
            }

            function describeCondition(condition, index) {
              if (!condition) return null;
              var conditionFont = safeCall(condition, "GetFont");
              return {
                index: index,
                type: safeCall(condition, "GetType"),
                dupeUnique: safeCall(condition, "GetDupeUnique"),
                operator: safeCall(condition, "GetOperator"),
                formula1: safeCall(condition, "GetFormula1"),
                formula2: safeCall(condition, "GetFormula2"),
                fillColor: describeColor(safeCall(condition, "GetFillColor")),
                fontColor: describeColor(safeCall(conditionFont, "GetColor")),
                bold: safeCall(conditionFont, "GetBold"),
              };
            }

            function conditionTypeForVerification(value) {
              var token = String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
              var types = {
                xlcellvalue: "cellValue",
                cellvalue: "cellValue",
                xlexpression: "expression",
                expression: "expression",
                unique: "uniqueOrDuplicateValues",
                uniquevalues: "uniqueOrDuplicateValues",
                xluniquevalues: "uniqueOrDuplicateValues",
                duplicatevalues: "uniqueOrDuplicateValues",
                colorscale: "colorScale",
                xlcolorscale: "colorScale",
                databar: "dataBar",
                xldatabar: "dataBar",
                iconset: "iconSet",
                xliconset: "iconSet",
                top10: "top10",
                xltop10: "top10",
                aboveaverage: "aboveAverage",
                xlaboveaverage: "aboveAverage",
              };
              return types[token] || (value === null || value === undefined ? null : String(value));
            }

            function conditionRuleExpectation(args, conditionType, expectedCount) {
              var expected = {};
              if (typeof expectedCount === "number") expected.count = expectedCount;
              expected.type = conditionTypeForVerification(
                conditionType === "uniqueValues" || conditionType === "duplicateValues"
                  ? "unique"
                  : conditionType
              );
              if (conditionType === "cellValue" || conditionType === "expression") {
                expected.operator = normalizeComparisonOperator(
                  args.operator,
                  "xlGreater",
                  "条件格式运算符"
                );
                if (hasOwn(args, "formula1")) expected.formula1 = args.formula1;
                if (hasOwn(args, "formula2")) expected.formula2 = args.formula2;
              }
              if (conditionType === "uniqueValues") expected.dupeUnique = "xlUnique";
              if (conditionType === "duplicateValues") expected.dupeUnique = "xlDuplicate";
              if (hasOwn(args, "fillColor")) expected.fillColor = args.fillColor;
              if (hasOwn(args, "fontColor")) expected.fontColor = args.fontColor;
              if (hasOwn(args, "bold")) expected.bold = Boolean(args.bold);
              return expected;
            }

            function conditionRuleObservation(rule, count) {
              var observed = {
                count: count,
              };
              if (!rule) return observed;
              observed.type = conditionTypeForVerification(rule.type);
              [
                "dupeUnique", "operator", "formula1", "formula2",
                "fillColor", "fontColor", "bold",
              ].forEach(function (field) {
                observed[field] = rule[field];
              });
              return observed;
            }

            function describeConditions(range) {
              var conditions = safeCall(range, "GetFormatConditions");
              if (!conditions) return { count: 0, rules: [] };
              var count = safeCall(conditions, "GetCount");
              var rules = [];
              if (typeof conditions.GetItem === "function" && typeof count === "number") {
                for (var conditionIndex = 0; conditionIndex < count; conditionIndex += 1) {
                  rules.push(describeCondition(safeCall(conditions, "GetItem", conditionIndex + 1), conditionIndex));
                }
              } else {
                var conditionItems = arrayValue(conditions);
                for (var itemIndex = 0; itemIndex < conditionItems.length; itemIndex += 1) {
                  rules.push(describeCondition(conditionItems[itemIndex], itemIndex));
                }
              }
              return { count: typeof count === "number" ? count : rules.length, rules: rules };
            }

            function describeValidation(validation) {
              if (!validation) return null;
              var rule = {
                type: safeCall(validation, "GetType"),
                alertStyle: safeCall(validation, "GetAlertStyle"),
                operator: safeCall(validation, "GetOperator"),
                formula1: safeCall(validation, "GetFormula1"),
                formula2: safeCall(validation, "GetFormula2"),
                ignoreBlank: safeCall(validation, "GetIgnoreBlank"),
                inCellDropdown: safeCall(validation, "GetInCellDropdown"),
                showInput: safeCall(validation, "GetShowInput"),
                showError: safeCall(validation, "GetShowError"),
                inputTitle: safeCall(validation, "GetInputTitle"),
                inputMessage: safeCall(validation, "GetInputMessage"),
                errorTitle: safeCall(validation, "GetErrorTitle"),
                errorMessage: safeCall(validation, "GetErrorMessage"),
              };
              var identifyingValues = [
                rule.type,
                rule.formula1,
                rule.formula2,
                rule.inputTitle,
                rule.inputMessage,
                rule.errorTitle,
                rule.errorMessage,
              ];
              var hasRule = identifyingValues.some(function (value) {
                return value !== null && value !== undefined && value !== "";
              });
              return hasRule ? rule : null;
            }

            function describeRange(range, includeValues, includeFormat, includeConditionalFormats, includeValidation) {
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
                format: includeFormat ? describeFormat(range) : undefined,
                conditionalFormats: includeConditionalFormats ? describeConditions(range) : undefined,
                validation: includeValidation
                  ? describeValidation(safeCall(range, "GetValidation"))
                  : undefined,
              };
            }

            function sheetViewSettings(sheet, partialMutationPossible) {
              var model = sheet && sheet.worksheet;
              var settings = model && typeof model.getSheetViewSettings === "function"
                ? model.getSheetViewSettings(true)
                : null;
              if (!settings) {
                throw sheetError(
                  "SHEETS_API_UNSUPPORTED",
                  "当前 ONLYOFFICE 版本无法读取工作表屏幕视图",
                  { partialMutationPossible: Boolean(partialMutationPossible) }
                );
              }
              return settings;
            }

            function describeDisplayState(sheet, partialMutationPossible) {
              var settings = sheetViewSettings(sheet, partialMutationPossible);
              return {
                displayGridlines: settings.showGridLines !== false,
                displayHeadings: settings.showRowColHeaders !== false,
              };
            }

            function freezeColumnNumber(column) {
              var number = 0;
              var normalized = String(column || "").toUpperCase();
              for (var columnIndex = 0; columnIndex < normalized.length; columnIndex += 1) {
                number = number * 26 + normalized.charCodeAt(columnIndex) - 64;
              }
              return number;
            }

            function freezeColumnName(number) {
              var remaining = Math.max(1, Number(number) || 1);
              var result = "";
              while (remaining > 0) {
                var remainder = (remaining - 1) % 26;
                result = String.fromCharCode(65 + remainder) + result;
                remaining = Math.floor((remaining - 1) / 26);
              }
              return result;
            }

            function freezeDimensions(address, rows, columns) {
              var maxRows = 1048576;
              var maxColumns = 16384;
              var localAddress = String(address || "")
                .split("!")
                .pop()
                .replace(/\$/g, "")
                .trim();
              var matched = /^([1-9][0-9]*):([1-9][0-9]*)$/.exec(localAddress);
              var frozenRows;
              var frozenColumns;
              if (matched) {
                frozenRows = Number(matched[2]);
                frozenColumns = 0;
              } else {
                matched = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(localAddress);
                if (matched) {
                  frozenRows = 0;
                  frozenColumns = freezeColumnNumber(matched[2]);
                } else {
                  var parts = localAddress.split(":");
                  matched = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(parts[parts.length - 1]);
                  if (matched) {
                    var lastColumn = freezeColumnNumber(matched[1]);
                    var lastRow = Number(matched[2]);
                    frozenRows = lastColumn >= maxColumns ? lastRow : (
                      lastRow >= maxRows ? 0 : lastRow
                    );
                    frozenColumns = lastRow >= maxRows ? lastColumn : (
                      lastColumn >= maxColumns ? 0 : lastColumn
                    );
                  }
                }
              }
              if (frozenRows === undefined || frozenColumns === undefined) {
                var rowCount = Number(rows);
                var columnCount = Number(columns);
                var validRows = isFinite(rowCount) && rowCount >= 1 && Math.floor(rowCount) === rowCount;
                var validColumns = isFinite(columnCount) && columnCount >= 1 && Math.floor(columnCount) === columnCount;
                if (validRows && validColumns) {
                  if (columnCount >= maxColumns && rowCount < maxRows) {
                    frozenRows = rowCount;
                    frozenColumns = 0;
                  } else if (rowCount >= maxRows && columnCount < maxColumns) {
                    frozenRows = 0;
                    frozenColumns = columnCount;
                  } else if (rowCount < maxRows && columnCount < maxColumns) {
                    frozenRows = rowCount;
                    frozenColumns = columnCount;
                  }
                }
              }
              if (
                frozenRows === undefined
                || frozenColumns === undefined
                || frozenRows < 0
                || frozenColumns < 0
                || frozenRows >= maxRows
                || frozenColumns >= maxColumns
              ) {
                return null;
              }
              return {
                frozenRows: frozenRows,
                frozenColumns: frozenColumns,
              };
            }

            function describeFreezeState(freezePanes) {
              var frozenRange = safeCall(freezePanes, "GetLocation");
              if (!frozenRange) {
                return {
                  location: null,
                  frozenRows: 0,
                  frozenColumns: 0,
                  topLeftCell: null,
                };
              }
              var location = {
                address: safeCall(frozenRange, "GetAddress", true, true, "xlA1", false),
                rows: safeCall(frozenRange, "GetRowsCount"),
                columns: safeCall(frozenRange, "GetColumnsCount"),
              };
              var dimensions = freezeDimensions(
                location.address,
                location.rows,
                location.columns
              );
              if (!dimensions) {
                throw sheetError(
                  "SHEETS_API_UNSUPPORTED",
                  "无法解析 ONLYOFFICE 返回的冻结区域",
                  { location: location, partialMutationPossible: false }
                );
              }
              return {
                location: location,
                frozenRows: dimensions.frozenRows,
                frozenColumns: dimensions.frozenColumns,
                topLeftCell: freezeColumnName(dimensions.frozenColumns + 1)
                  + String(dimensions.frozenRows + 1),
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

            function structuredTableCreateFields(args) {
              return [
                "sourceType",
                "name",
                "style",
                "showTotals",
                "showHeaders",
                "rowStripes",
                "columnStripes",
                "firstColumn",
                "lastColumn",
                "showAutoFilter",
                "showAutoFilterDropDown",
                "summary",
                "alternativeText",
              ].filter(function (name) {
                return hasOwn(args, name);
              });
            }

            function unexpectedArgumentFields(args, allowedFields) {
              return Object.keys(args).filter(function (name) {
                return allowedFields.indexOf(name) === -1;
              });
            }

            function isStructuredTableObject(value) {
              return Boolean(
                value
                && typeof value === "object"
                && typeof value.GetRange === "function"
              );
            }

            function rangeStyleHeaderReference(reference) {
              var localReference = String(reference || "")
                .split("!")
                .pop()
                .replace(/\$/g, "");
              var matched = /^([A-Za-z]{1,3})([1-9][0-9]*):([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(localReference);
              if (!matched) {
                var single = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(localReference);
                return single ? single[1] + single[2] : null;
              }
              return matched[1] + matched[2] + ":" + matched[3] + matched[2];
            }

            function formatRangeStyleTable(sheet, reference) {
              var tableRange = getRange(sheet, String(reference));
              if (!tableRange) throw new Error("无效表格区域：" + reference);
              var warnings = [];
              var borderColor = color("#D9E2F3");
              [
                "xlEdgeTop",
                "xlEdgeBottom",
                "xlEdgeLeft",
                "xlEdgeRight",
                "xlInsideHorizontal",
                "xlInsideVertical",
              ].forEach(function (side) {
                mutationCall(tableRange, "SetBorders", "设置普通区域表格边框", side, "xlContinuous", borderColor);
              });
              mutationCall(tableRange, "SetAlignVertical", "设置普通区域表格垂直对齐", "center");

              var headerReference = rangeStyleHeaderReference(reference);
              if (headerReference) {
                var headerRange = getRange(sheet, headerReference);
                mutationCall(headerRange, "SetBold", "设置普通区域表头粗体", true);
                mutationCall(headerRange, "SetFontColor", "设置普通区域表头字体颜色", color("#FFFFFF"));
                mutationCall(headerRange, "SetFillColor", "设置普通区域表头填充", color("#4472C4"));
                mutationCall(headerRange, "SetAlignHorizontal", "设置普通区域表头对齐", "center");
              } else {
                warnings.push("header_range_unavailable");
              }

              var filterApplied = false;
              if (typeof tableRange.SetAutoFilter === "function") {
                try {
                  tableRange.SetAutoFilter();
                  filterApplied = true;
                } catch (filterError) {
                  warnings.push("auto_filter_unavailable");
                }
              } else {
                warnings.push("auto_filter_unavailable");
              }
              return {
                formatted: true,
                filterApplied: filterApplied,
                warnings: warnings,
              };
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

            function reinforceChartTitleText(chart, horizontal, fontSize, bold) {
              var chartModel = chart && chart.Chart && chart.Chart.chart;
              if (!chartModel || typeof AscCommonWord === "undefined" || typeof AscCommonWord.ParaTextPr !== "function") {
                return false;
              }
              var title = null;
              if (horizontal === null) {
                title = chartModel.title;
              } else if (chartModel.plotArea) {
                var axisGetter = horizontal ? "getHorizontalAxis" : "getVerticalAxis";
                var axisModel = typeof chartModel.plotArea[axisGetter] === "function"
                  ? chartModel.plotArea[axisGetter]()
                  : null;
                title = axisModel && axisModel.title;
              }
              var content = title && typeof title.getDocContent === "function" ? title.getDocContent() : null;
              if (!content || typeof content.SetApplyToAll !== "function" || typeof content.AddToParagraph !== "function") {
                return false;
              }
              content.SetApplyToAll(true);
              try {
                content.AddToParagraph(new AscCommonWord.ParaTextPr({
                  FontSize: asFinite(fontSize, horizontal === null ? 13 : 11),
                  Bold: Boolean(bold),
                }), false);
              } finally {
                content.SetApplyToAll(false);
              }
              return true;
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
                var axisTitleFontSize = asFinite(axis.titleFontSize, 11);
                chart[titleMethod](String(axis.title), axisTitleFontSize, Boolean(axis.titleBold));
                reinforceChartTitleText(chart, horizontal, axisTitleFontSize, axis.titleBold);
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
                  typedSeries.ChangeChartType(normalizeChartType(
                    update.type,
                    "bar",
                    "seriesUpdates[" + updateIndex + "].type"
                  ));
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
                      chart.SetDataPointNumFormat(String(point.numberFormat), seriesIndex, targetPoint, Boolean(point.allSeries));
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
                var chartTitleFontSize = asFinite(args.titleFontSize, 13);
                chart.SetTitle(String(args.title), chartTitleFontSize, Boolean(args.titleBold));
                reinforceChartTitleText(chart, null, chartTitleFontSize, args.titleBold);
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

            function unsupportedFeature(feature, message, details) {
              throw sheetError(
                "SHEETS_API_UNSUPPORTED",
                message,
                Object.assign({
                  feature: feature,
                  retryable: false,
                  mutationState: "none",
                  partialMutationPossible: false,
                }, details || {})
              );
            }

            function invalidFilterArguments(message, field, details) {
              throw sheetError(
                "INVALID_TOOL_ARGUMENTS",
                message,
                Object.assign({
                  field: field || null,
                  mutationState: "none",
                  partialMutationPossible: false,
                }, details || {})
              );
            }

            function sheetAutoFilter(sheet) {
              return typeof sheet.GetAutoFilter === "function"
                ? sheet.GetAutoFilter()
                : sheet.AutoFilter;
            }

            function filterRangeIdentity(range) {
              var address = safeCall(range, "GetAddress", true, true, "xlA1", false);
              if (typeof address !== "string" || !address) address = safeCall(range, "GetAddress");
              if (typeof address !== "string" || !address) return null;
              return address
                .split("!")
                .pop()
                .replace(/\$/g, "")
                .replace(/^'|'$/g, "")
                .toUpperCase();
            }

            function resolveFilterTargetRange(sheet, args, autoFilter) {
              if (hasOwn(args, "range")) {
                if (typeof args.range !== "string" || !args.range.trim()) {
                  invalidFilterArguments("sheets_filter.range 必须是有效区域", "range");
                }
                try {
                  var requestedRange = getRange(sheet, args.range);
                  if (requestedRange) return requestedRange;
                } catch (error) {
                  invalidFilterArguments("sheets_filter.range 无法解析", "range");
                }
                invalidFilterArguments("sheets_filter.range 无法解析", "range");
              }
              var currentRange = safeCall(autoFilter, "GetRange");
              if (currentRange) return currentRange;
              var usedRange = safeCall(sheet, "GetUsedRange");
              if (usedRange) return usedRange;
              invalidFilterArguments(
                "未提供 sheets_filter.range，且当前工作表没有可筛选区域",
                "range"
              );
            }

            function isSafeFilterScalar(value) {
              return typeof value === "string"
                || typeof value === "boolean"
                || (typeof value === "number" && isFinite(value));
            }

            function isSafeFilterCriterion(value, allowArray) {
              if (isSafeFilterScalar(value)) return true;
              if (!allowArray || !Array.isArray(value) || !value.length) return false;
              for (var criterionIndex = 0; criterionIndex < value.length; criterionIndex += 1) {
                if (!isSafeFilterScalar(value[criterionIndex])) return false;
              }
              return true;
            }

            function validateFilterArguments(args) {
              var action = String(args.action || "");
              if (action !== "set" && action !== "showAll" && action !== "reapply") {
                invalidFilterArguments("不支持的筛选操作：" + action, "action");
              }
              var setOnlyFields = [
                "range",
                "field",
                "criteria1",
                "operator",
                "criteria2",
                "visibleDropDown",
              ];
              if (action !== "set") {
                for (var setFieldIndex = 0; setFieldIndex < setOnlyFields.length; setFieldIndex += 1) {
                  var setOnlyField = setOnlyFields[setFieldIndex];
                  if (hasOwn(args, setOnlyField)) {
                    invalidFilterArguments(
                      "sheets_filter." + setOnlyField + " 仅允许用于 action=set",
                      setOnlyField
                    );
                  }
                }
                return { action: action, operator: undefined };
              }

              if (!hasOwn(args, "field") || !Number.isInteger(args.field) || args.field < 1) {
                invalidFilterArguments(
                  "sheets_filter.action=set 需要从 1 开始的整数 field",
                  "field"
                );
              }
              if (hasOwn(args, "visibleDropDown") && typeof args.visibleDropDown !== "boolean") {
                invalidFilterArguments(
                  "sheets_filter.visibleDropDown 必须是布尔值",
                  "visibleDropDown"
                );
              }
              var criteria1Present = hasOwn(args, "criteria1");
              var criteria2Present = hasOwn(args, "criteria2");
              if (criteria1Present && !isSafeFilterCriterion(args.criteria1, true)) {
                invalidFilterArguments(
                  "sheets_filter.criteria1 必须是标量或非空标量数组",
                  "criteria1"
                );
              }
              if (criteria2Present && !isSafeFilterCriterion(args.criteria2, false)) {
                invalidFilterArguments("sheets_filter.criteria2 必须是标量", "criteria2");
              }
              var operator;
              try {
                operator = normalizeFilterOperator(args.operator);
              } catch (error) {
                invalidFilterArguments(error.message, "operator");
              }
              if (operator !== undefined && !criteria1Present) {
                invalidFilterArguments(
                  "提供 sheets_filter.operator 时必须提供 criteria1",
                  "criteria1"
                );
              }
              if (criteria2Present && operator !== "xlAnd" && operator !== "xlOr") {
                invalidFilterArguments(
                  "sheets_filter.criteria2 仅允许与 xlAnd 或 xlOr 一起使用",
                  "criteria2"
                );
              }
              if (criteria2Present && !criteria1Present) {
                invalidFilterArguments(
                  "提供 sheets_filter.criteria2 时必须提供 criteria1",
                  "criteria1"
                );
              }
              if (operator === "xlFilterValues" && !Array.isArray(args.criteria1)) {
                invalidFilterArguments(
                  "xlFilterValues 需要非空数组 criteria1",
                  "criteria1"
                );
              }
              if (Array.isArray(args.criteria1) && operator !== "xlFilterValues") {
                invalidFilterArguments(
                  "数组 criteria1 仅允许与 xlFilterValues 一起使用",
                  "operator"
                );
              }
              if (
                operator === "xlTop10Items"
                || operator === "xlBottom10Items"
                || operator === "xlTop10Percent"
                || operator === "xlBottom10Percent"
              ) {
                var ranking = Number(args.criteria1);
                if (!isFinite(ranking) || ranking <= 0) {
                  invalidFilterArguments(
                    "排名筛选需要正数 criteria1",
                    "criteria1"
                  );
                }
              }
              if (
                operator === "xlFilterCellColor"
                || operator === "xlFilterFontColor"
                || operator === "xlFilterIcon"
              ) {
                invalidFilterArguments(
                  "颜色和图标筛选需要运行时对象，不能通过 JSON bridge 安全调用",
                  "operator"
                );
              }
              return { action: action, operator: operator };
            }

            function validateFilterTarget(sheet, args, expectedRangeIdentity) {
              var autoFilter = sheetAutoFilter(sheet);
              var targetRange = resolveFilterTargetRange(sheet, args, autoFilter);
              if (typeof targetRange.SetAutoFilter !== "function") {
                unsupportedFeature(
                  "sheets.autoFilter.set",
                  "当前 ONLYOFFICE 运行时不支持安全设置自动筛选"
                );
              }
              var shape = targetRangeShape(targetRange, "自动筛选");
              if (shape.rows < 2) {
                invalidFilterArguments(
                  "自动筛选区域必须包含表头和至少一行数据",
                  "range",
                  { rows: shape.rows, columns: shape.columns }
                );
              }
              if (args.field > shape.columns) {
                invalidFilterArguments(
                  "sheets_filter.field 超出筛选区域列数",
                  "field",
                  { fieldValue: args.field, columns: shape.columns }
                );
              }
              var targetIdentity = filterRangeIdentity(targetRange);
              if (!targetIdentity) {
                unsupportedFeature(
                  "sheets.autoFilter.rangeReadback",
                  "当前 ONLYOFFICE 运行时无法读取自动筛选目标区域"
                );
              }
              var currentRange = safeCall(autoFilter, "GetRange");
              var currentIdentity = filterRangeIdentity(currentRange);
              if (currentIdentity && currentIdentity !== targetIdentity) {
                invalidFilterArguments(
                  "目标区域与工作表现有自动筛选区域冲突",
                  "range",
                  { requestedRange: targetIdentity, existingRange: currentIdentity }
                );
              }
              if (expectedRangeIdentity && expectedRangeIdentity !== targetIdentity) {
                invalidFilterArguments(
                  "同一批次不能为同一工作表设置不同的自动筛选区域",
                  "range",
                  { requestedRange: targetIdentity, existingRange: expectedRangeIdentity }
                );
              }
              return { range: targetRange, identity: targetIdentity };
            }

            function validateExistingFilterAction(sheet, action) {
              var autoFilter = sheetAutoFilter(sheet);
              var currentRange = safeCall(autoFilter, "GetRange");
              if (!autoFilter || !currentRange) {
                invalidFilterArguments(
                  "action=" + action + " 需要工作表已有自动筛选",
                  "action"
                );
              }
              var method = action === "showAll" ? "ShowAllData" : "ApplyFilter";
              if (typeof autoFilter[method] !== "function") {
                unsupportedFeature(
                  "sheets.autoFilter." + action,
                  "当前 ONLYOFFICE 运行时不支持筛选操作 " + action
                );
              }
              if (!filterRangeIdentity(currentRange)) {
                unsupportedFeature(
                  "sheets.autoFilter.rangeReadback",
                  "当前 ONLYOFFICE 运行时无法读取现有自动筛选区域"
                );
              }
              return autoFilter;
            }

            function containsNull(value) {
              if (value === null) return true;
              if (!Array.isArray(value)) return false;
              for (var nullIndex = 0; nullIndex < value.length; nullIndex += 1) {
                if (containsNull(value[nullIndex])) return true;
              }
              return false;
            }

            function batchReferencesSheet(key, value, sheetName) {
              if (Array.isArray(value)) {
                for (var valueIndex = 0; valueIndex < value.length; valueIndex += 1) {
                  if (batchReferencesSheet(key, value[valueIndex], sheetName)) return true;
                }
                return false;
              }
              if (isObject(value)) {
                for (var nestedKey in value) {
                  if (hasOwn(value, nestedKey) && batchReferencesSheet(nestedKey, value[nestedKey], sheetName)) return true;
                }
                return false;
              }
              if (typeof value !== "string") return false;
              if (
                (key === "sheet" || key === "destinationSheet" || key === "beforeSheet" || key === "sourceSheet")
                && value === sheetName
              ) return true;
              if (
                key !== "formula" && key !== "formula1" && key !== "formula2"
                && key !== "categoryRange" && key !== "valuesRange" && key !== "xValuesRange"
              ) return false;
              var escaped = sheetName.replace(/'/g, "''");
              return value.indexOf("'" + escaped + "'!") >= 0 || value.indexOf(sheetName + "!") >= 0;
            }

            function preflightBatchDependencies(toolCalls) {
              var createdSheets = {};
              for (var dependencyIndex = 0; dependencyIndex < toolCalls.length; dependencyIndex += 1) {
                var dependencyCall = toolCalls[dependencyIndex] || {};
                var dependencyArgs = getArgs(dependencyCall);
                for (var createdName in createdSheets) {
                  if (!hasOwn(createdSheets, createdName)) continue;
                  var referenced = false;
                  for (var dependencyKey in dependencyArgs) {
                    if (hasOwn(dependencyArgs, dependencyKey) && batchReferencesSheet(dependencyKey, dependencyArgs[dependencyKey], createdName)) {
                      referenced = true;
                      break;
                    }
                  }
                  if (referenced) {
                    throw sheetError(
                      "BATCH_DEPENDENCY_REQUIRES_SPLIT",
                      "同一批次不能引用刚新增或刚重命名的工作表；请拆分批次",
                      {
                        sheet: createdName,
                        sourceToolCallIndex: createdSheets[createdName],
                        toolCallIndex: dependencyIndex,
                        completedToolCalls: 0,
                        partialMutationPossible: false,
                        mutationState: "none",
                      }
                    );
                  }
                }
                var newSheetName = null;
                if (dependencyCall.name === "sheets_add_sheet") newSheetName = dependencyArgs.name;
                else if (dependencyCall.name === "sheets_rename_sheet") newSheetName = dependencyArgs.newName;
                else if (dependencyCall.name === "sheets_manage_sheet" && dependencyArgs.action === "copy") newSheetName = dependencyArgs.newName;
                if (typeof newSheetName === "string" && newSheetName) createdSheets[newSheetName] = dependencyIndex;
              }
            }

            function preflightRuntimeCalls(toolCalls) {
              preflightBatchDependencies(toolCalls);
              var expectedFilterRanges = {};
              for (var preflightIndex = 0; preflightIndex < toolCalls.length; preflightIndex += 1) {
                var preflightCall = toolCalls[preflightIndex] || {};
                var preflightArgs = getArgs(preflightCall);
                if (preflightCall.name === "sheets_filter") {
                  var preflightFilter = validateFilterArguments(preflightArgs);
                  var preflightFilterSheet = getSheet(preflightArgs.sheet);
                  if (preflightFilter.action === "set") {
                    var filterSheetName = String(preflightFilterSheet.GetName());
                    var preflightFilterTarget = validateFilterTarget(
                      preflightFilterSheet,
                      preflightArgs,
                      expectedFilterRanges[filterSheetName]
                    );
                    expectedFilterRanges[filterSheetName] = preflightFilterTarget.identity;
                  } else {
                    validateExistingFilterAction(preflightFilterSheet, preflightFilter.action);
                  }
                } else if (preflightCall.name === "sheets_manage_table") {
                  var preflightTableSheet = getSheet(preflightArgs.sheet);
                  var preflightTableAction = String(preflightArgs.action || "");
                  if (preflightTableAction === "create") {
                    var preflightTableMode = String(preflightArgs.tableMode || "auto");
                    if (preflightTableMode === "basic") preflightTableMode = "rangeStyle";
                    var requiresNativeTable = preflightTableMode === "structured"
                      || structuredTableCreateFields(preflightArgs).length > 0;
                    if (requiresNativeTable && typeof preflightTableSheet.AddListObject !== "function") {
                      unsupportedFeature(
                        "sheets.nativeTables.create",
                        "当前 ONLYOFFICE 运行时不支持创建原生结构化表格",
                        { requestedTableMode: preflightTableMode }
                      );
                    }
                    if (
                      preflightTableMode === "rangeStyle"
                      || (
                        preflightTableMode === "auto"
                        && typeof preflightTableSheet.AddListObject !== "function"
                      )
                    ) {
                      var preflightRange = getRange(preflightTableSheet, preflightArgs.range);
                      if (
                        !preflightRange
                        || typeof preflightRange.SetBold !== "function"
                        || typeof preflightRange.SetFillColor !== "function"
                        || typeof preflightRange.SetBorders !== "function"
                      ) {
                        unsupportedFeature(
                          "sheets.rangeStyleTables.create",
                          "当前 ONLYOFFICE 运行时不支持普通区域表格样式"
                        );
                      }
                    }
                  } else if (
                    preflightTableAction === "update"
                    || preflightTableAction === "resize"
                    || preflightTableAction === "delete"
                    || preflightTableAction === "unlist"
                  ) {
                    if (typeof preflightTableSheet.GetListObjects !== "function") {
                      unsupportedFeature(
                        "sheets.nativeTables.inspect",
                        "当前 ONLYOFFICE 运行时不支持访问原生结构化表格"
                      );
                    }
                  }
                } else if (preflightCall.name === "sheets_manage_conditional_format") {
                  var preflightConditionalSheet = getSheet(preflightArgs.sheet);
                  var preflightConditionalRange = getRange(preflightConditionalSheet, preflightArgs.range);
                  if (!preflightConditionalRange || typeof preflightConditionalRange.GetFormatConditions !== "function") {
                    unsupportedFeature(
                      "sheets.conditionalFormatting.create",
                      "当前 ONLYOFFICE 运行时不支持条件格式"
                    );
                  }
                } else if (preflightCall.name === "sheets_set_values" && containsNull(preflightArgs.values)) {
                  var preflightValuesSheet = getSheet(preflightArgs.sheet);
                  var preflightValuesRange = getRange(preflightValuesSheet, preflightArgs.range);
                  if (!preflightValuesRange || typeof preflightValuesRange.ClearContents !== "function") {
                    unsupportedFeature(
                      "sheets.values.clearNull",
                      "当前 ONLYOFFICE 运行时无法按 null=清空单元格 的语义安全写入"
                    );
                  }
                } else if (preflightCall.name === "sheets_set_array_formula") {
                  var preflightArraySheet = getSheet(preflightArgs.sheet);
                  var preflightArrayRange = getRange(preflightArraySheet, preflightArgs.range);
                  if (!preflightArrayRange || typeof preflightArrayRange.SetFormulaArray !== "function") {
                    unsupportedFeature("sheets.arrayFormula.set", "当前 ONLYOFFICE 运行时不支持数组公式");
                  }
                } else if (preflightCall.name === "sheets_manage_range") {
                  var preflightManagedSheet = getSheet(preflightArgs.sheet);
                  var preflightManagedRange = getRange(preflightManagedSheet, preflightArgs.range);
                  var fillMethods = {
                    fillDown: "FillDown",
                    fillUp: "FillUp",
                    fillLeft: "FillLeft",
                    fillRight: "FillRight",
                  };
                  var fillMethod = fillMethods[String(preflightArgs.action || "")];
                  if (fillMethod && (!preflightManagedRange || typeof preflightManagedRange[fillMethod] !== "function")) {
                    unsupportedFeature(
                      "sheets.rangeFill." + String(preflightArgs.action),
                      "当前 ONLYOFFICE 运行时不支持区域操作 " + String(preflightArgs.action)
                    );
                  }
                } else if (preflightCall.name === "sheets_manage_validation") {
                  var preflightValidationRange = getRange(getSheet(preflightArgs.sheet), preflightArgs.range);
                  var preflightValidation = preflightValidationRange && typeof preflightValidationRange.GetValidation === "function"
                    ? preflightValidationRange.GetValidation()
                    : null;
                  var validationMethod = preflightArgs.action === "add" ? "Add" : (preflightArgs.action === "modify" ? "Modify" : "Delete");
                  if (!preflightValidation || typeof preflightValidation[validationMethod] !== "function") {
                    unsupportedFeature("sheets.validation.manage", "当前 ONLYOFFICE 运行时不支持数据验证操作");
                  }
                } else if (preflightCall.name === "sheets_inspect_comments") {
                  if (typeof getSheet(preflightArgs.sheet).GetComments !== "function") {
                    unsupportedFeature("sheets.comments.inspect", "当前 ONLYOFFICE 运行时不支持读取批注");
                  }
                } else if (preflightCall.name === "sheets_manage_comments" && preflightArgs.action === "add") {
                  var preflightCommentRange = getRange(getSheet(preflightArgs.sheet), preflightArgs.range);
                  if (!preflightCommentRange || typeof preflightCommentRange.AddComment !== "function") {
                    unsupportedFeature("sheets.comments.create", "当前 ONLYOFFICE 运行时不支持添加批注");
                  }
                } else if (preflightCall.name === "sheets_inspect_freeze_panes") {
                  if (typeof getSheet(preflightArgs.sheet).GetFreezePanes !== "function") {
                    unsupportedFeature("sheets.freezePanes.inspect", "当前 ONLYOFFICE 运行时不支持读取冻结窗格");
                  }
                } else if (preflightCall.name === "sheets_manage_freeze_panes") {
                  var preflightFreezeSheet = getSheet(preflightArgs.sheet);
                  var preflightFreeze = typeof preflightFreezeSheet.GetFreezePanes === "function" ? preflightFreezeSheet.GetFreezePanes() : null;
                  var freezeMethod = {
                    unfreeze: "Unfreeze",
                    freezeAt: "FreezeAt",
                    freezeRows: "FreezeRows",
                    freezeColumns: "FreezeColumns",
                  }[String(preflightArgs.action || "")];
                  if (!preflightFreeze || !freezeMethod || typeof preflightFreeze[freezeMethod] !== "function") {
                    unsupportedFeature("sheets.freezePanes.manage", "当前 ONLYOFFICE 运行时不支持冻结窗格操作");
                  }
                } else if (preflightCall.name === "sheets_inspect_charts") {
                  if (typeof getSheet(preflightArgs.sheet).GetAllCharts !== "function") {
                    unsupportedFeature("sheets.charts.inspect", "当前 ONLYOFFICE 运行时不支持读取图表");
                  }
                } else if (preflightCall.name === "sheets_update_chart") {
                  var preflightUpdateChart = resolveChart(preflightArgs);
                  if (!preflightUpdateChart.chart || typeof preflightUpdateChart.chart.SetTitle !== "function") {
                    unsupportedFeature("sheets.charts.update", "当前 ONLYOFFICE 运行时不支持更新图表");
                  }
                } else if (preflightCall.name === "sheets_add_chart") {
                  var preflightChartSheet = getSheet(preflightArgs.sheet);
                  if (typeof preflightChartSheet.AddChart !== "function") {
                    unsupportedFeature(
                      "sheets.charts.create",
                      "当前 ONLYOFFICE 运行时不支持创建图表"
                    );
                  }
                  if (
                    hasOwn(preflightArgs, "categoryRange")
                    || (Array.isArray(preflightArgs.addSeries) && preflightArgs.addSeries.length > 0)
                    || (Array.isArray(preflightArgs.seriesUpdates) && preflightArgs.seriesUpdates.length > 0)
                    || (Array.isArray(preflightArgs.removeSeries) && preflightArgs.removeSeries.length > 0)
                  ) {
                    unsupportedFeature(
                      "sheets.charts.addSeriesOnCreate",
                      "当前 ONLYOFFICE 运行时不支持在创建图表时追加或重定向系列；请先创建基础图表，再检查并更新",
                      { recovery: "create_inspect_then_update" }
                    );
                  }
                } else if (preflightCall.name === "sheets_delete_chart") {
                  var preflightChart = resolveChart(preflightArgs);
                  if (typeof preflightChart.chart.Delete !== "function") {
                    unsupportedFeature(
                      "sheets.charts.delete",
                      "当前 ONLYOFFICE 运行时不支持删除图表"
                    );
                  }
                }
              }
            }

            function a1ColumnLabel(columnNumber) {
              var remainingColumn = columnNumber;
              var label = "";
              while (remainingColumn > 0) {
                var remainder = (remainingColumn - 1) % 26;
                label = String.fromCharCode(65 + remainder) + label;
                remainingColumn = Math.floor((remainingColumn - 1) / 26);
              }
              return label;
            }

            function a1TopLeft(address) {
              var localAddress = String(address || "").split("!").pop().replace(/\$/g, "");
              var first = localAddress.split(":")[0];
              var match = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(first);
              if (!match) return null;
              var column = 0;
              for (var columnIndex = 0; columnIndex < match[1].length; columnIndex += 1) {
                column = column * 26 + match[1].toUpperCase().charCodeAt(columnIndex) - 64;
              }
              return { column: column, row: Number(match[2]) };
            }

            function writeValuesWithNull(sheet, targetRange, address, values) {
              if (!containsNull(values)) return targetRange.SetValue(values);
              var clearResult = targetRange.ClearContents();
              if (clearResult === false) throw new Error("ONLYOFFICE 拒绝清空 null 目标单元格");
              if (values === null) return true;
              var topLeft = a1TopLeft(address);
              if (!topLeft || !Array.isArray(values)) {
                throw sheetError(
                  "INVALID_TOOL_ARGUMENTS",
                  "包含 null 的 values 必须对应明确的 A1 二维区域",
                  { partialMutationPossible: false }
                );
              }
              for (var valueRowIndex = 0; valueRowIndex < values.length; valueRowIndex += 1) {
                var valueRow = values[valueRowIndex];
                var runStart = -1;
                for (var valueColumnIndex = 0; valueColumnIndex <= valueRow.length; valueColumnIndex += 1) {
                  var atEnd = valueColumnIndex === valueRow.length;
                  var hasValue = !atEnd && valueRow[valueColumnIndex] !== null;
                  if (hasValue && runStart < 0) runStart = valueColumnIndex;
                  if ((!hasValue || atEnd) && runStart >= 0) {
                    var runEnd = valueColumnIndex - 1;
                    var segmentStart = a1ColumnLabel(topLeft.column + runStart) + String(topLeft.row + valueRowIndex);
                    var segmentEnd = a1ColumnLabel(topLeft.column + runEnd) + String(topLeft.row + valueRowIndex);
                    var segmentRange = getRange(sheet, segmentStart === segmentEnd ? segmentStart : segmentStart + ":" + segmentEnd);
                    var segmentAccepted = segmentRange.SetValue([valueRow.slice(runStart, runEnd + 1)]);
                    if (segmentAccepted === false) throw new Error("ONLYOFFICE 拒绝写入非空单元格片段");
                    runStart = -1;
                  }
                }
              }
              return true;
            }

            preflightRuntimeCalls(calls);
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
                    range: describeRange(
                      inspectedRange,
                      args.includeValues !== false,
                      args.includeFormat === true,
                      args.includeConditionalFormats === true,
                      args.includeValidation === true
                    ),
                  });
                  break;
                }

                case "sheets_set_values": {
                  if (!args.range) throw new Error("sheets_set_values.range 不能为空");
                  if (args.values === undefined) throw new Error("sheets_set_values.values 不能为空");
                  var valuesSheet = getSheet(args.sheet);
                  var range = getRange(valuesSheet, args.range);
                  if (!range) throw new Error("无效单元格区域：" + args.range);
                  var valueShapes = validateSizedInput(
                    range,
                    args.values,
                    "values",
                    isSheetCellValue
                  );
                  var accepted = writeValuesWithNull(valuesSheet, range, args.range, args.values);
                  if (accepted === false) throw new Error("ONLYOFFICE 拒绝写入单元格值");
                  changed += 1;
                  results.push({
                    name: call.name,
                    sheet: valuesSheet.GetName(),
                    range: String(args.range),
                    accepted: accepted !== false,
                    inputShape: valueShapes.input,
                    targetShape: valueShapes.target,
                    readback: safeCall(range, "GetValue"),
                  });
                  break;
                }

                case "sheets_set_array_formula": {
                  if (!args.range || args.formula === undefined) throw new Error("sheets_set_array_formula 需要 range 和 formula");
                  var arrayFormulaSheet = getSheet(args.sheet);
                  var arrayFormulaRange = getRange(arrayFormulaSheet, args.range);
                  if (!arrayFormulaRange) throw new Error("无效单元格区域：" + args.range);
                  var arrayFormula = String(args.formula);
                  if (arrayFormula.charAt(0) !== "=") arrayFormula = "=" + arrayFormula;
                  mutationCall(arrayFormulaRange, "SetFormulaArray", "设置数组公式", arrayFormula);
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
                  var formulaRange = getRange(formulaSheet, args.range);
                  if (!formulaRange) throw new Error("无效单元格区域：" + args.range);
                  var formulaShapes = validateSizedInput(
                    formulaRange,
                    args.formula,
                    "formula",
                    isFormulaValue
                  );
                  var formula = normalizeFormulaInput(args.formula);
                  var formulaAccepted = typeof formulaRange.SetFormula === "function"
                    ? formulaRange.SetFormula(formula)
                    : formulaRange.SetValue(formula);
                  if (formulaAccepted === false) throw new Error("ONLYOFFICE 拒绝写入公式");
                  var formulaReadback = safeCall(formulaRange, "GetFormula");
                  changed += 1;
                  results.push({
                    name: call.name,
                    sheet: formulaSheet.GetName(),
                    range: String(args.range),
                    accepted: formulaAccepted !== false,
                    inputShape: formulaShapes.input,
                    targetShape: formulaShapes.target,
                    formula: formulaReadback,
                    readback: formulaReadback,
                  });
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
                  var columnWidthChars = unitAlias(args, "columnWidthChars", "columnWidth");
                  var rowHeightPt = unitAlias(args, "rowHeightPt", "rowHeight");
                  var formatFields = [
                    "fontSize",
                    "fontName",
                    "bold",
                    "italic",
                    "underline",
                    "strikeout",
                    "fontColor",
                    "fillColor",
                    "horizontalAlign",
                    "verticalAlign",
                    "numberFormat",
                    "wrap",
                    "orientation",
                  ];
                  var hasRequestedFormat = formatFields.some(function (field) {
                    return hasOwn(args, field);
                  }) || columnWidthChars !== undefined
                    || rowHeightPt !== undefined
                    || (Array.isArray(args.borders) && args.borders.length > 0);
                  if (!hasRequestedFormat) {
                    throw new Error("sheets_format_range 至少需要一个格式属性");
                  }
                  if (args.fontSize !== undefined) mutationCall(formatRange, "SetFontSize", "设置字号", Math.max(1, Number(args.fontSize)));
                  if (args.fontName) mutationCall(formatRange, "SetFontName", "设置字体", String(args.fontName));
                  if (args.bold !== undefined) mutationCall(formatRange, "SetBold", "设置粗体", Boolean(args.bold));
                  if (args.italic !== undefined) mutationCall(formatRange, "SetItalic", "设置斜体", Boolean(args.italic));
                  if (args.underline !== undefined) mutationCall(formatRange, "SetUnderline", "设置下划线", Boolean(args.underline));
                  if (args.strikeout !== undefined) mutationCall(formatRange, "SetStrikeout", "设置删除线", Boolean(args.strikeout));
                  if (args.fontColor) mutationCall(formatRange, "SetFontColor", "设置字体颜色", color(args.fontColor));
                  if (args.fillColor) mutationCall(formatRange, "SetFillColor", "设置单元格填充", color(args.fillColor));
                  if (args.horizontalAlign) mutationCall(formatRange, "SetAlignHorizontal", "设置水平对齐", String(args.horizontalAlign));
                  if (args.verticalAlign) mutationCall(formatRange, "SetAlignVertical", "设置垂直对齐", String(args.verticalAlign));
                  if (args.numberFormat) mutationCall(formatRange, "SetNumberFormat", "设置数字格式", String(args.numberFormat));
                  if (args.wrap !== undefined) mutationCall(formatRange, "SetWrap", "设置自动换行", Boolean(args.wrap));
                  if (args.orientation !== undefined) mutationCall(formatRange, "SetOrientation", "设置文本旋转", args.orientation);
                  if (columnWidthChars !== undefined) mutationCall(formatRange, "SetColumnWidth", "设置列宽", Number(columnWidthChars));
                  if (rowHeightPt !== undefined) mutationCall(formatRange, "SetRowHeight", "设置行高", Number(rowHeightPt));
                  if (Array.isArray(args.borders)) {
                    for (var borderIndex = 0; borderIndex < args.borders.length; borderIndex += 1) {
                      var border = args.borders[borderIndex] || {};
                      if (!border.side || !border.style || !border.color) throw new Error("每个边框需要 side、style 和 color");
                      mutationCall(
                        formatRange,
                        "SetBorders",
                        "设置单元格边框",
                        String(border.side),
                        String(border.style),
                        color(border.color)
                      );
                    }
                  }
                  changed += 1;
                  var formatRequested = requestedFormat(args, columnWidthChars, rowHeightPt);
                  var formatReadback = describeFormat(formatRange);
                  results.push({
                    name: call.name,
                    sheet: formatSheet.GetName(),
                    range: String(args.range),
                    requestedFormat: formatRequested,
                    readback: formatReadback,
                    verification: verifyReadback(formatRequested, formatReadback),
                  });
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
                  sortArguments.push(normalizeSortHeader(args.header));
                  sortArguments.push(normalizeSortOrientation(args.orientation));
                  mutationCall.apply(null, [sortRange, "SetSort", "排序"].concat(sortArguments));
                  changed += 1;
                  results.push({ name: call.name, sheet: sortSheet.GetName(), range: String(args.range), keys: args.keys });
                  break;
                }

                case "sheets_filter": {
                  var filterSheet = getSheet(args.sheet);
                  var validatedFilter = validateFilterArguments(args);
                  var filterAction = validatedFilter.action;
                  var expectedFilterRange = null;
                  if (filterAction === "set") {
                    var filterTarget = validateFilterTarget(filterSheet, args, null);
                    var filterRange = filterTarget.range;
                    expectedFilterRange = filterTarget.identity;
                    mutationCall(
                      filterRange,
                      "SetAutoFilter",
                      "设置自动筛选",
                      args.field,
                      args.criteria1,
                      validatedFilter.operator,
                      args.criteria2,
                      hasOwn(args, "visibleDropDown") ? Boolean(args.visibleDropDown) : undefined
                    );
                  } else if (filterAction === "showAll") {
                    var autoFilter = validateExistingFilterAction(filterSheet, filterAction);
                    mutationCall(autoFilter, "ShowAllData", "显示全部筛选数据");
                  } else if (filterAction === "reapply") {
                    var reappliedFilter = validateExistingFilterAction(filterSheet, filterAction);
                    mutationCall(reappliedFilter, "ApplyFilter", "重新应用筛选");
                  }
                  var inspectedFilter = sheetAutoFilter(filterSheet);
                  var inspectedFilterRange = safeCall(inspectedFilter, "GetRange");
                  var inspectedFilterIdentity = filterRangeIdentity(inspectedFilterRange);
                  if (
                    filterAction === "set"
                    && (!inspectedFilter || inspectedFilterIdentity !== expectedFilterRange)
                  ) {
                    throw sheetError(
                      "SHEETS_RUNTIME_INCOMPATIBLE",
                      "ONLYOFFICE 未能可靠读取刚设置的自动筛选区域",
                      {
                        feature: "sheets.autoFilter.setReadback",
                        expectedRange: expectedFilterRange,
                        observedRange: inspectedFilterIdentity,
                        retryable: false,
                        mutationState: "partial",
                        partialMutationPossible: true,
                        recovery: "inspect_then_undo_or_reopen",
                      }
                    );
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: filterAction,
                    sheet: filterSheet.GetName(),
                    range: describeRange(inspectedFilterRange, false),
                    filterMode: safeCall(inspectedFilter, "GetFilterMode"),
                  });
                  break;
                }

                case "sheets_inspect_tables": {
                  var tableInspectionSheets = args.sheet ? [getSheet(args.sheet)] : Api.GetSheets();
                  var remainingTables = Math.max(1, Math.min(1000, Number(args.maxTables || 200)));
                  var inspectedTableSheets = [];
                  var tableInspectionSupported = true;
                  for (var inspectedSheetIndex = 0; inspectedSheetIndex < tableInspectionSheets.length; inspectedSheetIndex += 1) {
                    var inspectedTableSheet = tableInspectionSheets[inspectedSheetIndex];
                    if (typeof inspectedTableSheet.GetListObjects !== "function") {
                      tableInspectionSupported = false;
                      inspectedTableSheets.push({
                        sheet: inspectedTableSheet.GetName(),
                        supported: false,
                        tableCount: 0,
                        tables: [],
                      });
                      continue;
                    }
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
                      supported: true,
                      tableCount: sheetTables.length,
                      tables: describedTables,
                    });
                  }
                  results.push({
                    name: call.name,
                    supported: tableInspectionSupported,
                    feature: "sheets.nativeTables.inspect",
                    reason: tableInspectionSupported ? null : "runtime_api_unavailable",
                    tables: [],
                    sheets: inspectedTableSheets,
                  });
                  break;
                }

                case "sheets_manage_table": {
                  if (!args.action) throw new Error("sheets_manage_table.action 不能为空");
                  var tableSheet = getSheet(args.sheet);
                  var tableAction = String(args.action);
                  var table;
                  var managedTableIndex = null;
                  var tableKind = "structured";
                  var tableDegraded = false;
                  var tableFormatted = false;
                  var requestedTableKind = "structured";
                  var actualTableKind = "structured";
                  var tableWarnings = [];
                  var rangeStyleResult = null;
                  if (tableAction === "create") {
                    if (!args.range) throw new Error("create 需要 range");
                    var invalidCreateFields = unexpectedArgumentFields(args, [
                      "action",
                      "sheet",
                      "range",
                      "tableMode",
                      "sourceType",
                      "name",
                      "style",
                      "showTotals",
                      "showHeaders",
                      "rowStripes",
                      "columnStripes",
                      "firstColumn",
                      "lastColumn",
                      "showAutoFilter",
                      "showAutoFilterDropDown",
                      "summary",
                      "alternativeText",
                    ]);
                    if (invalidCreateFields.length) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "create 不支持字段：" + invalidCreateFields.join("、"),
                        { partialMutationPossible: false }
                      );
                    }
                    var receivedTableMode = String(args.tableMode || "auto");
                    var tableMode = receivedTableMode === "basic" ? "rangeStyle" : receivedTableMode;
                    if (["auto", "structured", "rangeStyle"].indexOf(tableMode) === -1) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "tableMode 必须是 auto、structured、rangeStyle 或兼容别名 basic"
                      );
                    }
                    requestedTableKind = tableMode;
                    var structuredFields = structuredTableCreateFields(args);
                    if (tableMode === "rangeStyle" && structuredFields.length) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "普通区域样式不支持结构化属性：" + structuredFields.join("、")
                      );
                    }
                    var structuredFailure = null;
                    if (tableMode !== "rangeStyle") {
                      if (typeof tableSheet.AddListObject === "function") {
                        try {
                          var createdTable = tableSheet.AddListObject(
                            String(args.sourceType || "xlSrcRange"),
                            String(args.range)
                          );
                          if (isStructuredTableObject(createdTable)) {
                            table = createdTable;
                          } else {
                            structuredFailure = "AddListObject 未返回 ApiListObject";
                          }
                        } catch (error) {
                          structuredFailure = error && error.message
                            ? error.message
                            : String(error);
                        }
                      } else {
                        structuredFailure = "AddListObject 不可用";
                      }
                    }
                    if (!table) {
                      if (tableMode === "structured" || structuredFields.length) {
                        throw sheetError(
                          "SHEETS_API_UNSUPPORTED",
                          "当前 ONLYOFFICE 版本无法创建结构化表格"
                            + (structuredFailure ? "：" + structuredFailure : ""),
                          {
                            feature: "sheets.nativeTables.create",
                            retryable: false,
                            mutationState: "none",
                            requestedTableMode: tableMode,
                            unsupportedFields: structuredFields,
                            partialMutationPossible: false,
                          }
                        );
                      }
                      rangeStyleResult = formatRangeStyleTable(tableSheet, args.range);
                      tableKind = "rangeStyle";
                      actualTableKind = "rangeStyle";
                      tableDegraded = tableMode === "auto";
                      tableFormatted = true;
                      if (tableDegraded) tableWarnings.push("native_table_unavailable");
                      tableWarnings = tableWarnings.concat(rangeStyleResult.warnings || []);
                    }
                    if (table && typeof table === "object") applyTableSettings(table, args, true);
                  } else if (tableAction === "format") {
                    if (!args.range) throw new Error("format 需要 range");
                    var formatFields = unexpectedArgumentFields(args, [
                      "action",
                      "sheet",
                      "range",
                    ]);
                    if (formatFields.length) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "action:format 只支持 sheet 和 range；不支持："
                          + formatFields.join("、")
                      );
                    }
                    rangeStyleResult = formatRangeStyleTable(tableSheet, args.range);
                    table = null;
                    tableKind = "rangeStyle";
                    requestedTableKind = "rangeStyle";
                    actualTableKind = "rangeStyle";
                    tableFormatted = true;
                    tableWarnings = tableWarnings.concat(rangeStyleResult.warnings || []);
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
                    tableKind: tableKind,
                    requestedKind: requestedTableKind,
                    actualKind: actualTableKind,
                    degraded: tableDegraded,
                    warnings: tableWarnings,
                    formatted: tableFormatted || undefined,
                    filterApplied: rangeStyleResult ? rangeStyleResult.filterApplied : undefined,
                    range: tableKind === "rangeStyle"
                      ? describeRange(getRange(tableSheet, args.range), false)
                      : undefined,
                    table: tableAction === "delete"
                      || tableAction === "unlist"
                      || tableKind === "rangeStyle"
                      ? null
                      : describeTable(table, managedTableIndex),
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
                  var conditionExternalIndex = null;
                  var conditionCountBefore = safeCall(conditions, "GetCount");
                  if (conditionalAction === "deleteAll") {
                    mutationCall(conditions, "Delete", "删除条件格式");
                    var deletedConditionCount = safeCall(conditions, "GetCount");
                    if (typeof deletedConditionCount === "number" && deletedConditionCount !== 0) {
                      throw sheetError(
                        "EXECUTION_FAILED",
                        "删除条件格式后规则数量仍为 " + deletedConditionCount
                      );
                    }
                  } else if (conditionalAction === "add") {
                    var conditionType = String(args.type || "cellValue");
                    if (conditionType === "cellValue" || conditionType === "expression") {
                      condition = objectMutationCall(
                        conditions,
                        "Add",
                        "添加条件格式规则",
                        conditionType === "expression" ? "xlExpression" : "xlCellValue",
                        normalizeComparisonOperator(args.operator, "xlGreater", "条件格式运算符"),
                        args.formula1,
                        args.formula2
                      );
                    } else if (conditionType === "uniqueValues" || conditionType === "duplicateValues") {
                      condition = objectMutationCall(conditions, "AddUniqueValues", "添加重复值条件格式");
                      mutationCall(
                        condition,
                        "SetDupeUnique",
                        "设置重复值条件格式类型",
                        conditionType === "duplicateValues" ? "xlDuplicate" : "xlUnique"
                      );
                    } else if (conditionType === "colorScale") {
                      condition = objectMutationCall(
                        conditions,
                        "AddColorScale",
                        "添加色阶条件格式",
                        clamp(Math.round(asFinite(args.scale, 3)), 2, 3)
                      );
                    } else if (conditionType === "dataBar") {
                      condition = objectMutationCall(conditions, "AddDatabar", "添加数据条条件格式");
                    } else if (conditionType === "iconSet") {
                      condition = objectMutationCall(conditions, "AddIconSetCondition", "添加图标集条件格式");
                    } else if (conditionType === "top10") {
                      condition = objectMutationCall(conditions, "AddTop10", "添加前十项条件格式");
                    } else if (conditionType === "aboveAverage") {
                      condition = objectMutationCall(conditions, "AddAboveAverage", "添加高于平均值条件格式");
                    } else {
                      throw new Error("不支持的条件格式类型：" + conditionType);
                    }
                    if (args.fillColor) mutationCall(condition, "SetFillColor", "设置条件格式填充", color(args.fillColor));
                    var conditionFont = safeCall(condition, "GetFont");
                    if (conditionFont) {
                      if (args.fontColor) mutationCall(conditionFont, "SetColor", "设置条件格式字体颜色", color(args.fontColor));
                      if (hasOwn(args, "bold")) mutationCall(conditionFont, "SetBold", "设置条件格式粗体", Boolean(args.bold));
                    }
                    var conditionCountAfter = safeCall(conditions, "GetCount");
                    if (
                      typeof conditionCountBefore === "number"
                      && typeof conditionCountAfter === "number"
                      && conditionCountAfter !== conditionCountBefore + 1
                    ) {
                      throw sheetError(
                        "EXECUTION_FAILED",
                        "添加条件格式后规则数量未增加"
                      );
                    }
                    if (typeof conditionCountAfter === "number" && conditionCountAfter > 0) {
                      conditionExternalIndex = conditionCountAfter - 1;
                      var addedConditionReadback = safeCall(
                        conditions,
                        "GetItem",
                        conditionCountAfter
                      );
                      if (addedConditionReadback) condition = addedConditionReadback;
                    }
                  } else {
                    throw new Error("条件格式 action 必须是 add 或 deleteAll");
                  }
                  changed += 1;
                  var conditionalCount = safeCall(conditions, "GetCount");
                  var conditionalRule = condition
                    ? describeCondition(condition, conditionExternalIndex)
                    : null;
                  var expectedConditionalRule = conditionalAction === "deleteAll"
                    ? { count: 0 }
                    : conditionRuleExpectation(
                      args,
                      String(args.type || "cellValue"),
                      typeof conditionCountBefore === "number"
                        ? conditionCountBefore + 1
                        : undefined
                    );
                  var observedConditionalRule = conditionRuleObservation(
                    conditionalRule,
                    conditionalCount
                  );
                  results.push({
                    name: call.name,
                    action: conditionalAction,
                    sheet: conditionalSheet.GetName(),
                    range: String(args.range),
                    count: conditionalCount,
                    rule: conditionalRule,
                    conditionalFormats: describeConditions(conditionalRange),
                    ruleVerification: verifyReadback(
                      expectedConditionalRule,
                      observedConditionalRule
                    ),
                    effectiveStyleVerification: {
                      status: "unavailable",
                      reason: "effective-cell-style-requires-visual-verification",
                    },
                  });
                  break;
                }

                case "sheets_manage_validation": {
                  if (!args.action || !args.range) throw new Error("sheets_manage_validation 需要 action 和 range");
                  var validationSheet = getSheet(args.sheet);
                  var validationRange = getRange(validationSheet, args.range);
                  var validation = requireMethod(validationRange, "GetValidation", "数据验证").call(validationRange);
                  var validationAction = String(args.action);
                  var validationResult = validation;
                  if (validationAction === "delete") {
                    mutationCall(validation, "Delete", "删除数据验证");
                  } else if (validationAction === "add" || validationAction === "modify") {
                    if (!hasOwn(args, "type")) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "数据验证 add/modify 必须显式提供 type"
                      );
                    }
                    var validationMethod = validationAction === "add" ? "Add" : "Modify";
                    var validationType = normalizeValidationType(args.type);
                    var validationOperator = normalizeComparisonOperator(
                      args.operator,
                      "xlBetween",
                      "数据验证运算符"
                    );
                    var validationNeedsFormula = validationType !== "xlValidateInputOnly";
                    if (validationNeedsFormula && !hasOwn(args, "formula1")) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "当前数据验证类型必须提供 formula1"
                      );
                    }
                    var validationComparisonTypes = {
                      xlValidateWholeNumber: true,
                      xlValidateDecimal: true,
                      xlValidateDate: true,
                      xlValidateTime: true,
                      xlValidateTextLength: true,
                    };
                    if (
                      validationComparisonTypes[validationType]
                      && (validationOperator === "xlBetween" || validationOperator === "xlNotBetween")
                      && !hasOwn(args, "formula2")
                    ) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "between/notBetween 数据验证必须提供 formula2"
                      );
                    }
                    validationResult = objectMutationCall(
                      validation,
                      validationMethod,
                      validationAction === "add" ? "添加数据验证规则" : "修改数据验证规则",
                      validationType,
                      normalizeValidationAlertStyle(args.alertStyle),
                      validationOperator,
                      args.formula1,
                      args.formula2
                    );
                    if (hasOwn(args, "ignoreBlank")) mutationCall(validationResult, "SetIgnoreBlank", "设置数据验证空值策略", Boolean(args.ignoreBlank));
                    if (hasOwn(args, "showInput")) mutationCall(validationResult, "SetShowInput", "设置数据验证输入提示", Boolean(args.showInput));
                    if (hasOwn(args, "showError")) mutationCall(validationResult, "SetShowError", "设置数据验证错误提示", Boolean(args.showError));
                    if (args.inputTitle !== undefined) mutationCall(validationResult, "SetInputTitle", "设置数据验证输入标题", String(args.inputTitle));
                    if (args.inputMessage !== undefined) mutationCall(validationResult, "SetInputMessage", "设置数据验证输入消息", String(args.inputMessage));
                    if (args.errorTitle !== undefined) mutationCall(validationResult, "SetErrorTitle", "设置数据验证错误标题", String(args.errorTitle));
                    if (args.errorMessage !== undefined) mutationCall(validationResult, "SetErrorMessage", "设置数据验证错误消息", String(args.errorMessage));
                  } else {
                    throw new Error("数据验证 action 必须是 add、modify 或 delete");
                  }
                  var validationReadback = describeValidation(
                    safeCall(validationRange, "GetValidation") || validationResult
                  );
                  if (validationAction !== "delete" && !validationReadback) {
                    throw sheetError(
                      "EXECUTION_FAILED",
                      "数据验证写入后无法读回有效规则"
                    );
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: validationAction,
                    sheet: validationSheet.GetName(),
                    range: String(args.range),
                    rule: validationAction === "delete" ? null : validationReadback,
                    validation: validationReadback,
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
                  var resolvedChartType = resolveChartType(args.type, "bar", "type");
                  var chart = null;
                  try {
                    chart = chartSheet.AddChart(
                      source,
                      Boolean(args.inRows),
                      resolvedChartType.normalized,
                      clamp(Math.round(asFinite(args.style, 2)), 1, 48),
                      Math.max(30, asFinite(args.widthMm, 120)) * EMU_PER_MM,
                      Math.max(20, asFinite(args.heightMm, 70)) * EMU_PER_MM,
                      Math.max(0, Math.round(asFinite(args.fromColumn, 0))),
                      mmToEmu(args.columnOffsetMm || 0),
                      Math.max(0, Math.round(asFinite(args.fromRow, 0))),
                      mmToEmu(args.rowOffsetMm || 0)
                    );
                    if (!chart) throw new Error("ONLYOFFICE 未返回图表对象");
                  } catch (chartError) {
                    throw chartCreationError(chartError, {
                      sheet: chartSheet.GetName(),
                      range: source,
                      receivedType: resolvedChartType.received,
                      normalizedType: resolvedChartType.normalized,
                    });
                  }
                  try {
                    applyChartOptions(chart, args, true);
                  } catch (chartOptionsError) {
                    var partialChartIndex = chartIndexOnSheet(chartSheet, chart);
                    var partialChartName = safeCall(chart, "GetName");
                    var chartRolledBack = false;
                    if (typeof chart.Delete === "function") {
                      try {
                        chartRolledBack = chart.Delete() !== false;
                      } catch (chartRollbackError) {
                        chartRolledBack = false;
                      }
                    }
                    if (chartRolledBack) {
                      throw sheetError(
                        "SHEETS_CHART_CREATE_FAILED",
                        "创建图表后应用选项失败，已回滚新图表："
                          + (chartOptionsError && chartOptionsError.message ? chartOptionsError.message : String(chartOptionsError)),
                        {
                          phase: "sheets-chart-options",
                          feature: "sheets.charts.create",
                          rolledBack: true,
                          partialMutationPossible: false,
                        }
                      );
                    }
                    throw sheetError(
                      "SHEETS_CHART_PARTIAL_MUTATION",
                      "图表已创建，但后续选项应用失败；当前运行时无法安全删除该对象",
                      {
                        phase: "sheets-chart-options",
                        feature: "sheets.charts.create",
                        created: true,
                        chartIndex: partialChartIndex,
                        chartName: partialChartName,
                        retryable: false,
                        mutationState: "partial",
                        recovery: "inspect_existing_chart_before_retry",
                        causeMessage: chartOptionsError && chartOptionsError.message
                          ? chartOptionsError.message
                          : String(chartOptionsError),
                        partialMutationPossible: true,
                      }
                    );
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    sheet: chartSheet.GetName(),
                    range: source,
                    type: resolvedChartType.normalized,
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
                  if (typeof chartToDelete.chart.Delete !== "function") {
                    throw sheetError(
                      "SHEETS_API_UNSUPPORTED",
                      "当前 ONLYOFFICE 运行时不支持删除图表",
                      {
                        feature: "sheets.charts.delete",
                        retryable: false,
                        mutationState: "none",
                        partialMutationPossible: false,
                      }
                    );
                  }
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
                  var inspectedFreezeState = describeFreezeState(freezePanes);
                  results.push({
                    name: call.name,
                    sheet: freezeSheet.GetName(),
                    location: inspectedFreezeState.location,
                    frozenRows: inspectedFreezeState.frozenRows,
                    frozenColumns: inspectedFreezeState.frozenColumns,
                    topLeftCell: inspectedFreezeState.topLeftCell,
                    verified: true,
                  });
                  break;
                }

                case "sheets_manage_freeze_panes": {
                  if (!args.action) throw new Error("sheets_manage_freeze_panes.action 不能为空");
                  var managedFreezeSheet = getSheet(args.sheet);
                  var managedFreeze = requireMethod(managedFreezeSheet, "GetFreezePanes", "冻结窗格").call(managedFreezeSheet);
                  var freezeAction = String(args.action);
                  if (freezeAction === "unfreeze") {
                    if (hasOwn(args, "range") || hasOwn(args, "count")) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "unfreeze 不接受 range 或 count",
                        { partialMutationPossible: false }
                      );
                    }
                    requireMethod(managedFreeze, "Unfreeze", "取消冻结窗格").call(managedFreeze);
                  }
                  else if (freezeAction === "freezeAt") {
                    if (!args.range) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "freezeAt 需要 range",
                        { partialMutationPossible: false }
                      );
                    }
                    if (hasOwn(args, "count")) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "freezeAt 不接受 count",
                        { partialMutationPossible: false }
                      );
                    }
                    requireMethod(managedFreeze, "FreezeAt", "按区域冻结窗格").call(managedFreeze, getRange(managedFreezeSheet, args.range));
                  } else if (freezeAction === "freezeRows") {
                    if (hasOwn(args, "range")) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "freezeRows 不接受 range",
                        { partialMutationPossible: false }
                      );
                    }
                    var frozenRowCount = asFinite(args.count, NaN);
                    if (
                      !isFinite(frozenRowCount)
                      || Math.floor(frozenRowCount) !== frozenRowCount
                      || frozenRowCount < 1
                    ) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "freezeRows.count 必须是正整数；解冻请使用 unfreeze",
                        { partialMutationPossible: false }
                      );
                    }
                    requireMethod(managedFreeze, "FreezeRows", "冻结行").call(managedFreeze, frozenRowCount);
                  } else if (freezeAction === "freezeColumns") {
                    if (hasOwn(args, "range")) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "freezeColumns 不接受 range",
                        { partialMutationPossible: false }
                      );
                    }
                    var frozenColumnCount = asFinite(args.count, NaN);
                    if (
                      !isFinite(frozenColumnCount)
                      || Math.floor(frozenColumnCount) !== frozenColumnCount
                      || frozenColumnCount < 1
                    ) {
                      throw sheetError(
                        "INVALID_TOOL_ARGUMENTS",
                        "freezeColumns.count 必须是正整数；解冻请使用 unfreeze",
                        { partialMutationPossible: false }
                      );
                    }
                    requireMethod(managedFreeze, "FreezeColumns", "冻结列").call(managedFreeze, frozenColumnCount);
                  } else {
                    throw sheetError(
                      "INVALID_TOOL_ARGUMENTS",
                      "不支持的冻结窗格操作：" + freezeAction,
                      { partialMutationPossible: false }
                    );
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: freezeAction,
                    sheet: managedFreezeSheet.GetName(),
                    location: null,
                    frozenRows: null,
                    frozenColumns: null,
                    topLeftCell: null,
                    verified: false,
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
                  var inspectedDisplayState = describeDisplayState(pageSheet, false);
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
                    displayGridlines: inspectedDisplayState.displayGridlines,
                    displayHeadings: inspectedDisplayState.displayHeadings,
                    verified: true,
                  });
                  break;
                }

                case "sheets_manage_page_layout": {
                  var managedMarginsPt = unitAlias(args, "marginsPt", "margins");
                  var hasManagedMargins = isObject(managedMarginsPt) && [
                    "top",
                    "right",
                    "bottom",
                    "left",
                  ].some(function (side) {
                    return hasOwn(managedMarginsPt, side);
                  });
                  if (
                    args.orientation === undefined
                    && !hasManagedMargins
                    && !hasOwn(args, "printGridlines")
                    && !hasOwn(args, "printHeadings")
                    && !hasOwn(args, "displayGridlines")
                    && !hasOwn(args, "displayHeadings")
                  ) {
                    throw new Error("sheets_manage_page_layout 至少需要一个页面布局属性");
                  }
                  var managedPageSheet = getSheet(args.sheet);
                  if (args.orientation !== undefined) {
                    mutationCall(
                      managedPageSheet,
                      "SetPageOrientation",
                      "设置页面方向",
                      normalizePageOrientation(args.orientation)
                    );
                  }
                  if (hasManagedMargins) {
                    if (hasOwn(managedMarginsPt, "top")) mutationCall(managedPageSheet, "SetTopMargin", "设置上页边距", Number(managedMarginsPt.top));
                    if (hasOwn(managedMarginsPt, "right")) mutationCall(managedPageSheet, "SetRightMargin", "设置右页边距", Number(managedMarginsPt.right));
                    if (hasOwn(managedMarginsPt, "bottom")) mutationCall(managedPageSheet, "SetBottomMargin", "设置下页边距", Number(managedMarginsPt.bottom));
                    if (hasOwn(managedMarginsPt, "left")) mutationCall(managedPageSheet, "SetLeftMargin", "设置左页边距", Number(managedMarginsPt.left));
                  }
                  if (hasOwn(args, "printGridlines")) mutationCall(managedPageSheet, "SetPrintGridlines", "设置打印网格线", Boolean(args.printGridlines));
                  if (hasOwn(args, "printHeadings")) mutationCall(managedPageSheet, "SetPrintHeadings", "设置打印行列标题", Boolean(args.printHeadings));
                  if (hasOwn(args, "displayGridlines")) mutationCall(managedPageSheet, "SetDisplayGridlines", "设置屏幕网格线", Boolean(args.displayGridlines));
                  if (hasOwn(args, "displayHeadings")) mutationCall(managedPageSheet, "SetDisplayHeadings", "设置屏幕行列标题", Boolean(args.displayHeadings));
                  changed += 1;
                  var persistedMarginsPt = {
                    top: safeCall(managedPageSheet, "GetTopMargin"),
                    right: safeCall(managedPageSheet, "GetRightMargin"),
                    bottom: safeCall(managedPageSheet, "GetBottomMargin"),
                    left: safeCall(managedPageSheet, "GetLeftMargin"),
                  };
                  var managedDisplayState = describeDisplayState(managedPageSheet, true);
                  var needsDisplayVerification = hasOwn(args, "displayGridlines")
                    || hasOwn(args, "displayHeadings");
                  results.push({
                    name: call.name,
                    sheet: managedPageSheet.GetName(),
                    orientation: safeCall(managedPageSheet, "GetPageOrientation"),
                    marginsPt: persistedMarginsPt,
                    margins: persistedMarginsPt,
                    printGridlines: safeCall(managedPageSheet, "GetPrintGridlines"),
                    printHeadings: safeCall(managedPageSheet, "GetPrintHeadings"),
                    displayGridlines: managedDisplayState.displayGridlines,
                    displayHeadings: managedDisplayState.displayHeadings,
                    verified: !needsDisplayVerification,
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
            var failedCallIndex = typeof callIndex === "number"
              && callIndex >= 0
              && calls
              && callIndex < calls.length
              ? callIndex
              : null;
            var failedCall = failedCallIndex !== null && calls[failedCallIndex]
              ? calls[failedCallIndex]
              : null;
            var priorMutationPossible = false;
            if (failedCallIndex !== null) {
              for (var mutationIndex = 0; mutationIndex < failedCallIndex; mutationIndex += 1) {
                if (calls[mutationIndex] && mutatingNames[calls[mutationIndex].name]) {
                  priorMutationPossible = true;
                  break;
                }
              }
            }
            var currentMutationPossible = Boolean(
              failedCall
              && mutatingNames[failedCall.name]
            );
            var existingDetails = error
              && error.details
              && typeof error.details === "object"
              ? error.details
              : {};
            if (typeof existingDetails.partialMutationPossible === "boolean") {
              currentMutationPossible = existingDetails.partialMutationPossible;
            }
            var details = {};
            Object.keys(existingDetails).forEach(function (name) {
              details[name] = existingDetails[name];
            });
            details.phase = existingDetails.phase || "sheets-command";
            details.completedToolCalls = failedCallIndex === null ? 0 : failedCallIndex;
            details.partialMutationPossible = priorMutationPossible || currentMutationPossible;
            if (failedCall && failedCall.name) details.tool = String(failedCall.name);
            if (failedCallIndex !== null) details.toolCallIndex = failedCallIndex;
            var errorMessage = error && error.message ? error.message : String(error);
            var normalizedErrorCode = error && error.code ? error.code : "EXECUTION_FAILED";
            if (errorMessage.indexOf("setDirtyConditionalFormatting") >= 0) {
              normalizedErrorCode = "SHEETS_RUNTIME_INCOMPATIBLE";
              details.feature = "sheets.conditionalFormatting";
              details.retryable = false;
              details.mutationState = details.partialMutationPossible ? "partial" : "unknown";
              details.recovery = "inspect_then_undo_or_reopen";
            }
            return JSON.stringify({
              ok: false,
              code: normalizedErrorCode,
              message: errorMessage,
              error: errorMessage,
              details: details,
            });
          }
        },
        false,
        true,
        function (rawResult) {
          try {
            const executionResult = parseResult(rawResult);
            verifyViewChanges(toolCalls, executionResult).then(resolve, reject);
          } catch (error) {
            reject(error);
          }
        },
      );
    });
  }

  window.AICopilotBridges.cell = {
    execute: execute,
    preflight: preflight,
    probeCapabilities: probeCapabilities,
    inspect: function () {
      return execute([{ name: "sheets_inspect", arguments: { maxCells: 1200 } }]);
    },
  };
})();
