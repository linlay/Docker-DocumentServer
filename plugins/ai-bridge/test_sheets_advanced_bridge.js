"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const bridgeRoot = process.env.AI_BRIDGE_DIR || __dirname;
const sheetsSource = fs.readFileSync(path.join(bridgeRoot, "bridges/sheets-bridge.js"), "utf8");
const pluginSource = fs.readFileSync(path.join(bridgeRoot, "plugin.js"), "utf8");
const hostSource = fs.readFileSync(path.join(bridgeRoot, "host-bridge.js"), "utf8");
const clientSource = fs.readFileSync(path.join(bridgeRoot, "client-sdk.js"), "utf8");
const declarationsSource = fs.readFileSync(path.join(bridgeRoot, "public-api.d.ts"), "utf8");
const publicContract = JSON.parse(fs.readFileSync(path.join(bridgeRoot, "public-api.json"), "utf8"));

function sheetsToolNames(source) {
  return [...new Set(source.match(/sheets_[a-z_]+/g) || [])].sort();
}

function a1Shape(address) {
  const match = String(address).match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i);
  if (!match) return { rows: 1, columns: 1 };
  const columnNumber = letters => [...letters.toUpperCase()]
    .reduce((value, letter) => (value * 26) + letter.charCodeAt(0) - 64, 0);
  const startColumn = columnNumber(match[1]);
  const endColumn = columnNumber(match[3] || match[1]);
  const startRow = Number(match[2]);
  const endRow = Number(match[4] || match[2]);
  return {
    rows: Math.abs(endRow - startRow) + 1,
    columns: Math.abs(endColumn - startColumn) + 1,
  };
}

test("all 43 Sheets tools stay aligned across the contract, plugin, SDKs, types, and adapter", () => {
  const expected = Object.keys(publicContract.tools.cell).sort();
  assert.equal(expected.length, 43);
  assert.deepEqual(sheetsToolNames(pluginSource), expected);
  assert.deepEqual(sheetsToolNames(hostSource), expected);
  assert.deepEqual(sheetsToolNames(clientSource), expected);
  assert.deepEqual(sheetsToolNames(declarationsSource), expected);
  for (const name of expected) {
    assert.match(sheetsSource, new RegExp(`["']${name}["']`));
  }
});

test("public Sheets contract names non-obvious units explicitly", () => {
  const format = publicContract.tools.cell.sheets_format_range.properties;
  assert.match(format.columnWidthChars.description, /character units/);
  assert.match(format.rowHeightPt.description, /points \(pt\)/);
  assert.equal(format.columnWidth.deprecated, true);
  assert.equal(format.rowHeight.deprecated, true);

  const pageLayout = publicContract.tools.cell.sheets_manage_page_layout;
  const freezePanes = publicContract.tools.cell.sheets_manage_freeze_panes;
  assert.ok(pageLayout.anyOf.some(entry => entry.required[0] === "marginsPt"));
  assert.match(pageLayout.properties.marginsPt.description, /points \(pt\)/);
  assert.equal(pageLayout.properties.margins.deprecated, true);
  assert.match(publicContract.unitConventions.sheet, /columnWidthChars in spreadsheet character-width units/);

  assert.equal(
    publicContract.tools.cell.sheets_set_formula.properties.formula.$ref,
    "#/$defs/sheetFormula",
  );
  assert.equal(
    publicContract.tools.cell.sheets_manage_conditional_format.properties.operator.$ref,
    "#/$defs/sheetComparisonOperator",
  );
  assert.ok(pageLayout.anyOf.some(entry => entry.required[0] === "displayGridlines"));
  assert.match(pageLayout.properties.printGridlines.description, /printed output only/);
  assert.match(pageLayout.properties.displayGridlines.description, /screen view/);
  assert.equal(freezePanes.properties.count.minimum, 1);
  assert.match(freezePanes.properties.range.description, /Exact/);
  assert.match(freezePanes.description, /verified:true/);
  assert.match(
    publicContract.tools.cell.sheets_inspect_freeze_panes.description,
    /frozenRows.*verified:true/,
  );
  assert.match(pageLayout.description, /verified:true/);
  assert.match(
    publicContract.tools.cell.sheets_inspect_page_layout.description,
    /真实布尔值.*verified:true/,
  );
  assert.ok(freezePanes.oneOf.some(entry => entry.required && entry.required[0] === "count"));
  assert.equal(
    publicContract.inputNormalization.sheet.formulaScalarMultiCellPolicy,
    "reject-scalar-multi-cell",
  );
});

function loadSheetsBridge(api, macroState, bridgeOptions = {}) {
  const asc = {
    scope: {},
    plugin: {
      info: {},
      callCommand(command, _close, _calc, callback) {
        callback(command());
      },
      executeMethod(method, params, callback) {
        macroState.calls.push([method, params]);
        if (method === "GetMacros") callback(JSON.stringify(macroState.content));
        else if (method === "GetVBAMacros") callback("<Module Name=\"Demo\"><SourceCode>Sub Demo()</SourceCode></Module>");
        else {
          macroState.content = JSON.parse(params[0]);
          callback(true);
        }
      },
    },
  };
  if (!bridgeOptions.withoutAscEditor) {
    asc.editor = {
      SetDocumentModified(value) {
        if (typeof bridgeOptions.onDocumentModified === "function") {
          bridgeOptions.onDocumentModified(value);
        }
      },
    };
  }
  const window = {
    Asc: asc,
    setTimeout,
    AICopilotSheetsViewVerificationIntervalMs: bridgeOptions.verifyIntervalMs,
    AICopilotSheetsViewVerificationTimeoutMs: bridgeOptions.verifyTimeoutMs,
  };
  const sandbox = {
    window,
    Asc: asc,
    Api: api,
    console,
    JSON,
    Number,
    String,
    Boolean,
    Math,
    RegExp,
    Error,
    Date,
    Object,
    Array,
    isFinite,
    isNaN,
    parseInt,
    setTimeout,
  };
  vm.runInNewContext(sheetsSource, sandbox);
  return window.AICopilotBridges.cell;
}

function createHarness(options = {}) {
  const state = {
    events: [],
    sheets: [],
    names: [],
    freezeDelayMs: Number(options.freezeDelayMs) || 0,
    freezeNeverApplies: options.freezeNeverApplies === true,
    pivots: [],
    core: {},
    custom: {},
    macros: {
      calls: [],
      content: { macrosArray: [{ name: "Demo", value: "return 1;", guid: "macro-1" }], current: 0 },
    },
    historyPoints: 0,
  };

  function event(name, ...args) {
    state.events.push([name, ...args]);
  }

  function font() {
    return {
      SetBold: value => event("font.bold", value),
      SetItalic: value => event("font.italic", value),
      SetUnderline: value => event("font.underline", value),
      SetStrikethrough: value => event("font.strikeout", value),
      SetName: value => event("font.name", value),
      SetSize: value => event("font.size", value),
      SetColor: value => event("font.color", value),
      GetName: () => "Arial",
    };
  }

  function comment(text = "comment") {
    const value = {
      text,
      replies: [],
      solved: false,
      GetId: () => "comment-1",
      GetText: () => value.text,
      GetQuoteText: () => "A1",
      GetAuthorName: () => "Tester",
      GetUserId: () => "user-1",
      IsSolved: () => value.solved,
      GetTime: () => new Date("2025-01-01T00:00:00Z"),
      GetRepliesCount: () => value.replies.length,
      GetReply: index => value.replies[index],
      SetText: next => { value.text = next; event("comment.text", next); },
      SetAuthorName: next => event("comment.author", next),
      SetUserId: next => event("comment.user", next),
      SetSolved: next => { value.solved = next; event("comment.solved", next); },
      AddReply: next => {
        value.replies.push({
          GetText: () => next,
          GetAuthorName: () => "Tester",
          GetUserId: () => "user-1",
          GetTime: () => new Date("2025-01-01T00:00:00Z"),
        });
        event("comment.reply", next);
      },
      RemoveReplies: (start, count) => { value.replies.splice(start, count); event("comment.removeReplies", start, count); },
      Delete: () => event("comment.delete"),
    };
    return value;
  }

  function conditions() {
    const value = {
      items: [],
      GetCount: () => value.items.length,
      GetItem: index => {
        event("conditions.getItem", index);
        return value.items[index - 1];
      },
      Delete: () => { value.items = []; event("conditions.delete"); },
    };
    for (const [method, kind] of [
      ["AddColorScale", "colorScale"],
      ["AddDatabar", "dataBar"],
      ["AddIconSetCondition", "iconSet"],
      ["AddTop10", "top10"],
      ["AddAboveAverage", "aboveAverage"],
      ["AddUniqueValues", "unique"],
    ]) {
      value[method] = (...args) => {
        const item = {
          GetFont: font,
          fillColor: null,
          GetType: () => kind,
          GetFillColor: () => item.fillColor,
          SetFillColor: color => { item.fillColor = color; event("condition.fill", color); },
          SetDupeUnique: type => event("condition.unique", type),
        };
        value.items.push(item);
        event(`conditions.${kind}`, ...args);
        return item;
      };
    }
    value.Add = (...args) => {
      const item = {
        fillColor: null,
        GetType: () => args[0],
        GetOperator: () => args[1],
        GetFormula1: () => args[2],
        GetFormula2: () => args[3],
        GetFillColor: () => item.fillColor,
        GetFont: font,
        SetFillColor: color => { item.fillColor = color; event("condition.fill", color); },
      };
      value.items.push(item);
      event("conditions.add", ...args);
      return item;
    };
    return value;
  }

  function validation() {
    const value = {
      type: null,
      alertStyle: null,
      operator: null,
      formula1: null,
      formula2: null,
      ignoreBlank: null,
      inCellDropdown: null,
      showInput: null,
      showError: null,
      inputTitle: null,
      inputMessage: null,
      errorTitle: null,
      errorMessage: null,
      GetType: () => value.type,
      GetAlertStyle: () => value.alertStyle,
      GetOperator: () => value.operator,
      GetFormula1: () => value.formula1,
      GetFormula2: () => value.formula2,
      GetIgnoreBlank: () => value.ignoreBlank,
      GetInCellDropdown: () => value.inCellDropdown,
      GetShowInput: () => value.showInput,
      GetShowError: () => value.showError,
      GetInputTitle: () => value.inputTitle,
      GetInputMessage: () => value.inputMessage,
      GetErrorTitle: () => value.errorTitle,
      GetErrorMessage: () => value.errorMessage,
    };
    for (const method of ["Add", "Modify"]) {
      value[method] = (type, alertStyle, operator, formula1, formula2) => {
        Object.assign(value, { type, alertStyle, operator, formula1, formula2 });
        event(`validation.${method}`, type, alertStyle, operator, formula1, formula2);
        return value;
      };
    }
    value.Delete = () => {
      Object.assign(value, {
        type: null,
        alertStyle: null,
        operator: null,
        formula1: null,
        formula2: null,
        inputTitle: null,
        inputMessage: null,
        errorTitle: null,
        errorMessage: null,
      });
      event("validation.Delete");
      return true;
    };
    for (const [method, field] of [
      ["SetIgnoreBlank", "ignoreBlank"],
      ["SetShowInput", "showInput"],
      ["SetShowError", "showError"],
      ["SetInputTitle", "inputTitle"],
      ["SetInputMessage", "inputMessage"],
      ["SetErrorTitle", "errorTitle"],
      ["SetErrorMessage", "errorMessage"],
    ]) {
      value[method] = next => {
        value[field] = next;
        event(`validation.${method}`, next);
        return true;
      };
    }
    return value;
  }

  function range(sheet, address) {
    const value = {
      sheet,
      address,
      values: [["Header", "Value"], ["A", 1], ["B", 2]],
      formula: null,
      arrayFormula: null,
      conditionSet: conditions(),
      validation: validation(),
      cellComment: null,
      format: {
        fontName: "Arial",
        fontSize: 11,
        bold: false,
        italic: false,
        underline: false,
        strikeout: false,
        fontColor: null,
        fillColor: null,
        horizontal: "left",
        vertical: "bottom",
        numberFormat: "General",
        wrap: false,
        orientation: 0,
      },
      GetAddress: () => value.address,
      GetValue: () => value.values,
      GetValue2: () => value.values,
      GetText: () => value.values,
      GetFormula: () => value.formula,
      GetFormulaArray: () => value.arrayFormula,
      GetNumberFormat: () => value.format.numberFormat,
      GetRowHeight: () => 15,
      GetColumnWidth: () => 10,
      GetHidden: () => false,
      GetWrapText: () => value.format.wrap,
      GetOrientation: () => value.format.orientation,
      GetFontName: () => value.format.fontName,
      GetFontSize: () => value.format.fontSize,
      GetBold: () => value.format.bold,
      GetItalic: () => value.format.italic,
      GetUnderline: () => value.format.underline,
      GetStrikeout: () => value.format.strikeout,
      GetFontColor: () => value.format.fontColor,
      GetFillColor: () => value.format.fillColor,
      GetAlignHorizontal: () => value.format.horizontal,
      GetAlignVertical: () => value.format.vertical,
      GetRowsCount: () => a1Shape(value.address).rows,
      GetColumnsCount: () => a1Shape(value.address).columns,
      SetValue: next => { value.values = next; event("range.value", address, next); return true; },
      SetFormula: next => { value.formula = next; event("range.formula", address, next); return true; },
      SetFormulaArray: next => { value.arrayFormula = next; event("range.arrayFormula", address, next); return true; },
      Replace: (search, replacement) => { event("range.replace", search, replacement); return true; },
      SetFontSize: next => { value.format.fontSize = next; event("range.fontSize", next); },
      SetFontName: next => { value.format.fontName = next; event("range.fontName", next); },
      SetBold: next => { value.format.bold = next; event("range.bold", next); },
      SetItalic: next => { value.format.italic = next; event("range.italic", next); },
      SetUnderline: next => { value.format.underline = next; event("range.underline", next); },
      SetStrikeout: next => { value.format.strikeout = next; event("range.strikeout", next); },
      SetFontColor: next => { value.format.fontColor = next; event("range.fontColor", next); },
      SetFillColor: next => { value.format.fillColor = next; event("range.fillColor", next); },
      SetAlignHorizontal: next => { value.format.horizontal = next; event("range.horizontal", next); },
      SetAlignVertical: next => { value.format.vertical = next; event("range.vertical", next); },
      SetNumberFormat: next => { value.format.numberFormat = next; event("range.numberFormat", next); },
      SetWrap: next => { value.format.wrap = next; event("range.wrap", next); },
      SetOrientation: next => { value.format.orientation = next; event("range.orientation", next); },
      SetColumnWidth: next => event("range.columnWidth", next),
      SetRowHeight: next => event("range.rowHeight", next),
      SetBorders: (...args) => event("range.border", ...args),
      Merge: next => event("range.merge", next),
      UnMerge: () => event("range.unmerge"),
      Insert: next => event("range.insert", next),
      Delete: next => event("range.delete", next),
      Copy: destination => event("range.copy", destination.address),
      Cut: destination => event("range.cut", destination.address),
      Clear: () => event("range.clear"),
      ClearContents: () => event("range.clearContents"),
      ClearFormats: () => event("range.clearFormats"),
      ClearHyperlinks: () => event("range.clearHyperlinks"),
      AutoFit: (...args) => event("range.autoFit", ...args),
      SetHidden: next => event("range.hidden", next),
      FillDown: () => event("range.fillDown"),
      FillUp: () => event("range.fillUp"),
      FillLeft: () => event("range.fillLeft"),
      FillRight: () => event("range.fillRight"),
      Select: () => event("range.select"),
      SetSort: (...args) => event("range.sort", ...args),
      SetAutoFilter: (...args) => event("range.filter", ...args),
      GetFormatConditions: () => value.conditionSet,
      GetValidation: () => value.validation,
      GetCharacters: (start, length) => ({
        SetText: next => event("characters.text", start, length, next),
        GetFont: font,
      }),
      AddComment: next => {
        value.cellComment = comment(next);
        sheet.comments.push(value.cellComment);
        event("range.comment", next);
        return value.cellComment;
      },
      GetComment: () => value.cellComment,
    };
    return value;
  }

  function drawing(sheet, classType = "shape") {
    const value = {
      sheet,
      name: "",
      width: 60 * 36000,
      height: 35 * 36000,
      rotation: 0,
      GetClassType: () => classType,
      GetName: () => value.name,
      GetWidth: () => value.width,
      GetHeight: () => value.height,
      GetRotation: () => value.rotation,
      GetFlipH: () => false,
      GetFlipV: () => false,
      SetName: next => { value.name = next; event("drawing.name", next); },
      SetSize: (width, height) => { value.width = width; value.height = height; event("drawing.size", width, height); },
      SetPosition: (...args) => event("drawing.position", ...args),
      SetRotation: next => { value.rotation = next; event("drawing.rotation", next); },
      SetFlipH: next => event("drawing.flipH", next),
      SetFlipV: next => event("drawing.flipV", next),
      Fill: next => event("drawing.fill", next),
      SetOutLine: next => event("drawing.line", next),
      GetDocContent: () => ({ GetElement: () => ({ AddText: next => event("drawing.text", next) }) }),
      Copy: () => drawing(sheet, classType),
      Delete: () => {
        const index = sheet.drawings.indexOf(value);
        if (index >= 0) sheet.drawings.splice(index, 1);
        event("drawing.delete");
      },
      ToJSON: () => JSON.stringify({ kind: classType, name: value.name }),
    };
    return value;
  }

  function freezePanes(sheet) {
    const value = {
      location: null,
      GetLocation: () => value.location,
      FreezeAt: target => {
        const lastCell = String(target.address).split(":").pop();
        apply(sheet.GetRange(`A1:${lastCell}`), "freeze.at", target.address);
      },
      FreezeRows: count => apply(sheet.GetRange(`A1:XFD${count}`), "freeze.rows", count),
      FreezeColumns: count => apply(sheet.GetRange(`A1:${columnName(count)}1048576`), "freeze.columns", count),
      Unfreeze: () => apply(null, "freeze.unfreeze"),
    };

    function columnName(number) {
      let remaining = number;
      let name = "";
      while (remaining > 0) {
        name = String.fromCharCode(65 + ((remaining - 1) % 26)) + name;
        remaining = Math.floor((remaining - 1) / 26);
      }
      return name;
    }

    function apply(location, eventName, eventValue) {
      event(eventName, eventValue);
      const update = () => {
        if (!state.freezeNeverApplies) value.location = location;
      };
      if (state.freezeDelayMs > 0) setTimeout(update, state.freezeDelayMs);
      else update();
    }

    return value;
  }

  function protectedRange(title, address) {
    const value = {
      title,
      address,
      anyoneType: "CanView",
      users: [],
      GetTitle: () => value.title,
      GetRange: () => value.address,
      GetAnyoneType: () => value.anyoneType,
      GetAllUsers: () => value.users,
      SetTitle: next => { value.title = next; event("protected.title", next); },
      SetRange: next => { value.address = next; event("protected.range", next); },
      SetAnyoneType: next => { value.anyoneType = next; event("protected.anyone", next); },
      AddUser: (id, name, type) => {
        value.users.push({ GetId: () => id, GetName: () => name, GetType: () => type });
        event("protected.user", id, name, type);
      },
      DeleteUser: id => {
        value.users = value.users.filter(user => user.GetId() !== id);
        event("protected.deleteUser", id);
      },
    };
    return value;
  }

  function table(sheet, address) {
    const value = {
      name: "Table1",
      style: "TableStyleMedium2",
      showTotals: false,
      showHeaders: true,
      rowStripes: true,
      columnStripes: false,
      firstColumn: false,
      lastColumn: false,
      showAutoFilter: true,
      showAutoFilterDropDown: true,
      summary: "",
      alternativeText: "",
      address,
      GetName: () => value.name,
      GetDisplayName: () => value.name,
      SetName: next => { value.name = next; event("table.name", next); },
      GetRange: () => sheet.GetRange(value.address),
      GetSourceType: () => "xlSrcRange",
      GetTableStyle: () => value.style,
      SetTableStyle: next => { value.style = next; event("table.style", next); },
      GetShowTotals: () => value.showTotals,
      SetShowTotals: next => { value.showTotals = next; event("table.totals", next); },
      GetShowHeaders: () => value.showHeaders,
      SetShowHeaders: next => { value.showHeaders = next; event("table.headers", next); },
      GetShowTableStyleRowStripes: () => value.rowStripes,
      SetShowTableStyleRowStripes: next => { value.rowStripes = next; event("table.rowStripes", next); },
      GetShowTableStyleColumnStripes: () => value.columnStripes,
      SetShowTableStyleColumnStripes: next => { value.columnStripes = next; event("table.columnStripes", next); },
      GetShowTableStyleFirstColumn: () => value.firstColumn,
      SetShowTableStyleFirstColumn: next => { value.firstColumn = next; },
      GetShowTableStyleLastColumn: () => value.lastColumn,
      SetShowTableStyleLastColumn: next => { value.lastColumn = next; },
      GetShowAutoFilter: () => value.showAutoFilter,
      SetShowAutoFilter: next => { value.showAutoFilter = next; },
      GetShowAutoFilterDropDown: () => value.showAutoFilterDropDown,
      SetShowAutoFilterDropDown: next => { value.showAutoFilterDropDown = next; },
      GetSummary: () => value.summary,
      SetSummary: next => { value.summary = next; },
      GetAlternativeText: () => value.alternativeText,
      SetAlternativeText: next => { value.alternativeText = next; },
      GetListColumns: () => [{}, {}],
      GetListRows: () => [{}, {}],
      Resize: next => { value.address = next; event("table.resize", next); },
      Delete: () => { sheet.tables = sheet.tables.filter(item => item !== value); event("table.delete"); },
      Unlist: () => { sheet.tables = sheet.tables.filter(item => item !== value); event("table.unlist"); },
    };
    return value;
  }

  function sheet(name) {
    const value = {
      name,
      visible: true,
      ranges: new Map(),
      drawings: [],
      charts: [],
      tables: [],
      comments: [],
      protectedRanges: [],
      page: {
        orientation: "xlLandscape",
        top: 20,
        right: 20,
        bottom: 20,
        left: 20,
        printGridlines: false,
        printHeadings: false,
        displayGridlines: true,
        displayHeadings: true,
      },
      GetName: () => value.name,
      SetName: next => { value.name = next; },
      GetIndex: () => state.sheets.indexOf(value),
      GetVisible: () => value.visible,
      SetVisible: next => { value.visible = next; event("sheet.visible", next); },
      SetActive: () => { state.activeSheet = value; event("sheet.active", name); },
      Move: before => {
        state.sheets.splice(state.sheets.indexOf(value), 1);
        state.sheets.splice(state.sheets.indexOf(before), 0, value);
        event("sheet.move", before.name);
      },
      GetRange(address) {
        const key = String(address);
        if (!value.ranges.has(key)) value.ranges.set(key, range(value, key));
        return value.ranges.get(key);
      },
      GetUsedRange: () => value.GetRange("A1:B3"),
      GetAllCharts: () => value.charts,
      GetAllDrawings: () => value.drawings,
      GetComments: () => value.comments,
      GetFreezePanes: () => value.freeze,
      GetAutoFilter: () => value.autoFilter,
      GetListObjects: () => value.tables,
      AddListObject: (_sourceType, address) => {
        const item = table(value, address);
        value.tables.push(item);
        event("table.create", address);
        return item;
      },
      FormatAsTable: address => { event("table.format", address); return true; },
      AddImage: () => { const item = drawing(value, "image"); value.drawings.push(item); event("drawing.image"); return item; },
      AddShape: () => { const item = drawing(value, "shape"); value.drawings.push(item); event("drawing.shape"); return item; },
      AddOleObject: () => { const item = drawing(value, "ole"); value.drawings.push(item); event("drawing.ole"); return item; },
      AddDrawing: item => { item.sheet = value; value.drawings.push(item); event("drawing.addCopy"); return item; },
      SetHyperlink: (...args) => event("sheet.hyperlink", ...args),
      AddProtectedRange: (title, address) => {
        const item = protectedRange(title, address);
        value.protectedRanges.push(item);
        event("protected.add", title, address);
        return item;
      },
      GetProtectedRange: title => value.protectedRanges.find(item => item.GetTitle() === title) || null,
      GetAllProtectedRanges: () => value.protectedRanges,
      GetPageOrientation: () => value.page.orientation,
      SetPageOrientation: next => { value.page.orientation = next; event("page.orientation", next); },
      GetTopMargin: () => value.page.top,
      GetRightMargin: () => value.page.right,
      GetBottomMargin: () => value.page.bottom,
      GetLeftMargin: () => value.page.left,
      SetTopMargin: next => { value.page.top = next; },
      SetRightMargin: next => { value.page.right = next; },
      SetBottomMargin: next => { value.page.bottom = next; },
      SetLeftMargin: next => { value.page.left = next; },
      GetPrintGridlines: () => value.page.printGridlines,
      SetPrintGridlines: next => { value.page.printGridlines = next; },
      GetPrintHeadings: () => value.page.printHeadings,
      SetPrintHeadings: next => { value.page.printHeadings = next; },
      SetDisplayGridlines: next => { value.page.displayGridlines = next; event("page.displayGridlines", next); },
      SetDisplayHeadings: next => { value.page.displayHeadings = next; event("page.displayHeadings", next); },
      Delete: () => { state.sheets.splice(state.sheets.indexOf(value), 1); return true; },
    };
    value.worksheet = {
      getSheetViewSettings: () => ({
        showGridLines: value.page.displayGridlines,
        showRowColHeaders: value.page.displayHeadings,
      }),
    };
    value.freeze = freezePanes(value);
    value.autoFilter = {
      GetRange: () => value.GetRange("A1:B3"),
      GetFilterMode: () => true,
      ShowAllData: () => event("filter.showAll"),
      ApplyFilter: () => event("filter.reapply"),
    };
    return value;
  }

  function definedName(name, refersTo) {
    const value = {
      name,
      refersTo,
      GetName: () => value.name,
      GetRefersTo: () => value.refersTo,
      GetRefersToRange: () => state.activeSheet.GetRange("A1:B3"),
      SetName: next => { value.name = next; event("name.rename", next); },
      SetRefersTo: next => { value.refersTo = next; event("name.refersTo", next); },
      Delete: () => { state.names.splice(state.names.indexOf(value), 1); event("name.delete"); },
    };
    return value;
  }

  function pivot(name = "Pivot1") {
    const value = {
      name,
      source: "Sheet1!A1:B3",
      style: "PivotStyleMedium9",
      GetName: () => value.name,
      SetName: next => { value.name = next; },
      GetTitle: () => value.name,
      GetDescription: () => "",
      GetSource: () => value.source,
      GetStyleName: () => value.style,
      GetTableRange2: () => state.activeSheet.GetRange("D1:F5"),
      GetRowFields: () => [],
      GetColumnFields: () => [],
      GetDataFields: () => [],
      AddFields: fields => event("pivot.fields", fields),
      AddDataField: field => { event("pivot.dataField", field); return { SetFunction: next => event("pivot.function", next) }; },
      SetStyleName: next => { value.style = next; event("pivot.style", next); },
      SetTitle: next => event("pivot.title", next),
      SetDescription: next => event("pivot.description", next),
      SetRowGrand: next => event("pivot.rowGrand", next),
      SetColumnGrand: next => event("pivot.columnGrand", next),
      RefreshTable: () => event("pivot.refresh"),
      ClearTable: () => event("pivot.clear"),
    };
    return value;
  }

  const sheet1 = sheet("Sheet1");
  const sheet2 = sheet("Sheet2");
  state.sheets.push(sheet1, sheet2);
  state.activeSheet = sheet1;
  for (const item of state.sheets) {
    item.worksheet.workbook = {
      oApi: {
        SetDocumentModified: value => event("workbook.document.modified", value),
      },
    };
  }

  const coreNames = [
    "Category", "ContentStatus", "Created", "Creator", "Description", "Identifier", "Keywords",
    "Language", "LastModifiedBy", "LastPrinted", "Modified", "Revision", "Subject", "Title", "Version",
  ];
  const core = {};
  for (const name of coreNames) {
    core[`Get${name}`] = () => state.core[name] ?? null;
    core[`Set${name}`] = next => { state.core[name] = next; event(`core.${name}`, next); };
  }
  const custom = {
    Add: (name, value) => { state.custom[name] = value; event("custom.add", name, value); },
    Get: name => state.custom[name],
  };

  const api = {
    GetSheets: () => state.sheets,
    GetActiveSheet: () => state.activeSheet,
    GetSelection: () => state.activeSheet.GetRange("A1"),
    GetRange: qualified => state.activeSheet.GetRange(String(qualified).split("!").at(-1)),
    CreateNewHistoryPoint: () => { state.historyPoints += 1; },
    AddSheet: name => {
      if (state.sheets.some(item => item.GetName() === name)) return null;
      const item = sheet(name);
      state.sheets.push(item);
      return item;
    },
    CreateColorFromRGB: (r, g, b) => ({ kind: "ApiColor", r, g, b }),
    CreateRGBColor: (r, g, b) => ({ kind: "ApiRGBColor", r, g, b }),
    CreateSolidFill: color => ({ type: "solid", color }),
    CreateNoFill: () => ({ type: "none" }),
    CreateStroke: (width, fill) => ({ width, fill }),
    AddDefName: (name, refersTo) => {
      const item = definedName(name, refersTo);
      state.names.push(item);
      event("name.add", name, refersTo);
      return item;
    },
    GetDefName: name => state.names.find(item => item.GetName() === name) || null,
    GetDefNames: () => state.names,
    RecalculateAllFormulas: () => event("recalculate.formulas"),
    RefreshAllPivots: () => event("recalculate.pivots"),
    GetAllPivotTables: () => state.pivots,
    InsertPivotNewWorksheet: () => {
      const item = pivot();
      state.pivots.push(item);
      event("pivot.createNew");
      return item;
    },
    InsertPivotExistingWorksheet: () => {
      const item = pivot();
      state.pivots.push(item);
      event("pivot.createExisting");
      return item;
    },
    GetPivotByName: name => state.pivots.find(item => item.GetName() === name) || null,
    GetCore: () => core,
    GetCustomProperties: () => custom,
  };

  return {
    bridge: loadSheetsBridge(api, state.macros, {
      ...options,
      onDocumentModified: value => event("document.modified", value),
    }),
    state,
    sheet1,
    sheet2,
  };
}

test("advanced Sheets bridge covers ranges, formulas, names, data rules, and tables", async () => {
  const { bridge, state } = createHarness();
  const result = await bridge.execute([
    { name: "sheets_inspect_range", arguments: { sheet: "Sheet1", range: "A1:B3" } },
    {
      name: "sheets_set_values",
      arguments: {
        sheet: "Sheet1",
        range: "C2:C5",
        values: [
          ["2026-07-23"],
          ["12:30:00"],
          [0.25],
          [true],
        ],
      },
    },
    { name: "sheets_set_array_formula", arguments: { sheet: "Sheet1", range: "D1:D3", formula: "{1;2;3}" } },
    {
      name: "sheets_format_range",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        strikeout: true,
        orientation: "xlUpward",
        columnWidthChars: 18,
        rowHeightPt: 24,
        borders: [{ side: "Bottom", style: "Thick", color: "#FF0000" }],
      },
    },
    { name: "sheets_manage_range", arguments: { sheet: "Sheet1", range: "A1:B3", action: "autoFit", rows: true, columns: true } },
    {
      name: "sheets_set_rich_text",
      arguments: { sheet: "Sheet1", range: "A1", text: "Revenue", runs: [{ start: 0, length: 3, bold: true, fontColor: "#0000FF" }] },
    },
    { name: "sheets_manage_names", arguments: { action: "add", name: "SalesData", refersTo: "Sheet1!$A$1:$B$3" } },
    { name: "sheets_inspect_names", arguments: {} },
    { name: "sheets_recalculate", arguments: { mode: "all" } },
    {
      name: "sheets_sort",
      arguments: { sheet: "Sheet1", range: "A1:B3", keys: [{ range: "B2:B3", order: "xlDescending" }], header: "xlYes" },
    },
    {
      name: "sheets_filter",
      arguments: { action: "set", sheet: "Sheet1", range: "A1:B3", field: 1, criteria1: ["A"], operator: "xlFilterValues" },
    },
    {
      name: "sheets_manage_table",
      arguments: { action: "create", sheet: "Sheet1", range: "A1:B3", name: "Sales", style: "TableStyleMedium4", showTotals: true },
    },
    {
      name: "sheets_manage_conditional_format",
      arguments: { action: "add", sheet: "Sheet1", range: "B2:B3", type: "duplicateValues", fillColor: "#FFFF00" },
    },
    {
      name: "sheets_manage_validation",
      arguments: {
        action: "add",
        sheet: "Sheet1",
        range: "A2:A3",
        type: "xlValidateList",
        formula1: "A,B",
        showInput: true,
        inputMessage: "Choose a value",
      },
    },
  ]);

  assert.equal(result.needsSave, true);
  assert.equal(state.historyPoints, 1);
  assert.equal(result.results[0].range.address, "A1:B3");
  assert.equal(result.results[7].names[0].name, "SalesData");
  const writtenValues = state.events.find(item => item[0] === "range.value" && item[1] === "C2:C5")[2];
  assert.deepEqual(Array.from(writtenValues, row => Array.from(row)), [
    ["2026-07-23"],
    ["12:30:00"],
    [0.25],
    [true],
  ]);
  assert.deepEqual(
    Array.from(result.results[1].readback, row => Array.from(row)),
    Array.from(writtenValues, row => Array.from(row)),
  );
  assert.ok(state.events.some(item => item[0] === "range.arrayFormula" && item[2] === "={1;2;3}"));
  assert.ok(state.events.some(item => item[0] === "range.columnWidth" && item[1] === 18));
  assert.ok(state.events.some(item => item[0] === "range.rowHeight" && item[1] === 24));
  assert.ok(state.events.some(item => item[0] === "range.border"));
  assert.ok(state.events.some(item => item[0] === "range.sort"));
  assert.ok(state.events.some(item => item[0] === "range.filter"));
  assert.ok(state.events.some(item => item[0] === "table.style" && item[1] === "TableStyleMedium4"));
  assert.ok(state.events.some(item => item[0] === "conditions.unique"));
  assert.ok(state.events.some(item => item[0] === "validation.SetInputMessage"));

  await assert.rejects(
    bridge.execute([{
      name: "sheets_format_range",
      arguments: {
        range: "A1",
        columnWidth: 18,
        columnWidthChars: 18,
      },
    }]),
    /columnWidthChars 与旧字段 columnWidth 不能同时提供/,
  );
});

test("Sheets bridge normalizes public enum aliases, preserves formula matrices, and exposes readback", async () => {
  const { bridge, state, sheet1 } = createHarness();
  const result = await bridge.execute([
    {
      name: "sheets_set_formula",
      arguments: {
        sheet: "Sheet1",
        range: "D1:E3",
        formula: [
          ["=A1+1", "=B1+1"],
          ["=A2+1", "=B2+1"],
          ["=A3+1", "=B3+1"],
        ],
      },
    },
    {
      name: "sheets_set_formula",
      arguments: { sheet: "Sheet1", range: "F1", formula: "=A1+1" },
    },
    {
      name: "sheets_format_range",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        fillColor: "#17365D",
        fontColor: "#FFFFFF",
        bold: true,
      },
    },
    {
      name: "sheets_manage_conditional_format",
      arguments: {
        action: "add",
        sheet: "Sheet1",
        range: "B2:B3",
        type: "cellValue",
        operator: "lessThan",
        formula1: 0,
        fillColor: "#FFF2CC",
      },
    },
    {
      name: "sheets_manage_validation",
      arguments: {
        action: "add",
        sheet: "Sheet1",
        range: "A2:A3",
        type: "list",
        alertStyle: "warning",
        operator: "equal",
        formula1: "A,B",
      },
    },
    {
      name: "sheets_sort",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        keys: [{ range: "B2:B3" }],
        header: true,
        orientation: "rows",
      },
    },
    {
      name: "sheets_filter",
      arguments: {
        action: "set",
        sheet: "Sheet1",
        range: "A1:B3",
        field: 1,
        criteria1: ">0",
        operator: "and",
      },
    },
    {
      name: "sheets_inspect_range",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        includeFormat: true,
        includeValidation: true,
      },
    },
    {
      name: "sheets_inspect_range",
      arguments: {
        sheet: "Sheet1",
        range: "B2:B3",
        includeConditionalFormats: true,
      },
    },
  ]);

  assert.deepEqual(
    sheet1.GetRange("D1:E3").formula,
    [
      ["=A1+1", "=B1+1"],
      ["=A2+1", "=B2+1"],
      ["=A3+1", "=B3+1"],
    ],
  );
  assert.deepEqual(result.results[0].inputShape, { rows: 3, columns: 2 });
  assert.deepEqual(result.results[0].targetShape, { rows: 3, columns: 2 });
  assert.deepEqual(
    Array.from(result.results[0].readback, row => Array.from(row)),
    [
      ["=A1+1", "=B1+1"],
      ["=A2+1", "=B2+1"],
      ["=A3+1", "=B3+1"],
    ],
  );
  assert.deepEqual(result.results[1].inputShape, { rows: 1, columns: 1 });
  assert.deepEqual(result.results[1].targetShape, { rows: 1, columns: 1 });
  assert.equal(result.results[2].readback.bold, true);
  assert.deepEqual(result.results[2].readback.fillColor, { r: 23, g: 54, b: 93 });
  assert.equal(result.results[3].rule.operator, "xlLess");
  assert.deepEqual(result.results[3].rule.fillColor, { r: 255, g: 242, b: 204 });
  assert.equal(result.results[4].rule.type, "xlValidateList");
  assert.equal(result.results[4].rule.alertStyle, "xlValidAlertWarning");
  assert.equal(result.results[4].rule.operator, "xlEqual");
  assert.equal(result.results[4].rule.formula1, "A,B");
  assert.ok(state.events.some(item => item[0] === "conditions.add" && item[2] === "xlLess"));
  assert.ok(state.events.some(item => item[0] === "range.fillColor" && item[1].kind === "ApiColor"));
  assert.ok(state.events.some(item => item[0] === "condition.fill" && item[1].kind === "ApiColor"));
  assert.ok(state.events.some(item => item[0] === "validation.Add"
    && item[1] === "xlValidateList"
    && item[2] === "xlValidAlertWarning"
    && item[3] === "xlEqual"));
  assert.ok(state.events.some(item => item[0] === "range.sort"
    && item.at(-2) === "xlYes"
    && item.at(-1) === "xlSortRows"));
  assert.ok(state.events.some(item => item[0] === "range.filter" && item[3] === "xlAnd"));
  assert.deepEqual(result.results[7].range.format.fillColor, { r: 23, g: 54, b: 93 });
  assert.equal(result.results[8].range.conditionalFormats.rules[0].operator, "xlLess");

  await assert.rejects(
    bridge.execute([{
      name: "sheets_set_formula",
      arguments: { range: "H1:I3", formula: [["=1"]] },
    }]),
    /formula 尺寸 1x1 与目标区域 3x2 不一致/,
  );
  await assert.rejects(
    bridge.execute([{
      name: "sheets_manage_conditional_format",
      arguments: { action: "add", range: "A1", type: "cellValue", operator: "smallerMaybe" },
    }]),
    /条件格式运算符 不支持/,
  );
});

test("Sheets bridge does not report rejected or empty mutations as changed", async () => {
  const { bridge, sheet1 } = createHarness();

  await assert.rejects(
    bridge.execute([{
      name: "sheets_format_range",
      arguments: { sheet: "Sheet1", range: "A1" },
    }]),
    /至少需要一个格式属性/,
  );

  sheet1.GetRange("B1").SetFillColor = () => false;
  await assert.rejects(
    bridge.execute([{
      name: "sheets_format_range",
      arguments: { sheet: "Sheet1", range: "B1", fillColor: "#17365D" },
    }]),
    /ONLYOFFICE 拒绝设置单元格填充/,
  );

  sheet1.GetRange("C1").SetValue = () => false;
  await assert.rejects(
    bridge.execute([{
      name: "sheets_set_values",
      arguments: { sheet: "Sheet1", range: "C1", values: 1 },
    }]),
    /ONLYOFFICE 拒绝写入单元格值/,
  );

  sheet1.GetRange("D1:D2").SetFormulaArray = () => false;
  await assert.rejects(
    bridge.execute([{
      name: "sheets_set_array_formula",
      arguments: { sheet: "Sheet1", range: "D1:D2", formula: "=ROW()" },
    }]),
    /ONLYOFFICE 拒绝设置数组公式/,
  );

  sheet1.GetRange("E1:E2").conditionSet.AddUniqueValues = () => null;
  await assert.rejects(
    bridge.execute([{
      name: "sheets_manage_conditional_format",
      arguments: {
        action: "add",
        sheet: "Sheet1",
        range: "E1:E2",
        type: "duplicateValues",
      },
    }]),
    /ONLYOFFICE 未创建添加重复值条件格式/,
  );

  sheet1.SetDisplayGridlines = () => false;
  await assert.rejects(
    bridge.execute([{
      name: "sheets_manage_page_layout",
      arguments: { sheet: "Sheet1", displayGridlines: false },
    }]),
    /ONLYOFFICE 拒绝设置屏幕网格线/,
  );
});

test("Sheets bridge rejects invalid value and formula shapes before mutation", async () => {
  const invalidCalls = [
    { name: "sheets_set_values", arguments: { range: "A1:B2", values: 1 } },
    { name: "sheets_set_values", arguments: { range: "A1:B2", values: [1, 2] } },
    { name: "sheets_set_values", arguments: { range: "A1:B2", values: [] } },
    { name: "sheets_set_values", arguments: { range: "A1:B2", values: [[1, 2], [3]] } },
    { name: "sheets_set_values", arguments: { range: "A1:B2", values: [[1], [2]] } },
    { name: "sheets_set_formula", arguments: { range: "A1:B2", formula: "=A1" } },
    { name: "sheets_set_formula", arguments: { range: "A1:B2", formula: [["=1"], ["=2"]] } },
  ];

  for (const call of invalidCalls) {
    const { bridge, state } = createHarness();
    let caught;
    try {
      await bridge.execute([call]);
    } catch (error) {
      caught = error;
    }
    assert.equal(caught && caught.code, "INVALID_TOOL_ARGUMENTS");
    assert.equal(caught.details.completedToolCalls, 0);
    assert.equal(caught.details.partialMutationPossible, false);
    assert.equal(
      state.events.some(item => item[0] === "range.value" || item[0] === "range.formula"),
      false,
    );
  }
});

test("Sheets bridge reports partial batch mutation without retrying failed calls", async () => {
  const { bridge, state } = createHarness();
  let caught;
  try {
    await bridge.execute([
      {
        name: "sheets_set_values",
        arguments: { sheet: "Sheet1", range: "A1", values: "committed" },
      },
      {
        name: "sheets_set_values",
        arguments: { sheet: "Sheet1", range: "B1:B2", values: "invalid" },
      },
    ]);
  } catch (error) {
    caught = error;
  }

  assert.equal(caught && caught.code, "INVALID_TOOL_ARGUMENTS");
  assert.equal(caught.details.phase, "sheets-command");
  assert.equal(caught.details.tool, "sheets_set_values");
  assert.equal(caught.details.toolCallIndex, 1);
  assert.equal(caught.details.completedToolCalls, 1);
  assert.equal(caught.details.partialMutationPossible, true);
  assert.equal(state.events.filter(item => item[0] === "range.value").length, 1);
});

test("Sheets bridge reads validation getters and rejects null validation mutations", async () => {
  const { bridge, sheet1 } = createHarness();
  const added = await bridge.execute([{
    name: "sheets_manage_validation",
    arguments: {
      action: "add",
      sheet: "Sheet1",
      range: "A2:A3",
      type: "list",
      formula1: "Red,Amber,Green",
      ignoreBlank: true,
      showInput: true,
      showError: true,
      inputTitle: "Status",
      inputMessage: "Choose one",
      errorTitle: "Invalid",
      errorMessage: "Use the list",
    },
  }]);
  const inspected = await bridge.execute([{
    name: "sheets_inspect_range",
    arguments: {
      sheet: "Sheet1",
      range: "A2:A3",
      includeValidation: true,
    },
  }]);

  assert.equal(added.results[0].validation.type, "xlValidateList");
  assert.equal(added.results[0].validation.formula1, "Red,Amber,Green");
  assert.equal(added.results[0].validation.ignoreBlank, true);
  assert.equal(added.results[0].validation.inputTitle, "Status");
  assert.equal(added.results[0].validation.errorMessage, "Use the list");
  assert.deepEqual(
    inspected.results[0].range.validation,
    added.results[0].validation,
  );

  sheet1.GetRange("B2:B3").validation.Add = () => null;
  let caught;
  try {
    await bridge.execute([{
      name: "sheets_manage_validation",
      arguments: {
        action: "add",
        sheet: "Sheet1",
        range: "B2:B3",
        type: "list",
        formula1: "A,B",
      },
    }]);
  } catch (error) {
    caught = error;
  }
  assert.equal(caught && caught.code, "EXECUTION_FAILED");
  assert.equal(caught.details.completedToolCalls, 0);
  assert.equal(caught.details.partialMutationPossible, true);
});

test("Sheets bridge uses one-based conditional format GetItem indices", async () => {
  const { bridge, state } = createHarness();
  const result = await bridge.execute([{
    name: "sheets_manage_conditional_format",
    arguments: {
      action: "add",
      sheet: "Sheet1",
      range: "A1:A2",
      type: "duplicateValues",
    },
  }]);

  assert.equal(result.results[0].rule.index, 0);
  assert.equal(result.results[0].conditionalFormats.rules[0].index, 0);
  assert.ok(state.events.some(item => item[0] === "conditions.getItem" && item[1] === 1));
});

test("Sheets bridge rejects null from every conditional-format Add API", async () => {
  const cases = [
    ["Add", "cellValue"],
    ["AddColorScale", "colorScale"],
    ["AddDatabar", "dataBar"],
    ["AddIconSetCondition", "iconSet"],
    ["AddTop10", "top10"],
    ["AddAboveAverage", "aboveAverage"],
    ["AddUniqueValues", "duplicateValues"],
  ];

  for (const [method, type] of cases) {
    const { bridge, sheet1 } = createHarness();
    const range = sheet1.GetRange("A1:A2");
    range.conditionSet[method] = () => null;
    let caught;
    try {
      await bridge.execute([{
        name: "sheets_manage_conditional_format",
        arguments: {
          action: "add",
          sheet: "Sheet1",
          range: "A1:A2",
          type,
          operator: "lessThan",
          formula1: 0,
        },
      }]);
    } catch (error) {
      caught = error;
    }
    assert.equal(caught && caught.code, "EXECUTION_FAILED", method);
    assert.equal(caught.details.completedToolCalls, 0, method);
  }
});

test("Sheets bridge verifies conditional-format deleteAll empties the collection", async () => {
  const { bridge, sheet1 } = createHarness();
  const range = sheet1.GetRange("A1:A2");
  range.conditionSet.items.push({ GetType: () => "xlCellValue" });
  range.conditionSet.Delete = () => true;

  await assert.rejects(
    bridge.execute([{
      name: "sheets_manage_conditional_format",
      arguments: {
        action: "deleteAll",
        sheet: "Sheet1",
        range: "A1:A2",
      },
    }]),
    /删除条件格式后规则数量仍为 1/,
  );
});

test("Sheets bridge rejects null from validation Add and Modify", async () => {
  for (const action of ["add", "modify"]) {
    const { bridge, sheet1 } = createHarness();
    sheet1.GetRange("A1:A2").validation[
      action === "add" ? "Add" : "Modify"
    ] = () => null;
    let caught;
    try {
      await bridge.execute([{
        name: "sheets_manage_validation",
        arguments: {
          action,
          sheet: "Sheet1",
          range: "A1:A2",
          type: "list",
          formula1: "A,B",
        },
      }]);
    } catch (error) {
      caught = error;
    }
    assert.equal(caught && caught.code, "EXECUTION_FAILED", action);
    assert.equal(caught.details.completedToolCalls, 0, action);
  }
});

test("Sheets bridge safely degrades table creation in community editions", async () => {
  const { bridge, state, sheet1 } = createHarness();
  sheet1.AddListObject = () => null;

  const basic = await bridge.execute([{
    name: "sheets_manage_table",
    arguments: {
      action: "create",
      sheet: "Sheet1",
      range: "A1:B3",
    },
  }]);
  assert.equal(basic.results[0].tableKind, "basic");
  assert.equal(basic.results[0].degraded, true);
  assert.equal(basic.results[0].formatted, true);
  assert.equal(basic.results[0].table, null);
  assert.equal(basic.results[0].range.address, "A1:B3");

  let structuredError;
  try {
    await bridge.execute([{
      name: "sheets_manage_table",
      arguments: {
        action: "create",
        tableMode: "structured",
        sheet: "Sheet1",
        range: "A1:B3",
      },
    }]);
  } catch (error) {
    structuredError = error;
  }
  assert.equal(structuredError && structuredError.code, "SHEETS_API_UNSUPPORTED");

  let autoStructuredError;
  try {
    await bridge.execute([{
      name: "sheets_manage_table",
      arguments: {
        action: "create",
        sheet: "Sheet1",
        range: "A1:B3",
        name: "MustNotBeDropped",
      },
    }]);
  } catch (error) {
    autoStructuredError = error;
  }
  assert.equal(autoStructuredError && autoStructuredError.code, "SHEETS_API_UNSUPPORTED");

  await assert.rejects(
    bridge.execute([{
      name: "sheets_manage_table",
      arguments: {
        action: "create",
        tableMode: "basic",
        sheet: "Sheet1",
        range: "A1:B3",
        name: "NotSupported",
      },
    }]),
    /基础格式表格不支持结构化属性/,
  );
  assert.equal(state.events.filter(item => item[0] === "table.format").length, 1);
});

test("structured table inspection and lifecycle use ApiListObject", async () => {
  const { bridge, state } = createHarness();

  await bridge.execute([{
    name: "sheets_manage_table",
    arguments: {
      action: "create",
      sheet: "Sheet1",
      range: "A1:B3",
      name: "Sales",
      style: "TableStyleMedium4",
    },
  }]);

  const inspected = await bridge.execute([{
    name: "sheets_inspect_tables",
    arguments: { sheet: "Sheet1", name: "Sales" },
  }]);
  assert.equal(inspected.needsSave, false);
  assert.equal(inspected.results[0].sheets[0].tables[0].name, "Sales");
  assert.equal(inspected.results[0].sheets[0].tables[0].columnCount, 2);

  await bridge.execute([{
    name: "sheets_manage_table",
    arguments: {
      action: "update",
      sheet: "Sheet1",
      tableName: "Sales",
      newName: "Revenue",
      style: "TableStyleLight9",
      showTotals: true,
      firstColumn: true,
      summary: "Revenue data",
    },
  }]);
  await bridge.execute([{
    name: "sheets_manage_table",
    arguments: { action: "resize", sheet: "Sheet1", tableName: "Revenue", range: "A1:C5" },
  }]);

  const resized = await bridge.execute([{
    name: "sheets_inspect_tables",
    arguments: { sheet: "Sheet1", tableIndex: 0 },
  }]);
  assert.equal(resized.results[0].sheets[0].tables[0].range.address, "A1:C5");
  assert.equal(resized.results[0].sheets[0].tables[0].style, "TableStyleLight9");
  assert.equal(resized.results[0].sheets[0].tables[0].summary, "Revenue data");

  await bridge.execute([{
    name: "sheets_manage_table",
    arguments: { action: "unlist", sheet: "Sheet1", tableName: "Revenue" },
  }]);
  assert.equal(state.sheets[0].tables.length, 0);
  assert.ok(state.events.some(item => item[0] === "table.resize" && item[1] === "A1:C5"));
  assert.ok(state.events.some(item => item[0] === "table.unlist"));
});

test("advanced Sheets bridge covers pivots, drawings, comments, freeze panes, metadata, protection, and page layout", async () => {
  const { bridge, state } = createHarness();
  const image = {
    url: "https://app.test/copilot-api/images/asset.png?token=signed",
    widthPx: 800,
    heightPx: 600,
  };
  const result = await bridge.execute([
    {
      name: "sheets_manage_pivot",
      arguments: {
        action: "createExisting",
        sourceSheet: "Sheet1",
        sourceRange: "A1:B3",
        destinationSheet: "Sheet2",
        destinationRange: "D1",
        name: "SalesPivot",
        fields: { rows: "Header" },
        dataFields: [{ name: "Value", function: "Sum" }],
      },
    },
    { name: "sheets_inspect_pivots", arguments: {} },
    {
      name: "sheets_manage_drawing",
      arguments: { action: "addImage", sheet: "Sheet1", _image: image, name: "Logo", widthMm: 50, heightMm: 30 },
    },
    { name: "sheets_inspect_drawings", arguments: { sheet: "Sheet1" } },
    { name: "sheets_manage_hyperlink", arguments: { action: "set", sheet: "Sheet1", range: "A1", url: "https://example.com" } },
    { name: "sheets_manage_comments", arguments: { action: "add", sheet: "Sheet1", range: "A1", text: "Review" } },
    { name: "sheets_inspect_comments", arguments: { sheet: "Sheet1" } },
    { name: "sheets_manage_freeze_panes", arguments: { action: "freezeAt", sheet: "Sheet1", range: "B2" } },
    { name: "sheets_inspect_freeze_panes", arguments: { sheet: "Sheet1" } },
    {
      name: "sheets_manage_properties",
      arguments: { core: { title: "Quarterly report", creator: "AI" }, custom: { Department: "Finance" } },
    },
    { name: "sheets_inspect_properties", arguments: { customNames: ["Department"] } },
    {
      name: "sheets_manage_protected_ranges",
      arguments: { action: "add", sheet: "Sheet1", title: "Inputs", range: "A1:B3" },
    },
    {
      name: "sheets_manage_protected_ranges",
      arguments: { action: "addUser", sheet: "Sheet1", title: "Inputs", userId: "u1", userName: "Alex", permission: "CanEdit" },
    },
    { name: "sheets_inspect_protected_ranges", arguments: { sheet: "Sheet1" } },
    {
      name: "sheets_manage_page_layout",
      arguments: {
        sheet: "Sheet1",
        orientation: "portrait",
        marginsPt: { top: 10, bottom: 10 },
        printGridlines: true,
        displayGridlines: false,
        displayHeadings: false,
      },
    },
    { name: "sheets_inspect_page_layout", arguments: { sheet: "Sheet1" } },
  ]);

  assert.equal(result.needsSave, true);
  assert.equal(state.historyPoints, 1);
  assert.equal(result.results[1].pivots[0].name, "SalesPivot");
  assert.equal(result.results[3].drawingCount, 1);
  assert.equal(result.results[6].sheets[0].comments[0].text, "Review");
  assert.equal(result.results[7].location.address, "A1:B2");
  assert.equal(result.results[7].frozenRows, 2);
  assert.equal(result.results[7].frozenColumns, 2);
  assert.equal(result.results[7].topLeftCell, "C3");
  assert.equal(result.results[7].verified, true);
  assert.equal(result.results[8].location.address, "A1:B2");
  assert.equal(result.results[10].core.title, "Quarterly report");
  assert.equal(result.results[10].custom.Department, "Finance");
  assert.equal(result.results[13].sheets[0].protectedRanges[0].users[0].name, "Alex");
  assert.equal(result.results[15].orientation, "xlPortrait");
  assert.equal(result.results[14].printGridlines, true);
  assert.equal(result.results[14].displayGridlines, false);
  assert.equal(result.results[14].verified, true);
  assert.equal(result.results[15].displayGridlines, false);
  assert.equal(result.results[15].displayHeadings, false);
  assert.equal(result.results[14].marginsPt.top, 10);
  assert.equal(result.results[14].marginsPt.bottom, 10);
  assert.equal(result.results[15].marginsPt.top, 10);
  assert.ok(state.events.some(item => item[0] === "drawing.image"));
  assert.ok(state.events.some(item => item[0] === "pivot.function" && item[1] === "Sum"));
  assert.ok(state.events.some(item => item[0] === "page.displayGridlines" && item[1] === false));

  await assert.rejects(
    bridge.execute([{
      name: "sheets_manage_page_layout",
      arguments: { margins: { top: 10 }, marginsPt: { top: 10 } },
    }]),
    /marginsPt 与旧字段 margins 不能同时提供/,
  );
  await assert.rejects(
    bridge.execute([{ name: "sheets_manage_page_layout", arguments: {} }]),
    /至少需要一个页面布局属性/,
  );
});

test("Sheets bridge waits for asynchronous freeze panes before reporting success", async () => {
  const { bridge, state } = createHarness({
    freezeDelayMs: 20,
    verifyIntervalMs: 2,
    verifyTimeoutMs: 250,
  });

  const result = await bridge.execute([
    {
      name: "sheets_manage_freeze_panes",
      arguments: { action: "freezeRows", sheet: "Sheet1", count: 2 },
    },
    {
      name: "sheets_inspect_freeze_panes",
      arguments: { sheet: "Sheet1" },
    },
  ]);

  assert.equal(result.results[0].location.address, "A1:XFD2");
  assert.equal(result.results[0].frozenRows, 2);
  assert.equal(result.results[0].frozenColumns, 0);
  assert.equal(result.results[0].topLeftCell, "A3");
  assert.equal(result.results[0].verified, true);
  assert.equal(result.results[1].location.address, "A1:XFD2");
  assert.equal(result.results[1].frozenRows, 2);
  assert.equal(result.results[1].topLeftCell, "A3");
  assert.equal(result.results[1].verified, true);
  assert.ok(state.events.some(item => (
    item[0] === "document.modified" && item[1] === true
  )));
});

test("Sheets bridge marks verified view changes through the workbook model fallback", async () => {
  const { bridge, state } = createHarness({
    withoutAscEditor: true,
    verifyIntervalMs: 2,
    verifyTimeoutMs: 100,
  });

  const result = await bridge.execute([{
    name: "sheets_manage_page_layout",
    arguments: { sheet: "Sheet1", displayGridlines: false },
  }]);

  assert.equal(result.results[0].verified, true);
  assert.equal(result.results[0].displayGridlines, false);
  assert.ok(state.events.some(item => (
    item[0] === "workbook.document.modified" && item[1] === true
  )));
});

test("Sheets bridge rejects unverified worksheet view mutations", async () => {
  const frozen = createHarness({
    freezeNeverApplies: true,
    verifyIntervalMs: 2,
    verifyTimeoutMs: 20,
  });
  await assert.rejects(
    frozen.bridge.execute([{
      name: "sheets_manage_freeze_panes",
      arguments: { action: "freezeRows", sheet: "Sheet1", count: 2 },
    }]),
    error => (
      error.code === "SHEETS_VIEW_STATE_NOT_APPLIED"
      && error.details.phase === "sheets-view-verification"
      && error.details.partialMutationPossible === true
    ),
  );

  const gridlines = createHarness({
    verifyIntervalMs: 2,
    verifyTimeoutMs: 20,
  });
  gridlines.sheet1.SetDisplayGridlines = () => true;
  await assert.rejects(
    gridlines.bridge.execute([{
      name: "sheets_manage_page_layout",
      arguments: { sheet: "Sheet1", displayGridlines: false },
    }]),
    error => (
      error.code === "SHEETS_VIEW_STATE_NOT_APPLIED"
      && error.details.expected.displayGridlines === false
      && error.details.observed.displayGridlines === true
    ),
  );
});

test("Sheets bridge requires explicit positive freeze counts and unfreeze action", async () => {
  const { bridge } = createHarness();
  for (const argumentsValue of [
    { action: "freezeRows", count: 0 },
    { action: "freezeRows", count: 1.5 },
    { action: "freezeColumns", count: -1 },
    { action: "freezeAt" },
    { action: "unfreeze", count: 0 },
  ]) {
    await assert.rejects(
      bridge.execute([{
        name: "sheets_manage_freeze_panes",
        arguments: argumentsValue,
      }]),
      error => error.code === "INVALID_TOOL_ARGUMENTS" || /freezeAt 需要 range/.test(error.message),
    );
  }
});

test("Sheets bridge unfreezes and restores screen gridlines without changing print gridlines", async () => {
  const { bridge } = createHarness({
    verifyIntervalMs: 2,
    verifyTimeoutMs: 100,
  });
  await bridge.execute([
    {
      name: "sheets_manage_freeze_panes",
      arguments: { action: "freezeRows", sheet: "Sheet1", count: 2 },
    },
    {
      name: "sheets_manage_page_layout",
      arguments: {
        sheet: "Sheet1",
        displayGridlines: false,
        printGridlines: true,
      },
    },
  ]);

  const result = await bridge.execute([
    {
      name: "sheets_manage_freeze_panes",
      arguments: { action: "unfreeze", sheet: "Sheet1" },
    },
    {
      name: "sheets_inspect_freeze_panes",
      arguments: { sheet: "Sheet1" },
    },
    {
      name: "sheets_manage_page_layout",
      arguments: { sheet: "Sheet1", displayGridlines: true },
    },
    {
      name: "sheets_inspect_page_layout",
      arguments: { sheet: "Sheet1" },
    },
  ]);

  for (const item of result.results.slice(0, 2)) {
    assert.equal(item.location, null);
    assert.equal(item.frozenRows, 0);
    assert.equal(item.frozenColumns, 0);
    assert.equal(item.topLeftCell, null);
    assert.equal(item.verified, true);
  }
  for (const item of result.results.slice(2)) {
    assert.equal(item.displayGridlines, true);
    assert.equal(item.printGridlines, true);
    assert.equal(item.verified, true);
  }
});

test("Sheets macro tools use the documented plugin methods and reject mixed batches", async () => {
  const { bridge, state } = createHarness();
  const inspected = await bridge.execute([{ name: "sheets_inspect_macros", arguments: {} }]);
  assert.equal(inspected.needsSave, false);
  assert.equal(inspected.results[0].content.macrosArray[0].name, "Demo");

  const content = { macrosArray: [{ name: "Updated", value: "return 2;", guid: "macro-2" }], current: 0 };
  const updated = await bridge.execute([{ name: "sheets_set_macros", arguments: { content } }]);
  assert.equal(updated.needsSave, true);
  assert.equal(state.macros.content.macrosArray[0].name, "Updated");
  assert.deepEqual(state.macros.calls.map(item => item[0]), ["GetMacros", "SetMacros"]);

  await assert.rejects(
    bridge.execute([
      { name: "sheets_inspect_macros", arguments: {} },
      { name: "sheets_inspect", arguments: {} },
    ]),
    /必须单独调用/,
  );
});
