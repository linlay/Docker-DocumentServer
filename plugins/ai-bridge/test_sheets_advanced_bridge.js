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

function loadSheetsBridge(api, macroState) {
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
  const window = { Asc: asc };
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
  };
  vm.runInNewContext(sheetsSource, sandbox);
  return window.AICopilotBridges.cell;
}

function createHarness() {
  const state = {
    events: [],
    sheets: [],
    names: [],
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
          SetFillColor: color => event("condition.fill", color),
          SetDupeUnique: type => event("condition.unique", type),
        };
        value.items.push(item);
        event(`conditions.${kind}`, ...args);
        return item;
      };
    }
    value.Add = (...args) => {
      const item = { GetFont: font, SetFillColor: color => event("condition.fill", color) };
      value.items.push(item);
      event("conditions.add", ...args);
      return item;
    };
    return value;
  }

  function validation() {
    const value = {};
    for (const method of [
      "Add", "Modify", "Delete", "SetIgnoreBlank", "SetShowInput", "SetShowError",
      "SetInputTitle", "SetInputMessage", "SetErrorTitle", "SetErrorMessage",
    ]) {
      value[method] = (...args) => event(`validation.${method}`, ...args);
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
      GetAddress: () => value.address,
      GetValue: () => value.values,
      GetValue2: () => value.values,
      GetText: () => value.values,
      GetFormula: () => value.formula,
      GetFormulaArray: () => value.arrayFormula,
      GetNumberFormat: () => "General",
      GetRowHeight: () => 15,
      GetColumnWidth: () => 10,
      GetHidden: () => false,
      GetWrapText: () => false,
      GetOrientation: () => 0,
      GetRowsCount: () => value.values.length,
      GetColumnsCount: () => value.values[0].length,
      SetValue: next => { value.values = next; event("range.value", address, next); return true; },
      SetFormula: next => { value.formula = next; event("range.formula", address, next); return true; },
      SetFormulaArray: next => { value.arrayFormula = next; event("range.arrayFormula", address, next); return true; },
      Replace: (search, replacement) => { event("range.replace", search, replacement); return true; },
      SetFontSize: next => event("range.fontSize", next),
      SetFontName: next => event("range.fontName", next),
      SetBold: next => event("range.bold", next),
      SetItalic: next => event("range.italic", next),
      SetUnderline: next => event("range.underline", next),
      SetStrikeout: next => event("range.strikeout", next),
      SetFontColor: next => event("range.fontColor", next),
      SetFillColor: next => event("range.fillColor", next),
      SetAlignHorizontal: next => event("range.horizontal", next),
      SetAlignVertical: next => event("range.vertical", next),
      SetNumberFormat: next => event("range.numberFormat", next),
      SetWrap: next => event("range.wrap", next),
      SetOrientation: next => event("range.orientation", next),
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
      FreezeAt: target => { value.location = target; event("freeze.at", target.address); },
      FreezeRows: count => { value.location = sheet.GetRange(`A${count + 1}`); event("freeze.rows", count); },
      FreezeColumns: count => { value.location = sheet.GetRange(`${String.fromCharCode(65 + count)}1`); event("freeze.columns", count); },
      Unfreeze: () => { value.location = null; event("freeze.unfreeze"); },
    };
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
      page: { orientation: "xlLandscape", top: 20, right: 20, bottom: 20, left: 20, gridlines: false, headings: false },
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
      FormatAsTable: address => { event("table.format", address); return table(value, address); },
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
      GetPrintGridlines: () => value.page.gridlines,
      SetPrintGridlines: next => { value.page.gridlines = next; },
      GetPrintHeadings: () => value.page.headings,
      SetPrintHeadings: next => { value.page.headings = next; },
      Delete: () => { state.sheets.splice(state.sheets.indexOf(value), 1); return true; },
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
    CreateColorFromRGB: (r, g, b) => ({ r, g, b }),
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

  return { bridge: loadSheetsBridge(api, state.macros), state, sheet1, sheet2 };
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
          [{ type: "date", value: "2026-07-23" }],
          [{ type: "time", value: "12:30:00" }],
          [{ type: "percent", value: "25%" }],
          [{ type: "boolean", value: "true" }],
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
  const typedValues = state.events.find(item => item[0] === "range.value" && item[1] === "C2:C5")[2];
  assert.ok(typedValues[0][0] instanceof Date);
  assert.equal(typedValues[1][0], 12.5 / 24);
  assert.equal(typedValues[2][0], 0.25);
  assert.equal(typedValues[3][0], true);
  assert.ok(state.events.some(item => item[0] === "range.arrayFormula" && item[2] === "={1;2;3}"));
  assert.ok(state.events.some(item => item[0] === "range.border"));
  assert.ok(state.events.some(item => item[0] === "range.sort"));
  assert.ok(state.events.some(item => item[0] === "range.filter"));
  assert.ok(state.events.some(item => item[0] === "table.style" && item[1] === "TableStyleMedium4"));
  assert.ok(state.events.some(item => item[0] === "conditions.unique"));
  assert.ok(state.events.some(item => item[0] === "validation.SetInputMessage"));
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
      arguments: { sheet: "Sheet1", orientation: "xlPortrait", margins: { top: 10, bottom: 10 }, printGridlines: true },
    },
    { name: "sheets_inspect_page_layout", arguments: { sheet: "Sheet1" } },
  ]);

  assert.equal(result.needsSave, true);
  assert.equal(state.historyPoints, 1);
  assert.equal(result.results[1].pivots[0].name, "SalesPivot");
  assert.equal(result.results[3].drawingCount, 1);
  assert.equal(result.results[6].sheets[0].comments[0].text, "Review");
  assert.equal(result.results[8].location.address, "B2");
  assert.equal(result.results[10].core.title, "Quarterly report");
  assert.equal(result.results[10].custom.Department, "Finance");
  assert.equal(result.results[13].sheets[0].protectedRanges[0].users[0].name, "Alex");
  assert.equal(result.results[15].orientation, "xlPortrait");
  assert.ok(state.events.some(item => item[0] === "drawing.image"));
  assert.ok(state.events.some(item => item[0] === "pivot.function" && item[1] === "Sum"));

  await assert.rejects(
    bridge.execute([{ name: "sheets_manage_page_layout", arguments: {} }]),
    /至少需要一个页面布局属性/,
  );
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
