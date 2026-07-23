"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const bridgeRoot = process.env.AI_BRIDGE_DIR || __dirname;
const slidesSource = fs.readFileSync(path.join(bridgeRoot, "bridges/slides-bridge.js"), "utf8");
const sheetsSource = fs.readFileSync(path.join(bridgeRoot, "bridges/sheets-bridge.js"), "utf8");

function loadBridge(source, api) {
  const sandbox = {
    window: {},
    Api: api,
    Asc: {
      scope: {},
      plugin: {
        info: {},
        callCommand(command, _close, _calc, callback) {
          callback(command());
        },
      },
    },
    console,
    JSON,
    Number,
    String,
    Boolean,
    Math,
    RegExp,
    Error,
    isFinite,
    parseInt,
  };
  vm.runInNewContext(source, sandbox);
  return sandbox.window.AICopilotBridges;
}

class SlideRun {
  constructor(text) {
    this.text = text;
    this.fontSize = 20;
    this.format = {};
  }
  GetClassType() { return "run"; }
  GetText() { return this.text; }
  RemoveAllElements() { this.text = ""; }
  AddText(text) { this.text += String(text); return this; }
  GetFontSize() { return this.fontSize; }
  SetFontSize(value) { this.fontSize = value; }
  SetFontFamily(value) { this.format.fontFamily = value; }
  SetBold(value) { this.format.bold = value; }
  SetItalic(value) { this.format.italic = value; }
  SetUnderline(value) { this.format.underline = value; }
  SetColor(value) { this.format.color = value; }
}

class SlideParagraph {
  constructor(text = "") {
    this.runs = text ? [new SlideRun(text)] : [];
    this.align = "left";
  }
  GetElementsCount() { return this.runs.length; }
  GetElement(index) { return this.runs[index]; }
  AddText(text) {
    const run = new SlideRun(String(text));
    this.runs.push(run);
    return run;
  }
  SetJc(value) { this.align = value; }
}

class SlideContent {
  constructor(text = "") {
    this.paragraphs = [new SlideParagraph(text)];
  }
  GetAllParagraphs() { return this.paragraphs; }
  GetText() {
    return this.paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join("")).join("\n");
  }
  RemoveAllElements() { this.paragraphs = []; }
  Push(paragraph) { this.paragraphs.push(paragraph); }
}

class SlideShape {
  constructor(text = "") {
    this.content = new SlideContent(text);
    this.position = [0, 0];
  }
  GetContent() { return this.content; }
  SetPosition(x, y) { this.position = [x, y]; }
}

class MockSlide {
  constructor(presentation, texts = []) {
    this.presentation = presentation;
    this.shapes = texts.map(text => new SlideShape(text));
    this.background = null;
  }
  GetAllShapes() { return this.shapes; }
  AddObject(shape) { this.shapes.push(shape); }
  SetBackground(fill) { this.background = fill; }
  Duplicate() {
    const copy = new MockSlide(this.presentation, this.shapes.map(shape => shape.GetContent().GetText()));
    const index = this.presentation.slides.indexOf(this);
    this.presentation.slides.splice(index + 1, 0, copy);
    return copy;
  }
  Delete() {
    const index = this.presentation.slides.indexOf(this);
    if (index < 0) return false;
    this.presentation.slides.splice(index, 1);
    return true;
  }
}

function slidesHarness() {
  const presentation = {
    slides: [],
    historyPoints: 0,
    GetSlidesCount() { return this.slides.length; },
    GetAllSlides() { return this.slides; },
    GetSlideByIndex(index) { return this.slides[index]; },
    GetCurSlideIndex() { return 0; },
    CreateNewHistoryPoint() { this.historyPoints += 1; },
    AddSlide(slide, index) {
      if (Number.isInteger(index)) this.slides.splice(index, 0, slide);
      else this.slides.push(slide);
    },
  };
  const first = new MockSlide(presentation, ["Alpha draft", "replace-me"]);
  const second = new MockSlide(presentation, ["Beta", "second slide"]);
  presentation.slides.push(first, second);
  const selection = { GetShapes: () => [first.shapes[0]] };
  const api = {
    GetPresentation: () => presentation,
    GetSelection: () => selection,
    Color: value => value,
    RGB: (r, g, b) => ({ r, g, b }),
    CreateSolidFill: value => ({ solid: value }),
    CreateNoFill: () => ({ none: true }),
    CreateStroke: (width, fill) => ({ width, fill }),
    CreateShape: () => new SlideShape(),
    CreateParagraph: () => new SlideParagraph(),
    CreateSlide: () => new MockSlide(presentation),
  };
  return { bridge: loadBridge(slidesSource, api).slide, presentation, first, selection };
}

class MockRange {
  constructor(address, values) {
    this.address = address;
    this.values = values;
    this.formula = null;
    this.format = {};
  }
  GetAddress() { return this.address; }
  GetValue() { return this.values; }
  SetValue(values) { this.values = values; return true; }
  SetFormula(formula) { this.formula = formula; return true; }
  Replace(search, replacement) {
    let changed = false;
    this.values = this.values.map(row => row.map(value => {
      if (typeof value !== "string" || !value.includes(search)) return value;
      changed = true;
      return value.split(search).join(replacement);
    }));
    return changed;
  }
  SetFontSize(value) { this.format.fontSize = value; }
  SetFontName(value) { this.format.fontName = value; }
  SetBold(value) { this.format.bold = value; }
  SetItalic(value) { this.format.italic = value; }
  SetUnderline(value) { this.format.underline = value; }
  SetFontColor(value) { this.format.fontColor = value; }
  SetFillColor(value) { this.format.fillColor = value; }
  SetAlignHorizontal(value) { this.format.horizontalAlign = value; }
  SetAlignVertical(value) { this.format.verticalAlign = value; }
  SetNumberFormat(value) { this.format.numberFormat = value; }
  SetWrap(value) { this.format.wrap = value; }
  SetColumnWidth(value) { this.format.columnWidth = value; }
  SetRowHeight(value) { this.format.rowHeight = value; }
}

class MockSheet {
  constructor(workbook, name, values) {
    this.workbook = workbook;
    this.name = name;
    this.range = new MockRange("A1:C3", values);
    this.charts = [];
  }
  GetName() { return this.name; }
  SetName(value) { this.name = value; }
  GetRange(address) { this.range.address = address; return this.range; }
  GetUsedRange() { return this.range; }
  AddChart(source, inRows, type, style, width, height, fromColumn, _columnOffset, fromRow) {
    const chart = {
      source, inRows, type, style, width, height, fromColumn, fromRow, title: null,
      SetTitle: (text, fontSize) => { chart.title = { text, fontSize }; },
    };
    this.charts.push(chart);
    return chart;
  }
  Delete() {
    const index = this.workbook.sheets.indexOf(this);
    if (index < 0) return false;
    this.workbook.sheets.splice(index, 1);
    return true;
  }
}

function sheetsHarness() {
  const workbook = { sheets: [], historyPoints: 0 };
  const sheet1 = new MockSheet(workbook, "Sheet1", [["draft", "value"], [1, 2]]);
  const sheet2 = new MockSheet(workbook, "Sheet2", [["rename-me"]]);
  workbook.sheets.push(sheet1, sheet2);
  const api = {
    GetSheets: () => workbook.sheets,
    GetActiveSheet: () => workbook.sheets[0],
    GetSelection: () => workbook.sheets[0].range,
    CreateNewHistoryPoint: () => { workbook.historyPoints += 1; },
    AddSheet: name => {
      if (workbook.sheets.some(sheet => sheet.GetName() === name)) return null;
      const sheet = new MockSheet(workbook, name, []);
      workbook.sheets.push(sheet);
      return sheet;
    },
    CreateColorFromRGB: (r, g, b) => ({ r, g, b }),
  };
  return { bridge: loadBridge(sheetsSource, api).cell, workbook, sheet1, sheet2 };
}

test("Slides bridge executes every public Slides tool in one history point", async () => {
  const { bridge, presentation, first } = slidesHarness();
  const result = await bridge.execute([
    { name: "slides_inspect", arguments: {} },
    { name: "slides_replace_text", arguments: { search: "draft", replace: "final", slide: 1 } },
    { name: "slides_scale_font", arguments: { scale: 1.1, slide: 1 } },
    { name: "slides_format_text", arguments: { slide: 1, bold: true, color: "#1F4E78" } },
    { name: "slides_format_selection", arguments: { italic: true } },
    { name: "slides_add_slide", arguments: { title: "Gamma" } },
    { name: "slides_duplicate_slide", arguments: { slide: 2 } },
    { name: "slides_add_textbox", arguments: { slide: 1, text: "bridge-added" } },
    { name: "slides_delete_slide", arguments: { slide: 3 } },
  ]);

  assert.deepEqual(Array.from(result.results, item => item.name), [
    "slides_inspect",
    "slides_replace_text",
    "slides_scale_font",
    "slides_format_text",
    "slides_format_selection",
    "slides_add_slide",
    "slides_duplicate_slide",
    "slides_add_textbox",
    "slides_delete_slide",
  ]);
  assert.equal(result.editorType, "slide");
  assert.equal(result.needsSave, true);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(presentation.slides.length, 3);
  assert.equal(first.shapes[0].GetContent().GetText(), "Alpha final");
  assert.equal(first.shapes[0].GetContent().GetAllParagraphs()[0].runs[0].format.bold, true);
  assert.equal(first.shapes[0].GetContent().GetAllParagraphs()[0].runs[0].format.italic, true);
  assert.equal(first.shapes.at(-1).GetContent().GetText(), "bridge-added");
});

test("Slides bridge rejects selection formatting when nothing is selected", async () => {
  const harness = slidesHarness();
  harness.selection.GetShapes = () => [];
  await assert.rejects(
    harness.bridge.execute([{ name: "slides_format_selection", arguments: { bold: true } }]),
    /请先在 PPT/,
  );
});

test("Slides bridge rejects a Sheets tool at the adapter boundary", async () => {
  const harness = slidesHarness();
  await assert.rejects(
    harness.bridge.execute([{ name: "sheets_inspect", arguments: {} }]),
    /Slides Bridge 不支持工具/,
  );
});

test("Sheets bridge executes every public Sheets tool in one history point", async () => {
  const { bridge, workbook, sheet1 } = sheetsHarness();
  const result = await bridge.execute([
    { name: "sheets_inspect", arguments: {} },
    { name: "sheets_set_values", arguments: { sheet: "Sheet1", range: "A1:B3", values: [["draft", "value"], [1, 2], [3, 4]] } },
    { name: "sheets_set_formula", arguments: { sheet: "Sheet1", range: "C2", formula: "SUM(A2:B2)" } },
    { name: "sheets_replace_text", arguments: { sheet: "Sheet1", range: "A1:B3", search: "draft", replace: "final" } },
    { name: "sheets_format_range", arguments: { sheet: "Sheet1", range: "A1:C3", bold: true, fillColor: "#D9EAF7" } },
    { name: "sheets_add_sheet", arguments: { name: "Temp" } },
    { name: "sheets_rename_sheet", arguments: { sheet: "Sheet2", newName: "Renamed" } },
    { name: "sheets_add_chart", arguments: { sheet: "Sheet1", range: "A1:B3", type: "bar", title: "Validation" } },
    { name: "sheets_delete_sheet", arguments: { sheet: "Temp" } },
  ]);

  assert.deepEqual(Array.from(result.results, item => item.name), [
    "sheets_inspect",
    "sheets_set_values",
    "sheets_set_formula",
    "sheets_replace_text",
    "sheets_format_range",
    "sheets_add_sheet",
    "sheets_rename_sheet",
    "sheets_add_chart",
    "sheets_delete_sheet",
  ]);
  assert.equal(result.editorType, "cell");
  assert.equal(result.needsSave, true);
  assert.equal(workbook.historyPoints, 1);
  assert.deepEqual(workbook.sheets.map(sheet => sheet.GetName()), ["Sheet1", "Renamed"]);
  assert.equal(sheet1.range.values[0][0], "final");
  assert.equal(sheet1.range.formula, "=SUM(A2:B2)");
  assert.equal(sheet1.range.format.bold, true);
  assert.equal(sheet1.range.format.fillColor.r, 217);
  assert.equal(sheet1.charts.length, 1);
  assert.deepEqual(sheet1.charts[0].title, { text: "Validation", fontSize: 13 });
});

test("Sheets bridge protects the last remaining worksheet", async () => {
  const harness = sheetsHarness();
  harness.workbook.sheets.splice(1, 1);
  await assert.rejects(
    harness.bridge.execute([{ name: "sheets_delete_sheet", arguments: { sheet: "Sheet1" } }]),
    /至少需要保留一个工作表/,
  );
});

test("Sheets bridge falls back to SetValue for formulas in community editions", async () => {
  const harness = sheetsHarness();
  harness.sheet1.range.SetFormula = undefined;
  const result = await harness.bridge.execute([
    { name: "sheets_set_formula", arguments: { sheet: "Sheet1", range: "C2", formula: "SUM(A2:B2)" } },
  ]);

  assert.equal(result.changed, 1);
  assert.equal(result.needsSave, true);
  assert.equal(harness.sheet1.range.values, "=SUM(A2:B2)");
});

test("Sheets bridge rejects a Slides tool at the adapter boundary", async () => {
  const harness = sheetsHarness();
  await assert.rejects(
    harness.bridge.execute([{ name: "slides_inspect", arguments: {} }]),
    /Sheets Bridge 不支持工具/,
  );
});
