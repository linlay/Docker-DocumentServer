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

let drawingId = 0;

function mockFill(type, data = {}) {
  return {
    kind: "fill",
    type,
    ...data,
    GetType() { return this.type; },
    ToJSON() {
      const { kind, type, color, stops, angle, pattern, background, foreground } = this;
      return JSON.stringify({ kind, type, color, stops, angle, pattern, background, foreground });
    },
  };
}

function mockStroke(width, fill) {
  return {
    kind: "line",
    width,
    fill,
    GetWidth() { return this.width; },
    GetFill() { return this.fill; },
    ToJSON() {
      return JSON.stringify({
        kind: "line",
        width: this.width,
        fill: JSON.parse(this.fill.ToJSON()),
      });
    },
  };
}

class SlideShape {
  constructor(text = "") {
    this.content = new SlideContent(text);
    this.position = [0, 0];
    this.size = [3600000, 2160000];
    this.rotation = 0;
    this.flipH = false;
    this.flipV = false;
    this.name = "";
    this.internalId = `drawing-${++drawingId}`;
    this.fill = mockFill("none");
    this.line = mockStroke(0, mockFill("none"));
    this.geometry = { preset: "rect", GetPreset() { return this.preset; } };
    this.parent = null;
  }
  GetClassType() { return "shape"; }
  GetInternalId() { return this.internalId; }
  GetName() { return this.name; }
  SetName(value) { this.name = value; }
  GetContent() { return this.content; }
  SetPosition(x, y) { this.position = [x, y]; }
  GetPosX() { return this.position[0]; }
  GetPosY() { return this.position[1]; }
  SetSize(width, height) { this.size = [width, height]; }
  GetWidth() { return this.size[0]; }
  GetHeight() { return this.size[1]; }
  SetRotation(value) { this.rotation = value; }
  GetRotation() { return this.rotation; }
  SetFlipH(value) { this.flipH = value; }
  SetFlipV(value) { this.flipV = value; }
  GetFlipH() { return this.flipH; }
  GetFlipV() { return this.flipV; }
  SetFill(value) { this.fill = value; }
  Fill(value) { this.fill = value; }
  GetFill() { return this.fill; }
  SetLine(value) { this.line = value; }
  SetOutLine(value) { this.line = value; }
  GetLine() { return this.line; }
  SetGeometry(value) { this.geometry = value; }
  GetGeometry() { return this.geometry; }
  SetVerticalTextAlign(value) { this.verticalAlign = value; }
  SetPaddings(left, top, right, bottom) { this.paddings = [left, top, right, bottom]; }
  Delete() {
    if (!this.parent) return false;
    return this.parent.RemoveObject(this);
  }
  ToJSON() {
    return JSON.stringify({
      kind: "shape",
      id: this.internalId,
      name: this.name,
      fill: JSON.parse(this.fill.ToJSON()),
      line: JSON.parse(this.line.ToJSON()),
      geometry: this.geometry.GetPreset(),
    });
  }
}

class SlideChart {
  constructor(type, series, seriesNames, categories, width, height, style) {
    this.type = type;
    this.values = series;
    this.seriesNames = seriesNames;
    this.categories = categories;
    this.size = [width, height];
    this.style = style;
    this.position = [0, 0];
    this.rotation = 0;
    this.name = "";
    this.title = "";
    this.internalId = `drawing-${++drawingId}`;
    this.options = {};
    this.parent = null;
  }
  GetClassType() { return "chart"; }
  GetInternalId() { return this.internalId; }
  GetName() { return this.name; }
  SetName(value) { this.name = value; }
  GetChartType() { return this.type; }
  GetTitle() { return this.title; }
  GetAllSeries() {
    return this.values.map((_, index) => ({
      GetChartType: () => this.type,
      ToJSON: () => JSON.stringify({ index, name: this.seriesNames[index], values: this.values[index] }),
    }));
  }
  GetSeries(index) {
    return {
      ChangeChartType: value => { this.options.seriesType = [index, value]; },
    };
  }
  SetPosition(x, y) { this.position = [x, y]; }
  GetPosX() { return this.position[0]; }
  GetPosY() { return this.position[1]; }
  SetSize(width, height) { this.size = [width, height]; }
  GetWidth() { return this.size[0]; }
  GetHeight() { return this.size[1]; }
  SetRotation(value) { this.rotation = value; }
  GetRotation() { return this.rotation; }
  ApplyChartStyle(value) { this.style = value; }
  SetTitle(value, fontSize) { this.title = value; this.options.titleFontSize = fontSize; }
  Fill(value) { this.options.fill = value; }
  SetOutLine(value) { this.options.line = value; }
  SetPlotAreaFill(value) { this.options.plotAreaFill = value; }
  SetPlotAreaOutLine(value) { this.options.plotAreaLine = value; }
  SetTitleFill(value) { this.options.titleFill = value; }
  SetTitleOutLine(value) { this.options.titleLine = value; }
  SetLegendPos(value) { this.options.legendPosition = value; }
  SetLegendFontSize(value) { this.options.legendFontSize = value; }
  SetLegendFill(value) { this.options.legendFill = value; }
  SetLegendOutLine(value) { this.options.legendLine = value; }
  SetHorAxisTitle(value, size) { this.options.horizontalTitle = [value, size]; }
  SetVerAxisTitle(value, size) { this.options.verticalTitle = [value, size]; }
  SetHorAxisLabelsFontSize(value) { this.options.horizontalLabelsFontSize = value; }
  SetVertAxisLabelsFontSize(value) { this.options.verticalLabelsFontSize = value; }
  SetHorAxisOrientation(value) { this.options.horizontalNormalOrder = value; }
  SetVertAxisOrientation(value) { this.options.verticalNormalOrder = value; }
  SetHorAxisMajorTickMark(value) { this.options.horizontalMajorTickMark = value; }
  SetVertAxisMajorTickMark(value) { this.options.verticalMajorTickMark = value; }
  SetHorAxisMinorTickMark(value) { this.options.horizontalMinorTickMark = value; }
  SetVertAxisMinorTickMark(value) { this.options.verticalMinorTickMark = value; }
  SetHorAxisTickLabelPosition(value) { this.options.horizontalTickLabelPosition = value; }
  SetVertAxisTickLabelPosition(value) { this.options.verticalTickLabelPosition = value; }
  SetAxieNumFormat(format, position) { this.options.axisNumberFormat = [format, position]; }
  SetShowDataLabels(...values) { this.options.dataLabels = values; }
  SetMajorHorizontalGridlines(value) { this.options.majorHorizontal = value; }
  SetMinorHorizontalGridlines(value) { this.options.minorHorizontal = value; }
  SetMajorVerticalGridlines(value) { this.options.majorVertical = value; }
  SetMinorVerticalGridlines(value) { this.options.minorVertical = value; }
  SetCategoryName(value, index) { this.categories[index] = value; }
  SetSeriaName(value, index) { this.seriesNames[index] = value; }
  SetSeriaValues(value, index) { this.values[index] = value; }
  SetXValues(value) { this.options.xValues = value; }
  SetSeriaNumFormat(value, index) { this.options.seriesNumberFormat = [index, value]; }
  SetSeriesFill(value, index, allSeries) { this.options.seriesFill = [index, allSeries, value]; }
  SetSeriesOutLine(value, index, allSeries) { this.options.seriesLine = [index, allSeries, value]; }
  SetDataPointFill(value, series, point, allSeries) { this.options.pointFill = [series, point, allSeries, value]; }
  SetDataPointOutLine(value, series, point, allSeries) { this.options.pointLine = [series, point, allSeries, value]; }
  SetMarkerFill(value, series, point, allMarkers) { this.options.markerFill = [series, point, allMarkers, value]; }
  SetMarkerOutLine(value, series, point, allMarkers) { this.options.markerLine = [series, point, allMarkers, value]; }
  SetDataPointNumFormat(value, series, point) { this.options.pointNumberFormat = [series, point, value]; }
  SetShowPointDataLabel(series, point, ...values) { this.options.pointDataLabels = [series, point, ...values]; }
  RemoveSeria(index) { this.values.splice(index, 1); this.seriesNames.splice(index, 1); }
  Delete() {
    if (!this.parent) return false;
    return this.parent.RemoveObject(this);
  }
  ToJSON() {
    return JSON.stringify({
      kind: "chart",
      id: this.internalId,
      name: this.name,
      type: this.type,
      title: this.title,
      style: this.style,
      categories: this.categories,
      seriesNames: this.seriesNames,
      values: this.values,
    });
  }
}

class MockSlide {
  constructor(presentation, texts = []) {
    this.presentation = presentation;
    this.shapes = texts.map(text => new SlideShape(text));
    this.shapes.forEach(shape => { shape.parent = this; });
    this.charts = [];
    this.background = null;
  }
  GetAllShapes() { return this.shapes; }
  GetAllCharts() { return this.charts; }
  GetAllDrawings() { return [...this.shapes, ...this.charts]; }
  GetWidth() { return 9144000; }
  GetHeight() { return 5143500; }
  AddObject(shape) {
    shape.parent = this;
    if (shape.GetClassType() === "chart") this.charts.push(shape);
    else this.shapes.push(shape);
  }
  RemoveObject(shape) {
    const collection = shape.GetClassType() === "chart" ? this.charts : this.shapes;
    const index = collection.indexOf(shape);
    if (index < 0) return false;
    collection.splice(index, 1);
    this.lastRemovedObject = shape;
    shape.parent = null;
    return true;
  }
  SetBackground(fill) { this.background = fill; }
  ClearBackground() { this.background = null; }
  FollowLayoutBackground() { this.background = "layout"; }
  FollowMasterBackground() { this.background = "master"; }
  ToJSON() { return JSON.stringify({ shapes: this.shapes.length, charts: this.charts.length }); }
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
    CreateSolidFill: value => mockFill("solid", { color: value }),
    CreateNoFill: () => mockFill("none"),
    CreateGradientStop: (color, position) => ({ color, position }),
    CreateLinearGradientFill: (stops, angle) => mockFill("linearGradient", { stops, angle }),
    CreateRadialGradientFill: stops => mockFill("radialGradient", { stops }),
    CreatePatternFill: (pattern, background, foreground) => mockFill("pattern", { pattern, background, foreground }),
    CreateStroke: (width, fill) => mockStroke(width, fill),
    CreateShape: (type, width, height, fill, line) => {
      const shape = new SlideShape();
      shape.geometry.preset = type;
      shape.size = [width, height];
      shape.fill = fill;
      shape.line = line;
      return shape;
    },
    CreatePresetGeometry: type => ({ preset: type, GetPreset() { return this.preset; } }),
    CreateChart: (...args) => new SlideChart(...args),
    FromJSON: raw => {
      if (typeof raw === "string") raw = JSON.parse(raw);
      if (!raw || typeof raw !== "object") return null;
      if (raw.kind === "fill") return mockFill(raw.type, raw);
      if (raw.kind === "line") return mockStroke(raw.width, mockFill(raw.fill.type, raw.fill));
      return null;
    },
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

class MockSheetChart {
  constructor(sheet, source, inRows, type, style, width, height, fromColumn, columnOffset, fromRow, rowOffset) {
    this.sheet = sheet;
    this.source = source;
    this.inRows = inRows;
    this.type = type;
    this.style = style;
    this.size = [width, height];
    this.position = [fromColumn, columnOffset, fromRow, rowOffset];
    this.rotation = 0;
    this.name = "";
    this.title = "";
    this.options = {};
    this.series = [{ name: "Series 1", valuesRange: source }];
  }
  GetName() { return this.name; }
  SetName(value) { this.name = value; }
  GetChartType() { return this.type; }
  GetTitle() { return this.title; }
  GetWidth() { return this.size[0]; }
  GetHeight() { return this.size[1]; }
  GetRotation() { return this.rotation; }
  GetAllSeries() {
    return this.series.map((series, index) => ({
      GetChartType: () => this.type,
      ToJSON: () => JSON.stringify({ index, ...series }),
    }));
  }
  GetSeries(index) {
    return {
      ChangeChartType: value => { this.options.seriesType = [index, value]; },
    };
  }
  SetPosition(fromColumn, columnOffset, fromRow, rowOffset) {
    this.position = [fromColumn, columnOffset, fromRow, rowOffset];
  }
  SetSize(width, height) { this.size = [width, height]; }
  SetRotation(value) { this.rotation = value; }
  ApplyChartStyle(value) { this.style = value; }
  SetTitle(value, fontSize) { this.title = value; this.options.titleFontSize = fontSize; }
  Fill(value) { this.options.fill = value; }
  SetOutLine(value) { this.options.line = value; }
  SetPlotAreaFill(value) { this.options.plotAreaFill = value; }
  SetPlotAreaOutLine(value) { this.options.plotAreaLine = value; }
  SetTitleFill(value) { this.options.titleFill = value; }
  SetTitleOutLine(value) { this.options.titleLine = value; }
  SetLegendPos(value) { this.options.legendPosition = value; }
  SetLegendFontSize(value) { this.options.legendFontSize = value; }
  SetLegendFill(value) { this.options.legendFill = value; }
  SetLegendOutLine(value) { this.options.legendLine = value; }
  SetHorAxisTitle(value, size) { this.options.horizontalTitle = [value, size]; }
  SetVerAxisTitle(value, size) { this.options.verticalTitle = [value, size]; }
  SetHorAxisLabelsFontSize(value) { this.options.horizontalLabelsFontSize = value; }
  SetVertAxisLabelsFontSize(value) { this.options.verticalLabelsFontSize = value; }
  SetHorAxisOrientation(value) { this.options.horizontalNormalOrder = value; }
  SetVertAxisOrientation(value) { this.options.verticalNormalOrder = value; }
  SetHorAxisMajorTickMark(value) { this.options.horizontalMajorTickMark = value; }
  SetVertAxisMajorTickMark(value) { this.options.verticalMajorTickMark = value; }
  SetHorAxisMinorTickMark(value) { this.options.horizontalMinorTickMark = value; }
  SetVertAxisMinorTickMark(value) { this.options.verticalMinorTickMark = value; }
  SetHorAxisTickLabelPosition(value) { this.options.horizontalTickLabelPosition = value; }
  SetVertAxisTickLabelPosition(value) { this.options.verticalTickLabelPosition = value; }
  SetAxieNumFormat(format, position) { this.options.axisNumberFormat = [format, position]; }
  SetShowDataLabels(...values) { this.options.dataLabels = values; }
  SetMajorHorizontalGridlines(value) { this.options.majorHorizontal = value; }
  SetMinorHorizontalGridlines(value) { this.options.minorHorizontal = value; }
  SetMajorVerticalGridlines(value) { this.options.majorVertical = value; }
  SetMinorVerticalGridlines(value) { this.options.minorVertical = value; }
  SetCatFormula(value) { this.options.categoryRange = value; }
  AddSeria(name, valuesRange, xValuesRange) { this.series.push({ name, valuesRange, xValuesRange }); }
  SetSeriaName(value, index) { this.series[index].name = value; }
  SetSeriaValues(value, index) { this.series[index].valuesRange = value; }
  SetSeriaXValues(value, index) { this.series[index].xValuesRange = value; }
  SetSeriaNumFormat(value, index) { this.options.seriesNumberFormat = [index, value]; }
  SetSeriesFill(value, index, allSeries) { this.options.seriesFill = [index, allSeries, value]; }
  SetSeriesOutLine(value, index, allSeries) { this.options.seriesLine = [index, allSeries, value]; }
  SetDataPointFill(value, series, point, allSeries) { this.options.pointFill = [series, point, allSeries, value]; }
  SetDataPointOutLine(value, series, point, allSeries) { this.options.pointLine = [series, point, allSeries, value]; }
  SetMarkerFill(value, series, point, allMarkers) { this.options.markerFill = [series, point, allMarkers, value]; }
  SetMarkerOutLine(value, series, point, allMarkers) { this.options.markerLine = [series, point, allMarkers, value]; }
  SetDataPointNumFormat(value, series, point) { this.options.pointNumberFormat = [series, point, value]; }
  SetShowPointDataLabel(series, point, ...values) { this.options.pointDataLabels = [series, point, ...values]; }
  RemoveSeria(index) { this.series.splice(index, 1); }
  Delete() {
    const index = this.sheet.charts.indexOf(this);
    if (index < 0) return false;
    this.sheet.charts.splice(index, 1);
    return true;
  }
  ToJSON() {
    return JSON.stringify({
      kind: "chart",
      name: this.name,
      type: this.type,
      title: this.title,
      style: this.style,
      position: this.position,
      size: this.size,
      series: this.series,
    });
  }
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
  GetAllCharts() { return this.charts; }
  AddChart(source, inRows, type, style, width, height, fromColumn, columnOffset, fromRow, rowOffset) {
    const chart = new MockSheetChart(
      this,
      source,
      inRows,
      type,
      style,
      width,
      height,
      fromColumn,
      columnOffset,
      fromRow,
      rowOffset,
    );
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
    CreateSolidFill: value => mockFill("solid", { color: value }),
    CreateNoFill: () => mockFill("none"),
    CreateGradientStop: (color, position) => ({ color, position }),
    CreateLinearGradientFill: (stops, angle) => mockFill("linearGradient", { stops, angle }),
    CreateRadialGradientFill: stops => mockFill("radialGradient", { stops }),
    CreatePatternFill: (pattern, background, foreground) => mockFill("pattern", { pattern, background, foreground }),
    CreateStroke: (width, fill) => mockStroke(width, fill),
    FromJSON: raw => {
      if (typeof raw === "string") raw = JSON.parse(raw);
      if (!raw || typeof raw !== "object") return null;
      if (raw.kind === "fill") return mockFill(raw.type, raw);
      if (raw.kind === "line") return mockStroke(raw.width, mockFill(raw.fill.type, raw.fill));
      return null;
    },
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

test("Slides bridge round-trips gradients and supports shape and chart CRUD", async () => {
  const { bridge, presentation, first } = slidesHarness();
  const result = await bridge.execute([
    {
      name: "slides_add_shape",
      arguments: {
        slide: 1,
        shapeType: "roundRect",
        name: "GradientShape",
        text: "Gradient",
        xMm: 20,
        yMm: 25,
        widthMm: 80,
        heightMm: 35,
        fill: {
          type: "linearGradient",
          angleDeg: 45,
          stops: [
            { position: 0, color: "#FF0000" },
            { position: 100, color: "#0000FF" },
          ],
        },
        line: { widthPt: 2, color: "#111111" },
      },
    },
    { name: "slides_inspect_objects", arguments: { slide: 1, includeRaw: true } },
    {
      name: "slides_update_shape",
      arguments: {
        slide: 1,
        name: "GradientShape",
        shapeType: "ellipse",
        rotationDeg: 12,
        bold: true,
        align: "center",
        fill: {
          type: "radialGradient",
          stops: [
            { position: 0, color: "#FFFFFF" },
            { position: 100, color: "#00AA88" },
          ],
        },
        includeRaw: true,
      },
    },
    {
      name: "slides_add_chart",
      arguments: {
        slide: 1,
        type: "bar",
        series: [[10, 20], [5, 15]],
        seriesNames: ["Revenue", "Cost"],
        categories: ["Q1", "Q2"],
        name: "SalesChart",
        title: "Sales",
        legend: { position: "bottom" },
        fill: {
          type: "linearGradient",
          stops: [
            { position: 0, color: "#F8FBFF" },
            { position: 100, color: "#DDEEFF" },
          ],
        },
      },
    },
    { name: "slides_inspect_charts", arguments: { slide: 1, includeRaw: true } },
    {
      name: "slides_update_chart",
      arguments: {
        slide: 1,
        chartIndex: 0,
        title: "Updated Sales",
        categories: ["H1", "H2"],
        horizontalAxis: { title: "Period", numberFormat: "General" },
        dataLabels: { showValue: true },
        seriesUpdates: [{
          index: 0,
          type: "area",
          name: "Net revenue",
          values: [12, 24],
          fill: { type: "solid", color: "#3366CC" },
          points: [{
            index: 1,
            fill: { type: "solid", color: "#FF9900" },
            markerFill: { type: "solid", color: "#FFFFFF" },
            markerLine: { widthPt: 1, color: "#222222" },
          }],
        }],
        includeRaw: true,
      },
    },
    { name: "slides_delete_chart", arguments: { slide: 1, chartIndex: 0 } },
    { name: "slides_delete_object", arguments: { slide: 1, name: "GradientShape" } },
  ]);

  assert.equal(result.changed, 6);
  assert.equal(result.needsSave, true);
  assert.equal(presentation.historyPoints, 1);

  const addShape = result.results[0].object;
  assert.equal(addShape.name, "GradientShape");
  assert.equal(addShape.fill.type, "linearGradient");
  assert.equal(addShape.fill.raw.angle, 45 * 60000);
  assert.deepEqual(Array.from(addShape.fill.raw.stops, stop => stop.position), [0, 100000]);

  const inspectedShape = result.results[1].slides[0].objects.find(object => object.name === "GradientShape");
  assert.equal(inspectedShape.fill.type, "linearGradient");
  assert.equal(inspectedShape.objectIndex, 2);

  const updatedShape = result.results[2].object;
  assert.equal(updatedShape.shapeType, "ellipse");
  assert.equal(updatedShape.rotationDeg, 12);
  assert.equal(updatedShape.fill.type, "radialGradient");
  assert.equal(first.lastRemovedObject.GetContent().GetAllParagraphs()[0].runs[0].format.bold, true);

  assert.equal(result.results[3].chart.name, "SalesChart");
  assert.equal(result.results[4].slides[0].charts[0].chartType, "bar");
  assert.equal(result.results[5].chart.title, "Updated Sales");
  assert.deepEqual(first.charts, []);
  assert.equal(first.shapes.some(shape => shape.GetName() === "GradientShape"), false);
});

test("Slides object and chart inspection is read-only", async () => {
  const { bridge, presentation } = slidesHarness();
  const result = await bridge.execute([
    { name: "slides_inspect_objects", arguments: {} },
    { name: "slides_inspect_charts", arguments: {} },
  ]);
  assert.equal(result.changed, 0);
  assert.equal(result.needsSave, false);
  assert.equal(presentation.historyPoints, 0);
});

test("Slides bridge replays inspected fill.raw through Api.FromJSON", async () => {
  const { bridge } = slidesHarness();
  const created = await bridge.execute([{
    name: "slides_add_shape",
    arguments: {
      slide: 1,
      shapeType: "rect",
      name: "RawReplay",
      fill: {
        type: "linearGradient",
        stops: [
          { position: 0, color: "#112233" },
          { position: 100, color: "#AABBCC" },
        ],
      },
    },
  }]);
  const raw = created.results[0].object.fill.raw;
  const replayed = await bridge.execute([{
    name: "slides_update_shape",
    arguments: { slide: 1, name: "RawReplay", fill: { raw } },
  }]);
  assert.equal(replayed.results[0].object.fill.type, "linearGradient");
  assert.deepEqual(Array.from(replayed.results[0].object.fill.raw.stops, stop => stop.position), [0, 100000]);
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
  assert.equal(sheet1.charts[0].title, "Validation");
  assert.equal(sheet1.charts[0].options.titleFontSize, 13);
});

test("Sheets bridge supports chart inspect, gradient formatting, update, and delete", async () => {
  const { bridge, workbook, sheet1 } = sheetsHarness();
  const result = await bridge.execute([
    {
      name: "sheets_add_chart",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        type: "line",
        name: "TrendChart",
        title: "Trend",
        widthMm: 150,
        heightMm: 80,
        fromColumn: 3,
        fromRow: 1,
        fill: {
          type: "linearGradient",
          angleDeg: 90,
          stops: [
            { position: 0, color: "#FFFFFF" },
            { position: 100, color: "#DDEEFF" },
          ],
        },
      },
    },
    { name: "sheets_inspect_charts", arguments: { sheet: "Sheet1", includeRaw: true } },
    {
      name: "sheets_update_chart",
      arguments: {
        sheet: "Sheet1",
        chartIndex: 0,
        title: "Updated Trend",
        categoryRange: "'Sheet1'!$A$2:$A$3",
        addSeries: [{ name: "Forecast", valuesRange: "'Sheet1'!$C$2:$C$3" }],
        horizontalAxis: { title: "Month" },
        verticalAxis: { title: "Value", numberFormat: "0.0" },
        dataLabels: { showValue: true },
        seriesUpdates: [{
          index: 0,
          type: "lineMarker",
          name: "Actual",
          valuesRange: "'Sheet1'!$B$2:$B$3",
          fill: { type: "solid", color: "#2255AA" },
        }],
        includeRaw: true,
      },
    },
    { name: "sheets_delete_chart", arguments: { sheet: "Sheet1", chartIndex: 0 } },
  ]);

  assert.equal(result.changed, 3);
  assert.equal(result.needsSave, true);
  assert.equal(workbook.historyPoints, 1);
  assert.equal(result.results[0].chart.name, "TrendChart");
  assert.equal(result.results[1].sheets[0].charts[0].chartIndex, 0);
  assert.equal(result.results[1].sheets[0].charts[0].raw.name, "TrendChart");
  assert.equal(result.results[2].chart.title, "Updated Trend");
  assert.equal(result.results[2].chart.raw.series[0].name, "Actual");
  assert.equal(result.results[2].chart.raw.series[1].name, "Forecast");
  assert.equal(sheet1.charts.length, 0);
});

test("Sheets chart inspection is read-only", async () => {
  const { bridge, workbook } = sheetsHarness();
  const result = await bridge.execute([
    { name: "sheets_inspect_charts", arguments: {} },
  ]);
  assert.equal(result.changed, 0);
  assert.equal(result.needsSave, false);
  assert.equal(workbook.historyPoints, 0);
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
