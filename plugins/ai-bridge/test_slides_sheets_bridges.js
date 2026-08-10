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
  const asc = {
    scope: {},
    plugin: {
      info: {},
      callCommand(command, _close, _calc, callback) {
        callback(command());
      },
      executeMethod(method, params, callback) {
        if (typeof api.__executeMethod === "function") {
          api.__executeMethod(method, params, callback);
          return;
        }
        callback(null);
      },
    },
  };
  const sandbox = {
    window: { Asc: asc },
    Api: api,
    Asc: asc,
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
  GetFontFamily() { return this.format.fontFamily ?? null; }
  GetBold() { return this.format.bold ?? false; }
  GetItalic() { return this.format.italic ?? false; }
  GetUnderline() { return this.format.underline ?? false; }
  GetColor() { return this.format.color ?? null; }
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
  GetParaPr() { return this; }
  GetElementsCount() { return this.runs.length; }
  GetElement(index) { return this.runs[index]; }
  AddText(text) {
    const run = new SlideRun(String(text));
    this.runs.push(run);
    return run;
  }
  GetJc() { return this.align; }
  GetBullet() { return this.bullet ?? null; }
  SetJc(value) { this.align = value; }
  SetIndFirstLine(value) { this.firstLineIndent = value; }
  SetIndLeft(value) { this.leftIndent = value; }
  SetIndRight(value) { this.rightIndent = value; }
  SetSpacingBefore(value, auto) { this.spacingBefore = [value, auto]; }
  SetSpacingAfter(value, auto) { this.spacingAfter = [value, auto]; }
  SetSpacingLine(value, rule) { this.spacingLine = [value, rule]; }
  SetOutlineLvl(value) { this.outlineLevel = value; }
  SetBullet(value) { this.bullet = value; }
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
    this.hyperlink = null;
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
  SetPlaceholder(value) { this.placeholder = value; return true; }
  GetPlaceholder() { return this.placeholder || null; }
  SetHyperlink(value) { this.hyperlink = value; return true; }
  GetHyperlink() { return this.hyperlink; }
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

class SlideImage extends SlideShape {
  constructor(url, width, height) {
    super();
    this.url = url;
    this.size = [width, height];
  }
  GetClassType() { return "image"; }
  ToJSON() {
    return JSON.stringify({
      kind: "image",
      id: this.internalId,
      name: this.name,
      url: this.url,
      position: this.position,
      size: this.size,
      rotation: this.rotation,
      flipH: this.flipH,
      flipV: this.flipV,
    });
  }
}

class MockOleObject extends SlideShape {
  constructor(url, width, height, data, applicationId) {
    super();
    this.url = url;
    this.size = [width, height];
    this.data = data;
    this.applicationId = applicationId;
  }
  GetClassType() { return "oleObject"; }
  GetApplicationId() { return this.applicationId; }
  GetData() { return this.data; }
  ToJSON() {
    return JSON.stringify({
      kind: "oleObject",
      id: this.internalId,
      name: this.name,
      applicationId: this.applicationId,
      data: this.data,
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

class MockLayout {
  constructor(name, layoutType, master = null) {
    this.name = name;
    this.layoutType = layoutType;
    this.drawings = [];
    this.master = master;
    this.background = null;
  }
  GetName() { return this.name; }
  SetName(value) { this.name = value; return true; }
  GetLayoutType() { return this.layoutType; }
  GetMaster() { return this.master; }
  GetAllDrawings() { return this.drawings; }
  AddObject(drawing) { drawing.parent = this; this.drawings.push(drawing); return true; }
  RemoveObject(drawing) {
    const index = this.drawings.indexOf(drawing);
    if (index < 0) return false;
    this.drawings.splice(index, 1);
    drawing.parent = null;
    return true;
  }
  SetBackground(fill) { this.background = fill; return true; }
  GetBackground() { return this.background; }
  ClearBackground() { this.background = null; return true; }
  FollowMasterBackground() { this.background = "master"; return true; }
  ToJSON() { return JSON.stringify({ name: this.name, layoutType: this.layoutType }); }
}

class MockTheme {
  constructor(name = "Office") {
    this.name = name;
    this.colorScheme = { name: "Office colors", ToJSON() { return JSON.stringify({ name: this.name }); } };
    this.fontScheme = { name: "Office fonts", ToJSON() { return JSON.stringify({ name: this.name }); } };
    this.formatScheme = { name: "Office format", ToJSON() { return JSON.stringify({ name: this.name }); } };
  }
  GetColorScheme() { return this.colorScheme; }
  GetFontScheme() { return this.fontScheme; }
  GetFormatScheme() { return this.formatScheme; }
  SetColorScheme(value) { this.colorScheme = value; return true; }
  SetFontScheme(value) { this.fontScheme = value; return true; }
  ToJSON() { return JSON.stringify({ name: this.name }); }
}

class MockMaster {
  constructor(layouts) {
    this.layouts = layouts;
    this.layouts.forEach(layout => { layout.master = this; });
    this.drawings = [];
    this.background = null;
    this.theme = new MockTheme();
  }
  GetLayoutsCount() { return this.layouts.length; }
  GetAllLayouts() { return this.layouts; }
  GetLayout(index) { return this.layouts[index] || null; }
  GetTheme() { return this.theme; }
  GetAllDrawings() { return this.drawings; }
  AddObject(drawing) { drawing.parent = this; this.drawings.push(drawing); return true; }
  RemoveObject(drawing) {
    const index = this.drawings.indexOf(drawing);
    if (index < 0) return false;
    this.drawings.splice(index, 1);
    drawing.parent = null;
    return true;
  }
  SetBackground(fill) { this.background = fill; return true; }
  GetBackground() { return this.background; }
  ClearBackground() { this.background = null; return true; }
  ToJSON() { return JSON.stringify({ layouts: this.layouts.map(layout => layout.name) }); }
}

class MockTableCell {
  constructor(text = "") {
    this.content = new SlideContent(text);
    this.fill = null;
    this.verticalAlign = null;
    this.borders = {};
    this.row = null;
  }
  GetClassType() { return "tableCell"; }
  GetContent() { return this.content; }
  GetText() { return this.content.GetText(); }
  SetText(text) {
    this.content.RemoveAllElements();
    this.content.Push(new SlideParagraph(String(text)));
    return new SlideRun(String(text));
  }
  SetShd(fill) { this.fill = fill; }
  SetVerticalAlign(value) { this.verticalAlign = value; }
  SetCellBorderTop(width, fill) { this.borders.top = [width, fill]; }
  SetCellBorderRight(width, fill) { this.borders.right = [width, fill]; }
  SetCellBorderBottom(width, fill) { this.borders.bottom = [width, fill]; }
  SetCellBorderLeft(width, fill) { this.borders.left = [width, fill]; }
  Split(rows, columns) { this.split = [rows, columns]; return true; }
}

class MockTableRow {
  constructor(count) {
    this.cells = Array.from({ length: count }, () => new MockTableCell());
    this.cells.forEach(cell => { cell.row = this; });
    this.height = null;
    this.table = null;
  }
  GetCell(index) { return this.cells[index] || null; }
  GetCellsCount() { return this.cells.length; }
  SetHeight(value) { this.height = value; return value; }
  GetHeight() { return this.height; }
}

class MockTable {
  constructor(rows, columns) {
    this.rows = Array.from({ length: rows }, () => new MockTableRow(columns));
    this.rows.forEach(row => { row.table = this; });
    this.columnWidths = Array.from({ length: columns }, () => 360000);
    this.position = [0, 0];
    this.size = [3600000, 1800000];
    this.rotation = 0;
    this.flipH = false;
    this.flipV = false;
    this.name = "";
    this.internalId = `drawing-${++drawingId}`;
    this.parent = null;
    this.tableLook = null;
  }
  GetClassType() { return "table"; }
  GetInternalId() { return this.internalId; }
  GetName() { return this.name; }
  SetName(value) { this.name = value; }
  GetRow(index) { return this.rows[index] || null; }
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
  SetColumnWidth(index, value) { this.columnWidths[index] = value; return value; }
  GetColumnWidth(index) { return this.columnWidths[index] ?? null; }
  SetTableLook(...values) { this.tableLook = values; }
  AddRow(cell, before) {
    const currentColumns = Math.max(1, ...this.rows.map(row => row.cells.length));
    const row = new MockTableRow(currentColumns);
    row.table = this;
    if (!cell) {
      this.rows.push(row);
      return row;
    }
    const index = this.rows.indexOf(cell.row);
    this.rows.splice(before ? index : index + 1, 0, row);
    return row;
  }
  AddColumn(cell, before) {
    const column = cell ? cell.row.cells.indexOf(cell) : -1;
    const target = column < 0 ? this.columnWidths.length : (before ? column : column + 1);
    for (const row of this.rows) {
      const inserted = new MockTableCell();
      inserted.row = row;
      row.cells.splice(target, 0, inserted);
    }
    this.columnWidths.splice(target, 0, 360000);
  }
  RemoveRow(cell) {
    const index = this.rows.indexOf(cell.row);
    if (index < 0 || this.rows.length <= 1) return false;
    this.rows.splice(index, 1);
    return true;
  }
  RemoveColumn(cell) {
    const index = cell.row.cells.indexOf(cell);
    if (index < 0 || this.columnWidths.length <= 1) return false;
    for (const row of this.rows) row.cells.splice(index, 1);
    this.columnWidths.splice(index, 1);
    return true;
  }
  MergeCells(cells) {
    if (cells.length < 2) return null;
    this.mergedCells = cells;
    return cells[0];
  }
  Delete() {
    if (!this.parent) return false;
    return this.parent.RemoveObject(this);
  }
  ToJSON() {
    return JSON.stringify({
      kind: "table",
      name: this.name,
      rows: this.rows.length,
      columns: Math.max(...this.rows.map(row => row.cells.length)),
    });
  }
}

class MockGroup {
  constructor(drawings) {
    this.drawings = drawings;
    this.name = "";
    this.internalId = `drawing-${++drawingId}`;
    this.parent = null;
  }
  GetClassType() { return "group"; }
  GetInternalId() { return this.internalId; }
  GetName() { return this.name; }
  SetName(value) { this.name = value; }
  GetPosX() { return Math.min(...this.drawings.map(drawing => drawing.GetPosX())); }
  GetPosY() { return Math.min(...this.drawings.map(drawing => drawing.GetPosY())); }
  GetWidth() {
    const left = this.GetPosX();
    return Math.max(...this.drawings.map(drawing => drawing.GetPosX() + drawing.GetWidth())) - left;
  }
  GetHeight() {
    const top = this.GetPosY();
    return Math.max(...this.drawings.map(drawing => drawing.GetPosY() + drawing.GetHeight())) - top;
  }
  GetRotation() { return 0; }
  GetFlipH() { return false; }
  GetFlipV() { return false; }
  Ungroup() {
    if (!this.parent) return false;
    const slide = this.parent;
    slide.RemoveObject(this);
    for (const drawing of this.drawings) slide.AddObject(drawing);
    return true;
  }
  ToJSON() { return JSON.stringify({ kind: "group", name: this.name, count: this.drawings.length }); }
}

class MockCustomPath {
  constructor() {
    this.commands = [];
  }
  SetWidth(value) { this.width = value; }
  SetHeight(value) { this.height = value; }
  SetStroke(value) { this.stroke = value; }
  SetFill(value) { this.fill = value; }
  MoveTo(...values) { this.commands.push(["moveTo", ...values]); }
  LineTo(...values) { this.commands.push(["lineTo", ...values]); }
  QuadBezTo(...values) { this.commands.push(["quadBezTo", ...values]); }
  CubicBezTo(...values) { this.commands.push(["cubicBezTo", ...values]); }
  ArcTo(...values) { this.commands.push(["arcTo", ...values]); }
  Close() { this.commands.push(["close"]); }
}

class MockCustomGeometry {
  constructor() {
    this.paths = [];
  }
  AddPath() {
    const path = new MockCustomPath();
    this.paths.push(path);
    return path;
  }
  GetPreset() { return null; }
  IsCustom() { return true; }
}

class MockAnimationEffect {
  constructor(sequence, shape, effectType, trigger) {
    this.sequence = sequence;
    this.shape = shape;
    this.effectType = effectType;
    this.trigger = trigger;
    this.duration = 500;
    this.delay = 0;
    this.repeatCount = 1;
  }
  GetShape() { return this.shape; }
  GetEffectType() { return this.effectType; }
  GetTriggerType() { return this.trigger; }
  SetTriggerType(value) { this.trigger = value; return true; }
  GetDuration() { return this.duration; }
  SetDuration(value) { this.duration = value; return true; }
  GetDelay() { return this.delay; }
  SetDelay(value) { this.delay = value; return true; }
  GetRepeatCount() { return this.repeatCount; }
  SetRepeatCount(value) { this.repeatCount = value; return true; }
  MoveTo(index) {
    const current = this.sequence.effects.indexOf(this);
    if (current < 0 || index < 0) return false;
    this.sequence.effects.splice(current, 1);
    this.sequence.effects.splice(Math.min(index, this.sequence.effects.length), 0, this);
    return true;
  }
  Delete() {
    const index = this.sequence.effects.indexOf(this);
    if (index < 0) return false;
    this.sequence.effects.splice(index, 1);
    return true;
  }
}

class MockAnimationSequence {
  constructor() { this.effects = []; }
  AddEffect(shape, effectType, trigger) {
    const effect = new MockAnimationEffect(this, shape, effectType, trigger);
    this.effects.push(effect);
    return effect;
  }
  GetCount() { return this.effects.length; }
  RemoveAllEffects() { this.effects.splice(0); return true; }
}

class MockTimeline {
  constructor() {
    this.main = new MockAnimationSequence();
    this.interactive = [];
  }
  GetMainSequence() { return this.main; }
  AddInteractiveSequence(trigger) {
    let entry = this.interactive.find(item => item.trigger === trigger);
    if (!entry) {
      entry = { trigger, sequence: new MockAnimationSequence() };
      this.interactive.push(entry);
    }
    return entry.sequence;
  }
  GetInteractiveSequences() { return this.interactive.map(item => item.sequence); }
  GetAllEffects() {
    return [
      ...this.main.effects,
      ...this.interactive.flatMap(item => item.sequence.effects),
    ];
  }
}

let commentId = 0;

class MockComment {
  constructor(presentation, x, y, text, author = "", userId = "", parent = null) {
    this.presentation = presentation;
    this.position = [x, y];
    this.text = text;
    this.author = author;
    this.userId = userId;
    this.parent = parent;
    this.id = `comment-${++commentId}`;
    this.solved = false;
    this.replies = [];
  }
  GetId() { return this.id; }
  GetText() { return this.text; }
  SetText(value) { this.text = value; return true; }
  GetAuthorName() { return this.author; }
  SetAuthorName(value) { this.author = value; return true; }
  GetUserId() { return this.userId; }
  SetUserId(value) { this.userId = value; return true; }
  GetPosition() { return this.position; }
  SetPosition(x, y) { this.position = [x, y]; return true; }
  GetTime() { return "2026-07-23T00:00:00"; }
  GetTimeUTC() { return "2026-07-22T16:00:00Z"; }
  IsSolved() { return this.solved; }
  SetSolved(value) { this.solved = value; return true; }
  GetRepliesCount() { return this.replies.length; }
  GetReply(index) { return this.replies[index] || null; }
  AddReply(text, author, userId) {
    const reply = new MockComment(this.presentation, 0, 0, text, author, userId, this);
    this.replies.push(reply);
    return reply;
  }
  RemoveReplies(start, count) {
    this.replies.splice(start, count);
    return true;
  }
  Delete() {
    const collection = this.parent ? this.parent.replies : this.presentation.comments;
    const index = collection.indexOf(this);
    if (index < 0) return false;
    collection.splice(index, 1);
    return true;
  }
}

class MockSlide {
  constructor(presentation, texts = []) {
    this.presentation = presentation;
    this.shapes = texts.map(text => new SlideShape(text));
    this.shapes.forEach(shape => { shape.parent = this; });
    this.charts = [];
    this.images = [];
    this.tables = [];
    this.groups = [];
    this.background = null;
    this.visible = true;
    this.layout = presentation.masters[0].GetLayout(0);
    this.notesBody = new SlideShape();
    this.transition = null;
    this.timeline = new MockTimeline();
  }
  GetAllShapes() { return this.shapes; }
  GetAllCharts() { return this.charts; }
  GetAllImages() { return this.images; }
  GetAllTables() { return this.tables; }
  GetAllDrawings() { return [...this.shapes, ...this.charts, ...this.images, ...this.tables, ...this.groups]; }
  GetWidth() { return this.presentation.width; }
  GetHeight() { return this.presentation.height; }
  GetVisible() { return this.visible; }
  SetVisible(value) { this.visible = Boolean(value); return true; }
  GetLayout() { return this.layout; }
  ApplyLayout(layout) {
    for (const shape of this.shapes.slice()) {
      if (shape.GetPlaceholder()) this.RemoveObject(shape);
    }
    this.layout = layout;
    for (const source of layout.GetAllDrawings()) {
      if (!source.GetPlaceholder()) continue;
      const clone = new SlideShape(source.GetContent().GetText());
      clone.position = source.position.slice();
      clone.size = source.size.slice();
      clone.geometry.preset = source.geometry.preset;
      clone.placeholder = source.placeholder;
      clone.name = source.name;
      this.AddObject(clone);
    }
    return true;
  }
  GetDrawingsByPlaceholderType(type) {
    return this.GetAllDrawings().filter(drawing => (
      drawing.GetPlaceholder()
      && drawing.GetPlaceholder().GetType() === type
    ));
  }
  GetTheme() { return this.theme || this.layout.GetMaster().GetTheme(); }
  ApplyTheme(theme) { this.theme = theme; return true; }
  GetNotesPage() {
    return {
      GetBodyShape: () => this.notesBody,
      GetBodyShapeText: () => this.notesBody.GetContent().GetText(),
    };
  }
  AddNotesText(text) {
    const paragraph = new SlideParagraph(String(text));
    this.notesBody.GetContent().Push(paragraph);
    return true;
  }
  AddComment(x, y, text, author, userId) {
    const comment = new MockComment(this.presentation, x, y, text, author, userId);
    this.presentation.comments.push(comment);
    return comment;
  }
  GetTimeLine() { return this.timeline; }
  Select() { this.presentation.selectedSlide = this; return true; }
  GetSlideShowTransition() { return this.transition; }
  SetSlideShowTransition(value) { this.transition = value; return true; }
  MoveTo(index) {
    const current = this.presentation.slides.indexOf(this);
    if (current < 0 || index < 0 || index >= this.presentation.slides.length) return false;
    this.presentation.slides.splice(current, 1);
    this.presentation.slides.splice(index, 0, this);
    return true;
  }
  GroupDrawings(drawings) {
    for (const drawing of drawings) {
      if (!this.RemoveObject(drawing)) return null;
    }
    const group = new MockGroup(drawings);
    this.AddObject(group);
    return group;
  }
  AddObject(shape) {
    shape.parent = this;
    if (shape.GetClassType() === "chart") this.charts.push(shape);
    else if (shape.GetClassType() === "image") this.images.push(shape);
    else if (shape.GetClassType() === "table") this.tables.push(shape);
    else if (shape.GetClassType() === "group") this.groups.push(shape);
    else this.shapes.push(shape);
  }
  RemoveObject(shape) {
    const collection = shape.GetClassType() === "chart"
      ? this.charts
      : shape.GetClassType() === "image"
        ? this.images
        : shape.GetClassType() === "table"
          ? this.tables
          : shape.GetClassType() === "group" ? this.groups : this.shapes;
    const index = collection.indexOf(shape);
    if (index < 0) return false;
    collection.splice(index, 1);
    this.lastRemovedObject = shape;
    shape.parent = null;
    return true;
  }
  SetBackground(fill) { this.background = fill; }
  GetBackground() { return this.background; }
  ClearBackground() { this.background = null; }
  FollowLayoutBackground() { this.background = "layout"; }
  FollowMasterBackground() { this.background = "master"; }
  ToJSON() {
    return JSON.stringify({
      shapes: this.shapes.length,
      charts: this.charts.length,
      images: this.images.length,
      tables: this.tables.length,
      groups: this.groups.length,
    });
  }
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
  const layouts = [
    new MockLayout("Title Slide", "title"),
    new MockLayout("Title and Content", "obj"),
  ];
  const centeredTitle = new SlideShape();
  centeredTitle.SetName("Centered Title");
  centeredTitle.SetPlaceholder({ type: "ctrTitle", GetType() { return this.type; } });
  layouts[0].AddObject(centeredTitle);
  const subtitle = new SlideShape();
  subtitle.SetName("Subtitle");
  subtitle.SetPlaceholder({ type: "subtitle", GetType() { return this.type; } });
  layouts[0].AddObject(subtitle);
  const title = new SlideShape();
  title.SetName("Title");
  title.SetPlaceholder({ type: "title", GetType() { return this.type; } });
  layouts[1].AddObject(title);
  const body = new SlideShape();
  body.SetName("Content");
  body.SetPlaceholder({ type: "body", GetType() { return this.type; } });
  layouts[1].AddObject(body);
  const presentation = {
    slides: [],
    masters: [new MockMaster(layouts)],
    width: 9144000,
    height: 5143500,
    loop: false,
    comments: [],
    pluginCalls: [],
    historyPoints: 0,
    GetSlidesCount() { return this.slides.length; },
    GetAllSlides() { return this.slides; },
    GetSlideByIndex(index) { return this.slides[index]; },
    GetCurSlideIndex() { return 0; },
    GetWidth() { return this.width; },
    GetHeight() { return this.height; },
    SetSizes(width, height) { this.width = width; this.height = height; return true; },
    GetMastersCount() { return this.masters.length; },
    GetAllSlideMasters() { return this.masters; },
    GetMaster(index) { return this.masters[index] || null; },
    GetLoopUntilStopped() { return this.loop; },
    SetLoopUntilStopped(value) { this.loop = Boolean(value); return true; },
    ApplyTheme(theme) {
      this.appliedTheme = theme;
      this.slides.forEach(slide => { slide.theme = theme; });
      return true;
    },
    GetAllComments() { return this.comments; },
    AddMathEquation(text, format) {
      const slide = this.selectedSlide || this.slides[this.GetCurSlideIndex()];
      const equation = new SlideShape(text);
      equation.mathFormat = format;
      slide.AddObject(equation);
      return true;
    },
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
    CreateBlipFill: (url, mode) => mockFill("blip", { url, mode }),
    CreateStroke: (width, fill) => mockStroke(width, fill),
    CreateShape: (type, width, height, fill, line) => {
      const shape = new SlideShape();
      shape.geometry.preset = type;
      shape.size = [width, height];
      shape.fill = fill;
      shape.line = line;
      return shape;
    },
    CreatePlaceholder: type => ({ type, GetType() { return this.type; } }),
    CreateLayout: master => {
      const layout = new MockLayout("", "custom", master);
      master.layouts.push(layout);
      return layout;
    },
    CreateThemeColorScheme: (colors, name) => ({
      colors,
      name,
      ToJSON() { return JSON.stringify({ colors: this.colors, name: this.name }); },
    }),
    CreateThemeFontScheme: (majorLatin, majorEastAsian, majorComplex, minorLatin, minorEastAsian, minorComplex, name) => ({
      majorLatin,
      majorEastAsian,
      majorComplex,
      minorLatin,
      minorEastAsian,
      minorComplex,
      name,
      ToJSON() {
        return JSON.stringify({
          majorLatin: this.majorLatin,
          majorEastAsian: this.majorEastAsian,
          majorComplex: this.majorComplex,
          minorLatin: this.minorLatin,
          minorEastAsian: this.minorEastAsian,
          minorComplex: this.minorComplex,
          name: this.name,
        });
      },
    }),
    CreatePresetGeometry: type => ({ preset: type, GetPreset() { return this.preset; } }),
    CreateBullet: symbol => ({ kind: "bullet", symbol }),
    CreateNumbering: (type, startAt) => ({ kind: "number", type, startAt }),
    CreateHyperlink: (link, tooltip) => ({ link, tooltip }),
    CreateSlideShowTransition: () => ({
      effect: null,
      speed: null,
      duration: null,
      advanceOnClick: null,
      advanceOnTime: null,
      advanceTime: null,
      SetEntryEffect(value) { this.effect = value; return true; },
      GetEntryEffect() { return this.effect; },
      SetSpeed(value) { this.speed = value; return true; },
      GetSpeed() { return this.speed; },
      SetDuration(value) { this.duration = value; return true; },
      GetDuration() { return this.duration; },
      SetAdvanceOnClick(value) { this.advanceOnClick = value; return true; },
      GetAdvanceOnClick() { return this.advanceOnClick; },
      SetAdvanceOnTime(value) { this.advanceOnTime = value; return true; },
      GetAdvanceOnTime() { return this.advanceOnTime; },
      SetAdvanceTime(value) { this.advanceTime = value; return true; },
      GetAdvanceTime() { return this.advanceTime; },
    }),
    CreateChart: (...args) => new SlideChart(...args),
    CreateImage: (url, width, height) => new SlideImage(url, width, height),
    CreateTextPr: () => ({
      SetFontSize(value) { this.fontSize = value; },
      SetFontFamily(value) { this.fontFamily = value; },
      SetBold(value) { this.bold = value; },
      SetItalic(value) { this.italic = value; },
      SetUnderline(value) { this.underline = value; },
      SetCaps(value) { this.caps = value; },
      SetColor(value) { this.color = value; },
    }),
    CreateWordArt: (_textPr, text, transform, fill, line, rotation, width, height) => {
      const shape = new SlideShape(text);
      shape.wordArtTransform = transform;
      shape.fill = fill;
      shape.line = line;
      shape.rotation = rotation;
      shape.size = [width, height];
      return shape;
    },
    CreateOleObject: (...args) => new MockOleObject(...args),
    CreateTable: (rows, columns) => new MockTable(rows, columns),
    CreateCustomGeometry: () => new MockCustomGeometry(),
    FromJSON: raw => {
      if (typeof raw === "string") raw = JSON.parse(raw);
      if (!raw || typeof raw !== "object") return null;
      if (raw.kind === "fill") return mockFill(raw.type, raw);
      if (raw.kind === "line") return mockStroke(raw.width, mockFill(raw.fill.type, raw.fill));
      return null;
    },
    CreateParagraph: () => new SlideParagraph(),
    CreateSlide: () => new MockSlide(presentation),
    __executeMethod(method, params, callback) {
      presentation.pluginCalls.push([method, params]);
      const values = {
        GetEditorThemes: JSON.stringify([{ id: 7, name: "Ion" }]),
        ApplyTheme: true,
        GetMacros: JSON.stringify({ macrosArray: [] }),
        GetVBAMacros: "Sub Main()\nEnd Sub",
        SetMacros: true,
        StartSlideShow: true,
        EndSlideShow: true,
        PauseSlideShow: true,
        ResumeSlideShow: true,
        GoToNextSlideInSlideShow: true,
        GoToPreviousSlideInSlideShow: true,
        GoToSlideInSlideShow: true,
      };
      callback(values[method]);
    },
  };
  return { bridge: loadBridge(slidesSource, api).slide, presentation, first, second, selection, api };
}

class MockRange {
  constructor(address, values) {
    this.address = address;
    this.values = values;
    this.formula = null;
    this.format = {};
    this.setValueCalls = [];
    this.clearContentsCalls = 0;
  }
  GetAddress() { return this.address; }
  GetValue() { return this.values; }
  GetRowsCount() {
    const match = String(this.address).match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i);
    return match ? Math.abs(Number(match[4] || match[2]) - Number(match[2])) + 1 : 1;
  }
  GetColumnsCount() {
    const match = String(this.address).match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i);
    if (!match) return 1;
    const columnNumber = letters => [...letters.toUpperCase()]
      .reduce((value, letter) => (value * 26) + letter.charCodeAt(0) - 64, 0);
    return Math.abs(columnNumber(match[3] || match[1]) - columnNumber(match[1])) + 1;
  }
  SetValue(values) { this.setValueCalls.push(values); this.values = values; return true; }
  ClearContents() { this.clearContentsCalls += 1; this.values = []; return true; }
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

test("Slides bridge adds a centered proportional image and exposes it to inspect/delete", async () => {
  const { bridge, presentation, first } = slidesHarness();
  const asset = {
    url: "https://app.test/copilot-api/images/asset.png?token=signed",
    assetId: "asset.png",
    widthPx: 800,
    heightPx: 600,
  };
  const result = await bridge.execute([
    {
      name: "slides_add_image",
      arguments: {
        slide: 1,
        _image: asset,
        rotationDeg: 15,
        flipH: true,
        name: "产品截图",
      },
    },
    { name: "slides_inspect_objects", arguments: { slide: 1 } },
    { name: "slides_delete_object", arguments: { slide: 1, name: "产品截图" } },
  ]);

  const added = result.results[0].object;
  assert.equal(result.changed, 2);
  assert.equal(result.needsSave, true);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(added.kind, "image");
  assert.equal(added.name, "产品截图");
  assert.equal(added.widthMm, 120);
  assert.equal(added.heightMm, 90);
  assert.ok(Math.abs(added.xMm - 67) < 0.001);
  assert.ok(Math.abs(added.yMm - 26.4375) < 0.001);
  assert.equal(added.rotationDeg, 15);
  assert.equal(added.flipH, true);
  assert.equal(result.results[1].slides[0].objects.at(-1).kind, "image");
  assert.equal(first.images.length, 0);
});

test("Slides bridge derives one dimension and only stretches when ratio preservation is disabled", async () => {
  const { bridge, first } = slidesHarness();
  const asset = {
    url: "https://app.test/copilot-api/images/asset.png?token=signed",
    widthPx: 400,
    heightPx: 200,
  };
  const result = await bridge.execute([
    {
      name: "slides_add_image",
      arguments: {
        slide: 1,
        _image: asset,
        widthMm: 80,
        xMm: 10,
        yMm: 20,
      },
    },
    {
      name: "slides_add_image",
      arguments: {
        slide: 1,
        _image: asset,
        widthMm: 80,
        heightMm: 80,
        preserveAspectRatio: false,
        flipV: true,
      },
    },
  ]);

  assert.equal(result.results[0].object.widthMm, 80);
  assert.equal(result.results[0].object.heightMm, 40);
  assert.equal(result.results[0].object.xMm, 10);
  assert.equal(result.results[0].object.yMm, 20);
  assert.equal(result.results[1].object.widthMm, 80);
  assert.equal(result.results[1].object.heightMm, 80);
  assert.equal(result.results[1].object.flipV, true);
  assert.equal(first.images.length, 2);
});

test("Slides bridge crops an image to a preset shape and applies a border", async () => {
  const { bridge, first } = slidesHarness();
  const asset = {
    url: "https://app.test/copilot-api/images/vector.svg?token=signed",
    assetId: "vector.svg",
    widthPx: 400,
    heightPx: 200,
  };
  const result = await bridge.execute([
    {
      name: "slides_add_image_shape",
      arguments: {
        slide: 1,
        _image: asset,
        shapeType: "ellipse",
        widthMm: 60,
        heightMm: 60,
        preserveAspectRatio: false,
        line: { widthPt: 2, color: "#336699" },
        name: "Avatar",
      },
    },
    {
      name: "slides_update_object",
      arguments: {
        slide: 1,
        name: "Avatar",
        line: { widthPt: 3, color: "#AA3300" },
      },
    },
  ]);

  assert.equal(result.changed, 2);
  assert.equal(result.results[0].object.kind, "shape");
  assert.equal(result.results[0].object.shapeType, "ellipse");
  assert.equal(result.results[0].object.fill.type, "blip");
  assert.equal(result.results[0].object.widthMm, 60);
  assert.equal(result.results[0].object.heightMm, 60);
  assert.equal(result.results[1].object.line.widthPt, 3);
  assert.equal(first.shapes.at(-1).GetName(), "Avatar");
});

test("Slides bridge sets native stretch and tile image backgrounds without adding drawings", async () => {
  const { bridge, presentation, first, second } = slidesHarness();
  const firstDrawingCount = first.GetAllDrawings().length;
  const secondDrawingCount = second.GetAllDrawings().length;
  const firstAsset = {
    url: "https://app.test/copilot-api/images/background-one.png?token=signed",
    assetId: "background-one.png",
    widthPx: 1600,
    heightPx: 900,
  };
  const secondAsset = {
    url: "https://app.test/copilot-api/images/background-two.jpg?token=signed",
    assetId: "background-two.jpg",
    widthPx: 1600,
    heightPx: 900,
  };

  const result = await bridge.execute([
    {
      name: "slides_set_background",
      arguments: { slide: 1, mode: "image", _image: firstAsset },
    },
    {
      name: "slides_set_background",
      arguments: { slide: 2, mode: "image", fillMode: "tile", _image: secondAsset },
    },
    {
      name: "slides_inspect_backgrounds",
      arguments: { includeTemplates: true },
    },
  ]);

  assert.equal(result.changed, 2);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(first.background.type, "blip");
  assert.equal(first.background.url, firstAsset.url);
  assert.equal(first.background.mode, "stretch");
  assert.equal(second.background.type, "blip");
  assert.equal(second.background.url, secondAsset.url);
  assert.equal(second.background.mode, "tile");
  assert.equal(first.GetAllDrawings().length, firstDrawingCount);
  assert.equal(second.GetAllDrawings().length, secondDrawingCount);
  assert.deepEqual(result.results[0].source, {
    assetId: "background-one.png",
    widthPx: 1600,
    heightPx: 900,
  });
  assert.equal(result.results[0].fillMode, "stretch");
  assert.equal(result.results[1].fillMode, "tile");
  assert.equal(result.results[2].slides[0].background.available, true);
  assert.equal(result.results[2].slides[0].background.fill.type, "blip");
  assert.equal(result.results[2].slides[1].background.fill.type, "blip");
  assert.equal(result.results[2].masters.length, 1);
  assert.equal(JSON.stringify(result.results).includes("token=signed"), false);
});

test("Slides bridge reads backgrounds through the installed SDK model fallback", async () => {
  const { bridge, first } = slidesHarness();
  first.GetBackground = undefined;
  first.GetBackgroundFill = undefined;
  first.Slide = {
    cSld: {
      Bg: {
        bgPr: {
          Fill: mockFill("blip", { url: "data:image/jpeg;base64,not-returned" }),
        },
      },
    },
  };

  const result = await bridge.execute([{
    name: "slides_inspect_backgrounds",
    arguments: { slide: 1, includeRaw: false },
  }]);
  const background = result.results[0].slides[0].background;

  assert.equal(background.available, true);
  assert.equal(background.source, "slide");
  assert.equal(background.inherited, false);
  assert.equal(background.hasDirectBackground, true);
  assert.equal(background.fill.type, "blip");
  assert.equal(Object.hasOwn(background.fill, "raw"), false);
  assert.equal(JSON.stringify(background).includes("base64"), false);
});

test("Slides bridge sets native image backgrounds on masters and layouts", async () => {
  const { bridge, presentation } = slidesHarness();
  const master = presentation.GetMaster(0);
  const layout = master.GetLayout(1);
  const masterDrawingCount = master.GetAllDrawings().length;
  const layoutDrawingCount = layout.GetAllDrawings().length;
  const asset = {
    url: "https://app.test/copilot-api/images/template-background.png?token=signed",
    assetId: "template-background.png",
    widthPx: 1600,
    heightPx: 900,
  };

  const result = await bridge.execute([
    {
      name: "slides_set_template_background",
      arguments: { scope: "master", masterIndex: 1, mode: "image", _image: asset },
    },
    {
      name: "slides_set_template_background",
      arguments: {
        scope: "layout",
        masterIndex: 1,
        layoutIndex: 2,
        mode: "image",
        fillMode: "tile",
        _image: asset,
      },
    },
  ]);

  assert.equal(result.changed, 2);
  assert.equal(master.background.type, "blip");
  assert.equal(master.background.mode, "stretch");
  assert.equal(layout.background.type, "blip");
  assert.equal(layout.background.mode, "tile");
  assert.equal(master.GetAllDrawings().length, masterDrawingCount);
  assert.equal(layout.GetAllDrawings().length, layoutDrawingCount);
  assert.equal(result.results[0].scope, "master");
  assert.equal(result.results[1].scope, "layout");
  assert.equal(result.results[1].fillMode, "tile");
});

test("Slides bridge reports unsupported or rejected native image backgrounds without adding drawings", async () => {
  const unsupported = slidesHarness();
  const unsupportedDrawingCount = unsupported.first.GetAllDrawings().length;
  unsupported.api.CreateBlipFill = undefined;
  await assert.rejects(
    unsupported.bridge.execute([{
      name: "slides_set_background",
      arguments: {
        slide: 1,
        mode: "image",
        _image: {
          url: "https://app.test/copilot-api/images/background.png?token=signed",
          widthPx: 1600,
          heightPx: 900,
        },
      },
    }]),
    /不支持图片背景填充/,
  );
  assert.equal(unsupported.first.background, null);
  assert.equal(unsupported.first.GetAllDrawings().length, unsupportedDrawingCount);

  const rejected = slidesHarness();
  const rejectedDrawingCount = rejected.first.GetAllDrawings().length;
  rejected.first.SetBackground = () => false;
  await assert.rejects(
    rejected.bridge.execute([{
      name: "slides_set_background",
      arguments: {
        slide: 1,
        mode: "image",
        _image: {
          url: "https://app.test/copilot-api/images/background.png?token=signed",
          widthPx: 1600,
          heightPx: 900,
        },
      },
    }]),
    /拒绝设置幻灯片图片背景/,
  );
  assert.equal(rejected.first.background, null);
  assert.equal(rejected.first.GetAllDrawings().length, rejectedDrawingCount);
});

test("Slides bridge rejects an invalid slide before creating an image", async () => {
  const { bridge, first } = slidesHarness();
  await assert.rejects(
    bridge.execute([{
      name: "slides_add_image",
      arguments: {
        slide: 99,
        _image: {
          url: "https://app.test/copilot-api/images/asset.png?token=signed",
          widthPx: 100,
          heightPx: 100,
        },
      },
    }]),
    /页码超出范围/,
  );
  assert.equal(first.images.length, 0);
});

test("Slides bridge inspects layouts and changes order, visibility, size, layout, and loop settings", async () => {
  const { bridge, presentation, first, second } = slidesHarness();
  const result = await bridge.execute([
    { name: "slides_inspect_layouts", arguments: {} },
    { name: "slides_apply_layout", arguments: { slide: 1, masterIndex: 1, layoutIndex: 2 } },
    { name: "slides_set_visibility", arguments: { slide: 2, visible: false } },
    { name: "slides_move_slide", arguments: { slide: 2, toIndex: 1 } },
    {
      name: "slides_set_size",
      arguments: { preset: "custom", widthMm: 300, heightMm: 170, orientation: "landscape" },
    },
    { name: "slides_set_show_settings", arguments: { loop: true } },
    { name: "slides_inspect", arguments: {} },
  ]);

  assert.equal(result.changed, 5);
  assert.equal(result.needsSave, true);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(result.results[0].masters[0].layouts[1].name, "Title and Content");
  assert.equal(first.GetLayout().GetName(), "Title and Content");
  assert.equal(second.GetVisible(), false);
  assert.equal(presentation.slides[0], second);
  assert.equal(Math.round(presentation.GetWidth() / 36000), 300);
  assert.equal(Math.round(presentation.GetHeight() / 36000), 170);
  assert.equal(presentation.GetLoopUntilStopped(), true);
  assert.equal(result.results[6].slides[0].visible, false);
  assert.equal(result.results[6].loopUntilStopped, true);
});

test("Slides bridge atomically applies a layout and creates native rich text", async () => {
  const { bridge, presentation } = slidesHarness();
  const result = await bridge.execute([
    {
      name: "slides_add_slide",
      arguments: { masterIndex: 1, layoutIndex: 2, title: "经营概览", titleFontSize: 32 },
    },
    {
      name: "slides_add_textbox",
      arguments: {
        slide: 3,
        name: "s3-actions",
        xMm: 20,
        yMm: 55,
        widthMm: 100,
        heightMm: 70,
        paragraphs: [
          { text: "优化产品结构", listType: "bullet", fontSize: 20 },
          { text: "提升交付效率", listType: "number", numberingType: "ArabicPeriod", fontSize: 18 },
        ],
      },
    },
    {
      name: "slides_inspect_objects",
      arguments: { slide: 3, includeTextStyles: true },
    },
  ]);

  assert.equal(result.results[0].slide, 3);
  assert.equal(result.results[0].layout.layoutIndex, 2);
  assert.equal(result.results[0].titleSource, "placeholder");
  assert.equal(result.results[0].titleObject.text, "经营概览");
  assert.equal(result.results[0].titleObject.placeholderType, "title");
  assert.equal(presentation.GetSlideByIndex(2).GetLayout().GetName(), "Title and Content");
  assert.equal(
    presentation.GetSlideByIndex(2).GetDrawingsByPlaceholderType("title").length,
    1,
  );
  assert.equal(
    presentation.GetSlideByIndex(2)
      .GetDrawingsByPlaceholderType("title")[0]
      .GetContent().GetAllParagraphs()[0].GetElement(0).GetFontSize(),
    64,
  );
  assert.equal(
    presentation.GetSlideByIndex(2).GetDrawingsByPlaceholderType("body").length,
    1,
  );
  const added = result.results[1].object;
  assert.equal(added.name, "s3-actions");
  const inspected = result.results[2].slides[0].objects.find(item => item.name === "s3-actions");
  assert.equal(inspected.textStyles.paragraphCount, 2);
  assert.equal(inspected.textStyles.minFontSize, 18);
  assert.equal(inspected.textStyles.paragraphs[0].list.type, "bullet");
  assert.equal(inspected.textStyles.paragraphs[1].list.type, "number");
});

test("Slides bridge supports automatic, strict, and legacy title placement", async () => {
  const automaticHarness = slidesHarness();
  const centered = await automaticHarness.bridge.execute([{
    name: "slides_add_slide",
    arguments: {
      masterIndex: 1,
      layoutIndex: 1,
      title: "Centered",
    },
  }]);
  assert.equal(centered.results[0].titleSource, "placeholder");
  assert.equal(centered.results[0].titleObject.placeholderType, "ctrTitle");
  assert.equal(
    automaticHarness.presentation.GetSlideByIndex(2)
      .GetDrawingsByPlaceholderType("ctrTitle")[0]
      .GetContent().GetAllParagraphs()[0].GetElement(0).GetFontSize(),
    20,
  );
  assert.equal(
    automaticHarness.presentation.GetSlideByIndex(2)
      .GetDrawingsByPlaceholderType("subtitle").length,
    1,
  );

  const titleOnlyLayout = new MockLayout("Title Only", "titleOnly");
  const titleOnlyPlaceholder = new SlideShape();
  titleOnlyPlaceholder.SetName("Title Only Placeholder");
  titleOnlyPlaceholder.SetPlaceholder({
    type: "title",
    GetType() { return this.type; },
  });
  titleOnlyLayout.AddObject(titleOnlyPlaceholder);
  const customLayout = new MockLayout("Custom", "custom");
  const customTitle = new SlideShape();
  customTitle.SetName("Custom Title");
  customTitle.SetPlaceholder({ type: "title", GetType() { return this.type; } });
  customLayout.AddObject(customTitle);
  const customBody = new SlideShape();
  customBody.SetName("Custom Body");
  customBody.SetPlaceholder({ type: "body", GetType() { return this.type; } });
  customLayout.AddObject(customBody);
  const automaticMaster = automaticHarness.presentation.GetMaster(0);
  titleOnlyLayout.master = automaticMaster;
  customLayout.master = automaticMaster;
  automaticMaster.layouts.push(
    titleOnlyLayout,
    customLayout,
  );

  const titleOnlyResult = await automaticHarness.bridge.execute([{
    name: "slides_add_slide",
    arguments: {
      masterIndex: 1,
      layoutIndex: 3,
      title: "Title only",
    },
  }]);
  assert.equal(titleOnlyResult.results[0].titleSource, "placeholder");
  assert.equal(titleOnlyResult.results[0].titleObject.placeholderType, "title");

  const customResult = await automaticHarness.bridge.execute([{
    name: "slides_add_slide",
    arguments: {
      masterIndex: 1,
      layoutIndex: 4,
      title: "Custom",
    },
  }]);
  assert.equal(customResult.results[0].titleSource, "placeholder");
  const customSlide = automaticHarness.presentation.GetSlideByIndex(4);
  assert.equal(customSlide.GetDrawingsByPlaceholderType("title").length, 1);
  assert.equal(customSlide.GetDrawingsByPlaceholderType("body").length, 1);

  const untitledResult = await automaticHarness.bridge.execute([{
    name: "slides_add_slide",
    arguments: {
      masterIndex: 1,
      layoutIndex: 2,
    },
  }]);
  assert.equal(untitledResult.results[0].titleSource, null);
  assert.equal(untitledResult.results[0].titleObject, null);
  const untitledSlide = automaticHarness.presentation.GetSlideByIndex(5);
  assert.equal(untitledSlide.GetAllShapes().length, 2);

  const fallback = await automaticHarness.bridge.execute([{
    name: "slides_add_slide",
    arguments: { title: "Freeform" },
  }]);
  assert.equal(fallback.results[0].titleSource, "textbox");
  assert.equal(fallback.results[0].titleObject.placeholderType, null);

  const inspected = await automaticHarness.bridge.execute([
    {
      name: "slides_inspect_objects",
      arguments: { slide: 6 },
    },
    {
      name: "slides_inspect_objects",
      arguments: { slide: 7 },
    },
  ]);
  assert.deepEqual(
    inspected.results[0].slides[0].objects.map(item => item.placeholderType),
    ["title", "body"],
  );
  assert.equal(inspected.results[1].slides[0].objects[0].placeholderType, null);

  const legacyHarness = slidesHarness();
  const legacy = await legacyHarness.bridge.execute([{
    name: "slides_add_slide",
    arguments: {
      masterIndex: 1,
      layoutIndex: 2,
      title: "Legacy",
      titlePlacement: "textbox",
    },
  }]);
  const legacySlide = legacyHarness.presentation.GetSlideByIndex(2);
  assert.equal(legacy.results[0].titleSource, "textbox");
  assert.equal(legacySlide.GetDrawingsByPlaceholderType("title").length, 1);
  assert.equal(legacySlide.GetDrawingsByPlaceholderType("body").length, 1);
  assert.equal(legacySlide.GetAllShapes().length, 3);

  const strictHarness = slidesHarness();
  await assert.rejects(
    strictHarness.bridge.execute([{
      name: "slides_add_slide",
      arguments: {
        title: "Strict",
        titlePlacement: "placeholder",
      },
    }]),
    error => (
      error.code === "TITLE_PLACEHOLDER_NOT_FOUND"
      && /没有可用的标题占位符/.test(error.message)
    ),
  );
  assert.equal(strictHarness.presentation.GetSlidesCount(), 2);
});

test("Slides placeholder title batches stay within the legacy textbox performance budget", async () => {
  const batch = Array.from({ length: 20 }, (_item, index) => ({
    name: "slides_add_slide",
    arguments: {
      masterIndex: 1,
      layoutIndex: 2,
      title: `Title ${index + 1}`,
    },
  }));

  async function sample(titlePlacement) {
    const { bridge } = slidesHarness();
    const calls = batch.map(call => ({
      name: call.name,
      arguments: { ...call.arguments, titlePlacement },
    }));
    const started = process.hrtime.bigint();
    for (let iteration = 0; iteration < 8; iteration += 1) {
      await bridge.execute(calls);
    }
    return Number(process.hrtime.bigint() - started) / 1_000_000;
  }

  const automaticSamples = [];
  const legacySamples = [];
  for (let trial = 0; trial < 7; trial += 1) {
    if (trial % 2 === 0) {
      legacySamples.push(await sample("textbox"));
      automaticSamples.push(await sample("auto"));
    } else {
      automaticSamples.push(await sample("auto"));
      legacySamples.push(await sample("textbox"));
    }
  }
  automaticSamples.sort((left, right) => left - right);
  legacySamples.sort((left, right) => left - right);
  const automaticMedian = automaticSamples[Math.floor(automaticSamples.length / 2)];
  const legacyMedian = legacySamples[Math.floor(legacySamples.length / 2)];
  assert.ok(
    automaticMedian <= legacyMedian * 1.2,
    `auto median ${automaticMedian.toFixed(3)}ms exceeded legacy median ${legacyMedian.toFixed(3)}ms by more than 20%`,
  );
});

test("Slides structural validation distinguishes containment, allowed overlap, and violations", async () => {
  const { bridge, presentation } = slidesHarness();
  await bridge.execute([
    { name: "slides_add_slide", arguments: {} },
    {
      name: "slides_add_shape",
      arguments: {
        slide: 3,
        shapeType: "rect",
        name: "panel",
        xMm: 20,
        yMm: 20,
        widthMm: 80,
        heightMm: 50,
        paragraphs: [{ text: "经营指标", fontSize: 20 }],
      },
    },
    {
      name: "slides_add_textbox",
      arguments: {
        slide: 3,
        name: "inside",
        xMm: 30,
        yMm: 30,
        widthMm: 20,
        heightMm: 10,
        text: "包含关系",
        fontSize: 18,
      },
    },
    {
      name: "slides_add_textbox",
      arguments: {
        slide: 3,
        name: "overlap",
        xMm: 90,
        yMm: 30,
        widthMm: 30,
        heightMm: 20,
        text: "允许交叉",
        fontSize: 18,
      },
    },
    {
      name: "slides_add_textbox",
      arguments: {
        slide: 3,
        name: "outside",
        xMm: 245,
        yMm: 150,
        widthMm: 20,
        heightMm: 20,
        text: "• 手写列表",
        fontSize: 12,
      },
    },
  ]);
  const historyBefore = presentation.historyPoints;
  const result = await bridge.execute([
    {
      name: "slides_validate_layout",
      arguments: {
        slide: 3,
        safeMarginMm: 15,
        minFontSize: 16,
        maxObjects: 4,
        expectedMasterIndex: 1,
        expectedLayoutIndex: 2,
        requireUniqueNames: true,
        allowedOverlapPairs: [{ firstName: "panel", secondName: "overlap" }],
      },
    },
  ]);

  assert.equal(presentation.historyPoints, historyBefore);
  assert.equal(result.changed, 0);
  assert.equal(result.needsSave, false);
  assert.equal(result.results[0].visualVerified, false);
  assert.equal(result.results[0].textOverflowVerified, false);
  const validation = result.results[0].slides[0];
  assert.equal(validation.valid, false);
  assert.ok(validation.intersections.some(item => item.kind === "containment"));
  assert.ok(validation.intersections.some(item => item.kind === "overlap" && item.allowed));
  assert.ok(validation.issues.some(item => item.code === "OBJECT_OUT_OF_BOUNDS"));
  assert.ok(validation.issues.some(item => item.code === "FONT_TOO_SMALL"));
  assert.ok(validation.issues.some(item => item.code === "MANUAL_LIST_PREFIX"));
});

test("Slides bridge inspects and customizes themes, masters, and layouts", async () => {
  const { bridge, presentation, first, second } = slidesHarness();
  const colors = [
    "#111111", "#FFFFFF", "#222222", "#F5F5F5",
    "#0066CC", "#CC3300", "#008866", "#993399",
    "#FF9900", "#33AACC", "#0000EE", "#551A8B",
  ];
  const result = await bridge.execute([
    { name: "slides_inspect_themes", arguments: { includeRaw: true } },
    {
      name: "slides_set_theme",
      arguments: {
        masterIndex: 1,
        colors,
        colorSchemeName: "AI Bridge",
        fonts: {
          majorLatin: "Aptos Display",
          minorLatin: "Aptos",
          majorEastAsian: "思源黑体",
          minorEastAsian: "思源黑体",
          name: "AI Bridge fonts",
        },
      },
    },
    {
      name: "slides_create_layout",
      arguments: {
        masterIndex: 1,
        name: "AI Two Column",
        followMasterBackground: true,
        placeholders: [
          { type: "title", text: "Title", xMm: 15, yMm: 10, widthMm: 220, heightMm: 25 },
          { type: "body", xMm: 15, yMm: 45, widthMm: 110, heightMm: 100 },
        ],
        applyToSlides: [1],
      },
    },
    {
      name: "slides_add_template_shape",
      arguments: {
        scope: "master",
        masterIndex: 1,
        shapeType: "rect",
        name: "Brand bar",
        xMm: 0,
        yMm: 180,
        widthMm: 254,
        heightMm: 10,
        fill: { type: "solid", color: "#0066CC" },
      },
    },
    {
      name: "slides_set_template_background",
      arguments: { scope: "layout", masterIndex: 1, layoutIndex: 3, mode: "master" },
    },
    {
      name: "slides_apply_theme",
      arguments: { sourceMasterIndex: 1, targetSlide: 2 },
    },
    {
      name: "slides_manage_template_object",
      arguments: {
        scope: "master",
        masterIndex: 1,
        action: "update",
        name: "Brand bar",
        newName: "Brand bar updated",
        heightMm: 12,
        fill: { type: "solid", color: "#CC3300" },
      },
    },
    { name: "slides_inspect_layouts", arguments: { includeObjects: true } },
    {
      name: "slides_manage_template_object",
      arguments: {
        scope: "master",
        masterIndex: 1,
        action: "delete",
        name: "Brand bar updated",
      },
    },
  ]);

  const master = presentation.GetMaster(0);
  const layout = master.GetLayout(2);
  assert.equal(result.changed, 7);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(result.results[0].masters[0].theme.raw.name, "Office");
  assert.equal(master.GetTheme().GetColorScheme().name, "AI Bridge");
  assert.equal(master.GetTheme().GetFontScheme().majorLatin, "Aptos Display");
  assert.equal(layout.GetName(), "AI Two Column");
  assert.equal(layout.GetAllDrawings().length, 2);
  assert.equal(layout.GetAllDrawings()[0].placeholder.type, "title");
  assert.equal(layout.background, "master");
  assert.equal(first.GetLayout(), layout);
  assert.equal(result.results[7].masters[0].drawings[0].name, "Brand bar updated");
  assert.equal(result.results[7].masters[0].drawings[0].heightMm, 12);
  assert.equal(master.GetAllDrawings().length, 0);
  assert.equal(second.GetTheme(), master.GetTheme());
});

test("Slides bridge manages rich paragraphs, object links, notes, comments, and transitions", async () => {
  const { bridge, presentation, first } = slidesHarness();
  first.shapes[0].SetName("Agenda");
  const result = await bridge.execute([
    {
      name: "slides_set_text_content",
      arguments: {
        slide: 1,
        name: "Agenda",
        paragraphs: [
          { text: "Agenda", fontSize: 28, bold: true, align: "center" },
          { text: "Discover", listType: "bullet", bulletSymbol: "•", level: 0 },
          {
            text: "Decide",
            listType: "number",
            numberingType: "ArabicPeriod",
            startAt: 1,
            level: 1,
          },
        ],
      },
    },
    {
      name: "slides_format_paragraphs",
      arguments: {
        slide: 1,
        name: "Agenda",
        paragraphIndexes: [2],
        leftIndentMm: 8,
        spacingAfterPt: 6,
        lineSpacing: 1.2,
        lineRule: "auto",
        italic: true,
      },
    },
    {
      name: "slides_update_object",
      arguments: { slide: 1, name: "Agenda", xMm: 20, yMm: 25, widthMm: 180, heightMm: 75, rotationDeg: 3 },
    },
    {
      name: "slides_set_hyperlink",
      arguments: { slide: 1, name: "Agenda", action: "slide", targetSlide: 2, tooltip: "Next topic" },
    },
    { name: "slides_set_notes", arguments: { slide: 1, text: "Mention the decision criteria." } },
    {
      name: "slides_add_comment",
      arguments: { slide: 1, text: "Confirm wording", xMm: 12, yMm: 14, author: "AI Bridge", userId: "ai" },
    },
    {
      name: "slides_set_transition",
      arguments: {
        slide: 1,
        effect: "effectFade",
        speed: "medium",
        durationMs: 600,
        advanceOnClick: true,
        advanceOnTime: true,
        advanceTimeMs: 3500,
      },
    },
    { name: "slides_inspect", arguments: {} },
  ]);

  const paragraphs = first.shapes[0].GetContent().GetAllParagraphs();
  assert.equal(result.changed, 7);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(paragraphs.length, 3);
  assert.equal(paragraphs[0].align, "center");
  assert.equal(paragraphs[1].bullet.symbol, "•");
  assert.equal(paragraphs[1].leftIndent, Math.round(8 * 1440 / 25.4));
  assert.equal(paragraphs[1].spacingAfter[0], 120);
  assert.deepEqual(paragraphs[1].spacingLine, [288, "auto"]);
  assert.equal(paragraphs[1].runs[0].format.italic, true);
  assert.equal(paragraphs[2].bullet.kind, "number");
  assert.equal(first.shapes[0].GetPosX(), 20 * 36000);
  assert.equal(first.shapes[0].GetHyperlink().link, "ppaction://hlinksldjumpslide1");
  assert.equal(first.GetNotesPage().GetBodyShapeText(), "Mention the decision criteria.");
  assert.equal(presentation.comments.length, 1);
  assert.equal(first.GetSlideShowTransition().GetEntryEffect(), "effectFade");
  assert.equal(result.results[7].slides[0].notes, "Mention the decision criteria.");
  assert.equal(result.results[7].slides[0].transition.advanceTimeMs, 3500);
  assert.equal(result.results[7].commentCount, 1);
});

test("Slides bridge creates, edits, merges, and formats tables", async () => {
  const { bridge, presentation, first } = slidesHarness();
  const result = await bridge.execute([
    {
      name: "slides_add_table",
      arguments: {
        slide: 1,
        rows: 2,
        columns: 2,
        name: "Metrics",
        data: [["Metric", "Value"], ["Revenue", "42"]],
        xMm: 20,
        yMm: 35,
        widthMm: 180,
        heightMm: 45,
        header: { bold: true, backgroundColor: "#DDEEFF", align: "center" },
      },
    },
    {
      name: "slides_set_table_cell",
      arguments: {
        slide: 1,
        name: "Metrics",
        row: 2,
        column: 2,
        text: "43",
        backgroundColor: "#FFF2CC",
        verticalAlign: "center",
      },
    },
    {
      name: "slides_edit_table",
      arguments: { slide: 1, name: "Metrics", action: "addRow", row: 1, position: "after" },
    },
    {
      name: "slides_edit_table",
      arguments: { slide: 1, name: "Metrics", action: "addColumn", row: 1, column: 1, position: "after" },
    },
    {
      name: "slides_edit_table",
      arguments: {
        slide: 1,
        name: "Metrics",
        action: "mergeCells",
        rowStart: 1,
        rowEnd: 1,
        columnStart: 1,
        columnEnd: 2,
      },
    },
    {
      name: "slides_edit_table",
      arguments: {
        slide: 1,
        name: "Metrics",
        action: "splitCell",
        row: 2,
        column: 2,
        rows: 2,
        columns: 2,
      },
    },
    {
      name: "slides_format_table",
      arguments: {
        slide: 1,
        name: "Metrics",
        columnWidthsMm: [70, 40, 70],
        rowHeightsMm: [14, 12, 12],
        tableLook: { firstRow: true, horizontalBanding: true },
        border: { widthMm: 0.5, color: "#333333" },
        fontSize: 12,
      },
    },
    { name: "slides_inspect_objects", arguments: { slide: 1, kinds: ["table"], includeRaw: true } },
  ]);

  const table = first.tables[0];
  assert.equal(result.changed, 7);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(table.GetName(), "Metrics");
  assert.equal(table.GetRow(2).GetCell(2).GetText(), "43");
  assert.equal(table.rows.length, 3);
  assert.equal(table.columnWidths.length, 3);
  assert.equal(table.mergedCells.length, 2);
  assert.deepEqual(table.GetRow(1).GetCell(1).split, [2, 2]);
  assert.equal(table.GetColumnWidth(0), 70 * 36000);
  assert.equal(table.GetRow(0).GetHeight(), 14 * 36000);
  assert.equal(table.GetRow(0).GetCell(0).borders.top[0], 0.5);
  assert.equal(result.results[7].slides[0].objects[0].kind, "table");
  assert.equal(result.results[7].slides[0].objects[0].rows, 3);
  assert.equal(result.results[7].slides[0].objects[0].columns, 3);
});

test("Slides bridge writes scalar and formatted table cells before applying header defaults", async () => {
  const { bridge, presentation, first } = slidesHarness();
  const result = await bridge.execute([{
    name: "slides_add_table",
    arguments: {
      slide: 1,
      rows: 2,
      columns: 2,
      data: [
        [
          {
            text: "指标",
            fontFamily: "微软雅黑",
            bold: false,
            color: "#112233",
            backgroundColor: "#E8EEF5",
            align: "right",
          },
          42,
        ],
        [true],
      ],
      header: {
        bold: true,
        backgroundColor: "#DDEEFF",
        align: "center",
      },
    },
  }]);

  const table = first.tables[0];
  const formattedCell = table.GetRow(0).GetCell(0);
  const formattedParagraph = formattedCell.GetContent().GetAllParagraphs()[0];
  const formattedRun = formattedParagraph.runs[0];

  assert.equal(result.changed, 1);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(formattedCell.GetText(), "指标");
  assert.equal(formattedRun.format.fontFamily, "微软雅黑");
  assert.equal(formattedRun.format.color, "#112233");
  assert.equal(formattedRun.format.bold, true);
  assert.deepEqual(formattedCell.fill.color, { r: 221, g: 238, b: 255 });
  assert.equal(formattedParagraph.align, "center");
  assert.equal(table.GetRow(0).GetCell(1).GetText(), "42");
  assert.equal(table.GetRow(1).GetCell(0).GetText(), "true");
  assert.equal(table.GetRow(1).GetCell(1).GetText(), "");
  assert.equal(table.GetRow(0).GetCell(1).GetContent().GetAllParagraphs()[0].runs[0].format.bold, true);
  assert.equal(
    table.rows.flatMap(row => row.cells).some(cell => cell.GetText().includes("[object Object]")),
    false,
  );
});

test("Slides bridge rejects invalid formatted table cells before mutation", () => {
  const { bridge, presentation, first } = slidesHarness();
  let error;
  try {
    bridge.execute([{
      name: "slides_add_table",
      arguments: {
        slide: 1,
        rows: 1,
        columns: 5,
        data: [[
          { text: "错误字段", textColor: "#112233" },
          { bold: true },
          { text: { nested: "value" } },
          { text: "错误嵌套颜色", fill: { type: "solid", color: { nested: "value" } } },
          { text: "错误边框字段", border: { mystery: true } },
        ]],
      },
    }]);
  } catch (caught) {
    error = caught;
  }

  assert.equal(error && error.code, "INVALID_TOOL_ARGUMENTS");
  assert.deepEqual(
    Array.from(error.details.validationErrors, entry => entry.path),
    [
      "arguments.data[0][0].textColor",
      "arguments.data[0][1].text",
      "arguments.data[0][2].text",
      "arguments.data[0][3].fill.color",
      "arguments.data[0][4].border.mystery",
    ],
  );
  assert.equal(error.details.completedToolCalls, 0);
  assert.equal(error.details.partialMutationPossible, false);
  assert.equal(presentation.historyPoints, 0);
  assert.equal(first.tables.length, 0);
});

test("Slides bridge aligns, distributes, groups, reorders, and creates custom geometry", async () => {
  const { bridge, presentation, first } = slidesHarness();
  first.shapes[0].SetName("A");
  first.shapes[1].SetName("B");
  first.shapes[0].SetPosition(10 * 36000, 20 * 36000);
  first.shapes[1].SetPosition(100 * 36000, 50 * 36000);
  const third = new SlideShape("C");
  third.SetName("C");
  third.SetPosition(200 * 36000, 80 * 36000);
  first.AddObject(third);

  const result = await bridge.execute([
    {
      name: "slides_align_objects",
      arguments: {
        slide: 1,
        targets: [{ name: "A" }, { name: "B" }, { name: "C" }],
        align: "top",
        distribute: "horizontal",
        relativeTo: "selection",
      },
    },
    {
      name: "slides_group_objects",
      arguments: { slide: 1, action: "group", targets: [{ name: "A" }, { name: "B" }], name: "AB" },
    },
    { name: "slides_group_objects", arguments: { slide: 1, action: "ungroup", name: "AB" } },
    { name: "slides_reorder_object", arguments: { slide: 1, name: "C", action: "back" } },
    {
      name: "slides_add_connector",
      arguments: {
        slide: 1,
        connectorType: "straightConnector1",
        startXmm: 20,
        startYmm: 30,
        endXmm: 120,
        endYmm: 70,
        name: "Flow",
        line: { widthPt: 2, color: "#2255AA" },
      },
    },
    {
      name: "slides_add_freeform",
      arguments: {
        slide: 1,
        name: "Triangle",
        xMm: 140,
        yMm: 30,
        widthMm: 50,
        heightMm: 40,
        fill: { type: "solid", color: "#FFCC66" },
        paths: [{
          commands: [
            { type: "moveTo", xMm: 25, yMm: 0 },
            { type: "lineTo", xMm: 50, yMm: 40 },
            { type: "lineTo", xMm: 0, yMm: 40 },
            { type: "close" },
          ],
        }],
      },
    },
  ]);

  assert.equal(result.changed, 8);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(first.shapes.find(shape => shape.GetName() === "A").GetPosY(), 20 * 36000);
  assert.equal(first.groups.length, 0);
  assert.equal(first.GetAllDrawings()[0].GetName(), "C");
  assert.equal(first.shapes.some(shape => shape.GetName() === "Flow"), true);
  const freeform = first.shapes.find(shape => shape.GetName() === "Triangle");
  assert.equal(freeform.GetGeometry().IsCustom(), true);
  assert.equal(freeform.GetGeometry().paths[0].commands.length, 4);
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

test("Slides bridge adds WordArt, math equations, and OLE objects", async () => {
  const { bridge, presentation } = slidesHarness();
  const preview = {
    url: "https://app.test/copilot-api/images/asset.png?token=signed",
    assetId: "asset.png",
    widthPx: 800,
    heightPx: 600,
  };
  const result = await bridge.execute([
    {
      name: "slides_add_word_art",
      arguments: {
        slide: 1,
        text: "AI Bridge",
        transform: "textArchUp",
        fontSize: 32,
        bold: true,
        color: "#112233",
        widthMm: 90,
        heightMm: 25,
        name: "WordArt",
      },
    },
    {
      name: "slides_add_math",
      arguments: {
        slide: 1,
        text: "\\frac{a}{b}",
        format: "latex",
        xMm: 20,
        yMm: 50,
        widthMm: 60,
        heightMm: 20,
        name: "Equation",
      },
    },
    {
      name: "slides_add_ole_object",
      arguments: {
        slide: 1,
        _image: preview,
        data: "embedded-data",
        appId: "asc.custom",
        widthMm: 80,
        heightMm: 50,
        name: "Embedded",
        includeRaw: true,
      },
    },
  ]);

  assert.equal(result.changed, 3);
  assert.equal(result.needsSave, true);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(result.results[0].object.name, "WordArt");
  assert.equal(result.results[1].object.name, "Equation");
  assert.equal(result.results[1].format, "latex");
  assert.equal(result.results[2].object.kind, "oleObject");
  assert.equal(result.results[2].object.applicationId, "asc.custom");
  assert.equal(result.results[2].object.dataLength, "embedded-data".length);
});

test("Slides bridge manages animation timelines and exposes timing", async () => {
  const { bridge, presentation, first } = slidesHarness();
  first.shapes[0].SetName("Animated title");
  const result = await bridge.execute([
    {
      name: "slides_manage_animation",
      arguments: {
        slide: 1,
        action: "add",
        objectIndex: 0,
        effectType: "entranceFade",
        trigger: "afterprevious",
        durationMs: 900,
        delayMs: 125,
        repeatCount: 2,
      },
    },
    { name: "slides_inspect_animations", arguments: { slide: 1 } },
    {
      name: "slides_manage_animation",
      arguments: {
        slide: 1,
        action: "update",
        effectIndex: 0,
        trigger: "onclick",
        durationMs: 1200,
      },
    },
    { name: "slides_inspect_animations", arguments: { slide: 1 } },
  ]);

  assert.equal(result.changed, 2);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(result.results[1].slides[0].effects[0].objectName, "Animated title");
  assert.equal(result.results[1].slides[0].effects[0].durationMs, 900);
  assert.equal(result.results[1].slides[0].effects[0].delayMs, 125);
  assert.equal(result.results[3].slides[0].effects[0].trigger, "onclick");
  assert.equal(result.results[3].slides[0].effects[0].durationMs, 1200);
});

test("Slides bridge inspects, edits, replies to, and deletes comments", async () => {
  const { bridge, presentation } = slidesHarness();
  const result = await bridge.execute([
    {
      name: "slides_add_comment",
      arguments: { slide: 1, text: "Review this", author: "Alice", userId: "alice" },
    },
    {
      name: "slides_manage_comment",
      arguments: { action: "addReply", commentIndex: 0, text: "Done", author: "Bob", userId: "bob" },
    },
    {
      name: "slides_manage_comment",
      arguments: { action: "update", commentIndex: 0, text: "Reviewed", solved: true, xMm: 10, yMm: 20 },
    },
    { name: "slides_inspect_comments", arguments: {} },
    { name: "slides_manage_comment", arguments: { action: "delete", commentIndex: 0 } },
  ]);

  assert.equal(result.changed, 4);
  assert.equal(presentation.historyPoints, 1);
  assert.equal(result.results[3].comments[0].text, "Reviewed");
  assert.equal(result.results[3].comments[0].solved, true);
  assert.equal(result.results[3].comments[0].replies[0].text, "Done");
  assert.deepEqual(Array.from(result.results[3].comments[0].position), [360000, 720000]);
  assert.equal(presentation.comments.length, 0);
});

test("Slides bridge uses plugin methods for built-in themes, macros, and slideshow control", async () => {
  const { bridge, presentation } = slidesHarness();
  const themes = await bridge.execute([
    { name: "slides_inspect_builtin_themes", arguments: {} },
  ]);
  const applied = await bridge.execute([
    { name: "slides_apply_builtin_theme", arguments: { theme: 7 } },
  ]);
  const macros = await bridge.execute([
    { name: "slides_set_macros", arguments: { content: { macrosArray: [] } } },
  ]);
  const slideshow = await bridge.execute([
    { name: "slides_control_slideshow", arguments: { action: "goto", slide: 2 } },
  ]);

  assert.equal(themes.results[0].content[0].name, "Ion");
  assert.equal(applied.needsSave, true);
  assert.equal(macros.needsSave, true);
  assert.equal(macros.results[0].kind, "office");
  assert.equal(slideshow.needsSave, false);
  assert.deepEqual(Array.from(presentation.pluginCalls, entry => entry[0]), [
    "GetEditorThemes",
    "ApplyTheme",
    "SetMacros",
    "GoToSlideInSlideShow",
  ]);
  assert.deepEqual(Array.from(presentation.pluginCalls[3][1]), [1]);
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
  ]);

  const deleted = await bridge.execute([
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
  ]);
  assert.equal(deleted.results[0].name, "sheets_delete_sheet");
  assert.equal(result.editorType, "cell");
  assert.equal(result.needsSave, true);
  assert.equal(workbook.historyPoints, 2);
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
  const created = await bridge.execute([
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
  ]);
  const inspected = await bridge.execute([
    { name: "sheets_inspect_charts", arguments: { sheet: "Sheet1", includeRaw: true } },
  ]);
  const updated = await bridge.execute([
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
  ]);
  const deleted = await bridge.execute([
    { name: "sheets_delete_chart", arguments: { sheet: "Sheet1", chartIndex: 0 } },
  ]);

  assert.equal(created.changed, 1);
  assert.equal(updated.changed, 1);
  assert.equal(deleted.changed, 1);
  assert.equal(created.needsSave, true);
  assert.equal(workbook.historyPoints, 3);
  assert.equal(created.results[0].chart.name, "TrendChart");
  assert.equal(inspected.results[0].sheets[0].charts[0].chartIndex, 0);
  assert.equal(inspected.results[0].sheets[0].charts[0].raw.name, "TrendChart");
  assert.equal(updated.results[0].chart.title, "Updated Trend");
  assert.equal(updated.results[0].chart.raw.series[0].name, "Actual");
  assert.equal(updated.results[0].chart.raw.series[1].name, "Forecast");
  assert.equal(sheet1.charts.length, 0);
});

test("Sheets bridge rejects same-batch references to a newly created sheet before history", async () => {
  const { bridge, workbook } = sheetsHarness();
  await assert.rejects(
    bridge.execute([
      { name: "sheets_add_sheet", arguments: { name: "NewData" } },
      { name: "sheets_set_values", arguments: { sheet: "NewData", range: "A1", values: 1 } },
    ]),
    error => {
      assert.equal(error.code, "BATCH_DEPENDENCY_REQUIRES_SPLIT");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      return true;
    },
  );
  assert.equal(workbook.historyPoints, 0);
  assert.deepEqual(workbook.sheets.map(sheet => sheet.GetName()), ["Sheet1", "Sheet2"]);
});

test("Sheets bridge clears null cells and writes only contiguous non-null segments", async () => {
  const { bridge, workbook, sheet1 } = sheetsHarness();
  const target = sheet1.range;
  await bridge.execute([
    {
      name: "sheets_set_values",
      arguments: {
        sheet: "Sheet1",
        range: "A1:C2",
        values: [["A", null, "B"], [null, "C", "D"]],
      },
    },
  ]);
  assert.equal(workbook.historyPoints, 1);
  assert.equal(target.clearContentsCalls, 1);
  assert.equal(JSON.stringify(target.setValueCalls), JSON.stringify([[["A"]], [["B"]], [["C", "D"]]]));
  assert.equal(target.setValueCalls.flat(Infinity).includes(null), false);
});

test("Sheets bridge rejects unsupported fill before creating history", async () => {
  const { bridge, workbook } = sheetsHarness();
  await assert.rejects(
    bridge.execute([
      { name: "sheets_manage_range", arguments: { sheet: "Sheet1", range: "A1:A3", action: "fillDown" } },
    ]),
    error => {
      assert.equal(error.code, "SHEETS_API_UNSUPPORTED");
      assert.equal(error.details.feature, "sheets.rangeFill.fillDown");
      return true;
    },
  );
  assert.equal(workbook.historyPoints, 0);
});

test("Sheets bridge rejects unsupported chart types before AddChart", async () => {
  const { bridge, sheet1 } = sheetsHarness();
  let caught;
  try {
    await bridge.execute([{
      name: "sheets_add_chart",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        type: "columnClustered",
      },
    }]);
  } catch (error) {
    caught = error;
  }
  assert.equal(caught && caught.code, "INVALID_TOOL_ARGUMENTS");
  assert.equal(caught.details.phase, "sheets-chart-validation");
  assert.equal(caught.details.receivedType, "columnClustered");
  assert.deepEqual(Array.from(caught.details.suggestedTypes), ["column", "bar"]);
  assert.equal(caught.details.completedToolCalls, 0);
  assert.equal(caught.details.partialMutationPossible, false);
  assert.equal(sheet1.charts.length, 0);
});

test("Sheets bridge rejects advanced series before creating a chart", async () => {
  const { bridge, sheet1 } = sheetsHarness();
  let caught;
  try {
    await bridge.execute([{
      name: "sheets_add_chart",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        type: "line",
        addSeries: [{ name: "Forecast", valuesRange: "'Sheet1'!$C$2:$C$3" }],
      },
    }]);
  } catch (error) {
    caught = error;
  }
  assert.equal(caught && caught.code, "SHEETS_API_UNSUPPORTED");
  assert.equal(caught.details.feature, "sheets.charts.addSeriesOnCreate");
  assert.equal(caught.details.completedToolCalls, 0);
  assert.equal(caught.details.partialMutationPossible, false);
  assert.equal(sheet1.charts.length, 0);
});

test("Sheets bridge reports an unremovable chart created before option failure", async () => {
  const { bridge, sheet1 } = sheetsHarness();
  const originalAddChart = sheet1.AddChart.bind(sheet1);
  sheet1.AddChart = (...args) => {
    const chart = originalAddChart(...args);
    chart.Delete = undefined;
    chart.SetTitle = () => { throw new Error("chart option failed"); };
    return chart;
  };

  let caught;
  try {
    await bridge.execute([{
      name: "sheets_add_chart",
      arguments: { sheet: "Sheet1", range: "A1:B3", type: "line", title: "Trend" },
    }]);
  } catch (error) {
    caught = error;
  }

  assert.equal(caught && caught.code, "SHEETS_CHART_PARTIAL_MUTATION");
  assert.equal(caught.details.created, true);
  assert.equal(caught.details.chartIndex, 0);
  assert.equal(caught.details.mutationState, "partial");
  assert.equal(caught.details.completedToolCalls, 0);
  assert.equal(caught.details.partialMutationPossible, true);
  assert.equal(sheet1.charts.length, 1);
});

test("Sheets bridge reports structured details when AddChart returns null", async () => {
  const { bridge, sheet1 } = sheetsHarness();
  sheet1.AddChart = () => null;
  let caught;
  try {
    await bridge.execute([{
      name: "sheets_add_chart",
      arguments: {
        sheet: "Sheet1",
        range: "A1:B3",
        type: "column",
      },
    }]);
  } catch (error) {
    caught = error;
  }
  assert.equal(caught && caught.code, "SHEETS_CHART_CREATE_FAILED");
  assert.equal(caught.details.apiMethod, "ApiWorksheet.AddChart");
  assert.equal(caught.details.receivedType, "column");
  assert.equal(caught.details.normalizedType, "bar");
  assert.equal(caught.details.completedToolCalls, 0);
  assert.equal(caught.details.partialMutationPossible, true);
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
