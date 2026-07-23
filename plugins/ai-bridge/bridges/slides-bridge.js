(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) throw new Error((value && value.error) || "Slides Bridge 执行失败");
    return value;
  }

  function execute(toolCalls) {
    return new Promise(function (resolve, reject) {
      Asc.scope.copilotSlideCalls = toolCalls;
      Asc.plugin.info.recalculate = true;

      Asc.plugin.callCommand(
        function () {
          try {
            var calls = Asc.scope.copilotSlideCalls || [];
            var presentation = Api.GetPresentation();
            var results = [];
            var changed = 0;
            var mutating = false;
            var EMU_PER_MM = 36000;
            var EMU_PER_POINT = 12700;
            var mutatingNames = {
              slides_replace_text: true,
              slides_scale_font: true,
              slides_format_text: true,
              slides_format_selection: true,
              slides_add_slide: true,
              slides_duplicate_slide: true,
              slides_delete_slide: true,
              slides_add_textbox: true,
              slides_set_background: true,
              slides_add_shape: true,
              slides_update_shape: true,
              slides_delete_object: true,
              slides_add_chart: true,
              slides_update_chart: true,
              slides_delete_chart: true,
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

            function emuToPoints(value) {
              return typeof value === "number" && isFinite(value)
                ? Math.round(value / EMU_PER_POINT * 1000) / 1000
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

            function slideCount() {
              if (typeof presentation.GetSlidesCount === "function") return presentation.GetSlidesCount();
              if (typeof presentation.GetSlideCount === "function") return presentation.GetSlideCount();
              return presentation.GetAllSlides().length;
            }

            function getSlide(oneBasedIndex) {
              var index = Number(oneBasedIndex);
              if (!Number.isInteger(index) || index < 1 || index > slideCount()) {
                throw new Error("幻灯片页码超出范围，当前共 " + slideCount() + " 页");
              }
              return presentation.GetSlideByIndex(index - 1);
            }

            function slideShapes(slide) {
              return typeof slide.GetAllShapes === "function" ? slide.GetAllShapes() : [];
            }

            function slideCharts(slide) {
              return typeof slide.GetAllCharts === "function" ? slide.GetAllCharts() : [];
            }

            function pushUnique(target, values) {
              for (var index = 0; index < values.length; index += 1) {
                if (target.indexOf(values[index]) === -1) target.push(values[index]);
              }
            }

            function slideDrawings(slide) {
              if (typeof slide.GetAllDrawings === "function") return slide.GetAllDrawings();
              var drawings = [];
              pushUnique(drawings, slideShapes(slide));
              pushUnique(drawings, slideCharts(slide));
              if (typeof slide.GetAllImages === "function") pushUnique(drawings, slide.GetAllImages());
              if (typeof slide.GetAllTables === "function") pushUnique(drawings, slide.GetAllTables());
              if (typeof slide.GetAllOleObjects === "function") pushUnique(drawings, slide.GetAllOleObjects());
              return drawings;
            }

            function drawingKind(drawing) {
              var kind = safeCall(drawing, "GetClassType");
              return kind ? String(kind) : "drawing";
            }

            function drawingText(drawing) {
              var content;
              try {
                content = drawing && typeof drawing.GetContent === "function" ? drawing.GetContent() : null;
              } catch (error) {
                content = null;
              }
              return content && typeof content.GetText === "function" ? String(content.GetText() || "") : "";
            }

            function getRunsFromShape(shape) {
              var runs = [];
              var content;
              try {
                content = shape.GetContent();
              } catch (error) {
                return runs;
              }
              if (!content || typeof content.GetAllParagraphs !== "function") return runs;
              var paragraphs = content.GetAllParagraphs();
              for (var p = 0; p < paragraphs.length; p += 1) {
                var paragraph = paragraphs[p];
                var count = paragraph.GetElementsCount();
                for (var e = 0; e < count; e += 1) {
                  var element = paragraph.GetElement(e);
                  if (!element || typeof element.GetClassType !== "function") continue;
                  if (element.GetClassType() === "run") runs.push(element);
                }
              }
              return runs;
            }

            function applyRunFormat(run, format) {
              if (format.fontSize !== undefined) run.SetFontSize(Math.max(2, Math.round(Number(format.fontSize) * 2)));
              if (format.fontFamily) run.SetFontFamily(String(format.fontFamily));
              if (format.bold !== undefined) run.SetBold(Boolean(format.bold));
              if (format.italic !== undefined) run.SetItalic(Boolean(format.italic));
              if (format.underline !== undefined) run.SetUnderline(Boolean(format.underline));
              if (format.color) run.SetColor(Api.Color(String(format.color)));
            }

            function replaceShapeText(shape, args) {
              var content = shape.GetContent();
              if (!content) throw new Error("目标图形不支持文本内容");
              if (typeof content.RemoveAllElements === "function") content.RemoveAllElements();
              var paragraph = Api.CreateParagraph();
              paragraph.SetJc(String(args.align || "left"));
              var run = paragraph.AddText(String(args.text || ""));
              applyRunFormat(run, args);
              content.Push(paragraph);
            }

            function formatShapeText(shape, args) {
              var hasRunFormat = hasOwn(args, "fontSize") ||
                hasOwn(args, "fontFamily") ||
                hasOwn(args, "bold") ||
                hasOwn(args, "italic") ||
                hasOwn(args, "underline") ||
                hasOwn(args, "color");
              if (hasRunFormat) {
                var runs = getRunsFromShape(shape);
                for (var runIndex = 0; runIndex < runs.length; runIndex += 1) {
                  applyRunFormat(runs[runIndex], args);
                }
              }
              if (hasOwn(args, "align")) {
                var content = safeCall(shape, "GetContent");
                var paragraphs = content && typeof content.GetAllParagraphs === "function"
                  ? content.GetAllParagraphs()
                  : [];
                for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
                  if (typeof paragraphs[paragraphIndex].SetJc === "function") {
                    paragraphs[paragraphIndex].SetJc(String(args.align));
                  }
                }
              }
            }

            function apiColor(value) {
              var hex = String(value || "").replace("#", "");
              if (!/^[0-9a-fA-F]{6}$/.test(hex)) throw new Error("颜色必须是 #RRGGBB 格式");
              var red = parseInt(hex.slice(0, 2), 16);
              var green = parseInt(hex.slice(2, 4), 16);
              var blue = parseInt(hex.slice(4, 6), 16);
              if (typeof Api.CreateRGBColor === "function") return Api.CreateRGBColor(red, green, blue);
              return Api.RGB(red, green, blue);
            }

            function createGradientStops(stops) {
              if (!Array.isArray(stops) || stops.length < 2 || stops.length > 16) {
                throw new Error("渐变色 stops 必须包含 2 到 16 个颜色节点");
              }
              if (typeof Api.CreateGradientStop !== "function") throw new Error("当前 ONLYOFFICE 版本不支持渐变节点");
              return stops.map(function (stop) {
                if (!isObject(stop) || !stop.color) throw new Error("每个渐变节点都需要 color");
                var position = clamp(asFinite(stop.position, 0), 0, 100);
                return Api.CreateGradientStop(apiColor(stop.color), Math.round(position * 1000));
              });
            }

            function createFill(spec, fallbackColor) {
              if (spec === undefined || spec === null) {
                return fallbackColor ? Api.CreateSolidFill(apiColor(fallbackColor)) : Api.CreateNoFill();
              }
              if (typeof spec === "string") return Api.CreateSolidFill(apiColor(spec));
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
              if (type === "solid") return Api.CreateSolidFill(apiColor(spec.color));
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
                  apiColor(spec.backgroundColor),
                  apiColor(spec.foregroundColor)
                );
              }
              throw new Error("不支持的填充类型：" + type);
            }

            function describeFill(fill) {
              if (!fill) return null;
              return {
                type: safeCall(fill, "GetType"),
                raw: serialized(fill),
              };
            }

            function createStroke(spec, legacyColor, legacyWidthPt) {
              if (spec === undefined || spec === null) {
                if (!legacyColor && legacyWidthPt === undefined) return Api.CreateStroke(0, Api.CreateNoFill());
                return Api.CreateStroke(
                  Math.max(0, asFinite(legacyWidthPt, 1)) * EMU_PER_POINT,
                  legacyColor ? Api.CreateSolidFill(apiColor(legacyColor)) : Api.CreateNoFill()
                );
              }
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
              var lineFill = spec.fill !== undefined
                ? createFill(spec.fill)
                : (spec.color ? Api.CreateSolidFill(apiColor(spec.color)) : Api.CreateNoFill());
              return Api.CreateStroke(Math.max(0, asFinite(spec.widthPt, 1)) * EMU_PER_POINT, lineFill);
            }

            function describeStroke(stroke) {
              if (!stroke) return null;
              var fill = safeCall(stroke, "GetFill");
              return {
                widthPt: emuToPoints(safeCall(stroke, "GetWidth")),
                fill: describeFill(fill),
                raw: serialized(stroke),
              };
            }

            function getDrawingFrame(drawing) {
              return {
                x: safeCall(drawing, "GetPosX"),
                y: safeCall(drawing, "GetPosY"),
                width: safeCall(drawing, "GetWidth"),
                height: safeCall(drawing, "GetHeight"),
              };
            }

            function applyDrawingFrame(drawing, args) {
              var frame = getDrawingFrame(drawing);
              if (hasOwn(args, "xMm") || hasOwn(args, "yMm")) {
                var x = hasOwn(args, "xMm") ? mmToEmu(args.xMm) : (typeof frame.x === "number" ? frame.x : 0);
                var y = hasOwn(args, "yMm") ? mmToEmu(args.yMm) : (typeof frame.y === "number" ? frame.y : 0);
                if (typeof drawing.SetPosition === "function") drawing.SetPosition(x, y);
                else {
                  if (typeof drawing.SetPosX === "function") drawing.SetPosX(x);
                  if (typeof drawing.SetPosY === "function") drawing.SetPosY(y);
                }
              }
              if (hasOwn(args, "widthMm") || hasOwn(args, "heightMm")) {
                var width = hasOwn(args, "widthMm") ? mmToEmu(args.widthMm) : (typeof frame.width === "number" ? frame.width : mmToEmu(100));
                var height = hasOwn(args, "heightMm") ? mmToEmu(args.heightMm) : (typeof frame.height === "number" ? frame.height : mmToEmu(60));
                if (typeof drawing.SetSize === "function") drawing.SetSize(Math.max(1, width), Math.max(1, height));
              }
              if (hasOwn(args, "rotationDeg") && typeof drawing.SetRotation === "function") {
                drawing.SetRotation(asFinite(args.rotationDeg, 0));
              }
              if (hasOwn(args, "flipH") && typeof drawing.SetFlipH === "function") drawing.SetFlipH(Boolean(args.flipH));
              if (hasOwn(args, "flipV") && typeof drawing.SetFlipV === "function") drawing.SetFlipV(Boolean(args.flipV));
              if (hasOwn(args, "name") && typeof drawing.SetName === "function") drawing.SetName(String(args.name));
            }

            function applyShapeOptions(shape, args) {
              applyDrawingFrame(shape, args);
              if (hasOwn(args, "shapeType")) {
                if (typeof Api.CreatePresetGeometry !== "function" || typeof shape.SetGeometry !== "function") {
                  throw new Error("当前 ONLYOFFICE 版本不支持修改图形几何");
                }
                var geometry = Api.CreatePresetGeometry(String(args.shapeType));
                if (!geometry) throw new Error("无效 shapeType：" + args.shapeType);
                shape.SetGeometry(geometry);
              }
              if (hasOwn(args, "fill") || hasOwn(args, "fillColor")) {
                var fill = createFill(args.fill, args.fillColor);
                if (typeof shape.SetFill === "function") shape.SetFill(fill);
                else if (typeof shape.Fill === "function") shape.Fill(fill);
              }
              if (hasOwn(args, "line") || hasOwn(args, "lineColor") || hasOwn(args, "lineWidthPt")) {
                var stroke = createStroke(args.line, args.lineColor, args.lineWidthPt);
                if (typeof shape.SetLine === "function") shape.SetLine(stroke);
                else if (typeof shape.SetOutLine === "function") shape.SetOutLine(stroke);
              }
              if (hasOwn(args, "verticalAlign") && typeof shape.SetVerticalTextAlign === "function") {
                shape.SetVerticalTextAlign(String(args.verticalAlign));
              }
              if (isObject(args.paddingMm) && typeof shape.SetPaddings === "function") {
                shape.SetPaddings(
                  mmToEmu(args.paddingMm.left || 0),
                  mmToEmu(args.paddingMm.top || 0),
                  mmToEmu(args.paddingMm.right || 0),
                  mmToEmu(args.paddingMm.bottom || 0)
                );
              }
              if (hasOwn(args, "text")) replaceShapeText(shape, args);
              else formatShapeText(shape, args);
            }

            function describeDrawing(drawing, objectIndex, includeRaw) {
              var kind = drawingKind(drawing);
              var geometry = kind === "shape" ? safeCall(drawing, "GetGeometry") : null;
              var frame = getDrawingFrame(drawing);
              var info = {
                objectIndex: objectIndex,
                objectId: safeCall(drawing, "GetInternalId"),
                name: safeCall(drawing, "GetName"),
                kind: kind,
                xMm: emuToMm(frame.x),
                yMm: emuToMm(frame.y),
                widthMm: emuToMm(frame.width),
                heightMm: emuToMm(frame.height),
                rotationDeg: safeCall(drawing, "GetRotation"),
                flipH: safeCall(drawing, "GetFlipH"),
                flipV: safeCall(drawing, "GetFlipV"),
              };
              if (kind === "shape") {
                info.shapeType = geometry ? safeCall(geometry, "GetPreset") : null;
                info.text = drawingText(drawing);
                info.fill = describeFill(safeCall(drawing, "GetFill"));
                info.line = describeStroke(safeCall(drawing, "GetLine"));
              }
              if (kind === "chart") {
                info.chartIndex = objectIndex;
                info.chartType = safeCall(drawing, "GetChartType");
                info.title = safeCall(drawing, "GetTitle");
                var series = safeCall(drawing, "GetAllSeries");
                info.series = Array.isArray(series) ? series.map(function (item, index) {
                  return {
                    index: index,
                    chartType: safeCall(item, "GetChartType"),
                    raw: includeRaw ? serialized(item) : undefined,
                  };
                }) : [];
              }
              if (includeRaw) info.raw = serialized(drawing);
              return info;
            }

            function drawingMatches(drawing, args, objectIndex) {
              var requestedId = args.objectId || args.chartId;
              if (requestedId && String(safeCall(drawing, "GetInternalId")) !== String(requestedId)) return false;
              if (args.name && String(safeCall(drawing, "GetName")) !== String(args.name)) return false;
              var hasObjectIndex = hasOwn(args, "objectIndex");
              var hasChartIndex = hasOwn(args, "chartIndex");
              var requestedIndex = hasObjectIndex ? args.objectIndex : args.chartIndex;
              if ((hasObjectIndex || hasChartIndex) && objectIndex !== Number(requestedIndex)) return false;
              return Boolean(requestedId || args.name || hasObjectIndex || hasChartIndex);
            }

            function resolveDrawing(args, expectedKind) {
              var slide = getSlide(args.slide);
              var drawings = expectedKind === "chart" ? slideCharts(slide) : slideDrawings(slide);
              for (var index = 0; index < drawings.length; index += 1) {
                if (!drawingMatches(drawings[index], args, index)) continue;
                var kind = drawingKind(drawings[index]);
                if (expectedKind && kind !== expectedKind) {
                  throw new Error("目标对象类型为 " + kind + "，需要 " + expectedKind);
                }
                return { slide: slide, drawing: drawings[index], index: index };
              }
              throw new Error("找不到目标对象；请提供有效的 objectId、name 或对象序号");
            }

            function applyChartAxis(chart, axis, horizontal) {
              if (!isObject(axis)) return;
              var prefix = horizontal ? "Hor" : "Vert";
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

            function applyChartSeries(chart, updates, spreadsheetMode) {
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
                if (spreadsheetMode && hasOwn(update, "valuesRange") && typeof chart.SetSeriaValues === "function") {
                  chart.SetSeriaValues(String(update.valuesRange), seriesIndex);
                } else if (!spreadsheetMode && Array.isArray(update.values) && typeof chart.SetSeriaValues === "function") {
                  chart.SetSeriaValues(update.values.map(Number), seriesIndex);
                }
                if (spreadsheetMode && hasOwn(update, "xValuesRange") && typeof chart.SetSeriaXValues === "function") {
                  chart.SetSeriaXValues(String(update.xValuesRange), seriesIndex);
                }
                if (!spreadsheetMode && Array.isArray(update.xValues) && typeof chart.SetXValues === "function") {
                  chart.SetXValues(update.xValues.map(Number));
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
              if (!hasOwn(value || {}, "line")) return;
              if (typeof chart[method] === "function") chart[method](createStroke(value.line));
            }

            function applyChartOptions(chart, args, spreadsheetMode) {
              applyDrawingFrame(chart, args);
              if (hasOwn(args, "style") && typeof chart.ApplyChartStyle === "function") {
                chart.ApplyChartStyle(clamp(Math.round(asFinite(args.style, 2)), 1, 48));
              }
              if (hasOwn(args, "title") && typeof chart.SetTitle === "function") {
                chart.SetTitle(String(args.title), asFinite(args.titleFontSize, 13));
              }
              if (hasOwn(args, "fill") && typeof chart.Fill === "function") chart.Fill(createFill(args.fill));
              if (hasOwn(args, "line") && typeof chart.SetOutLine === "function") chart.SetOutLine(createStroke(args.line));
              if (hasOwn(args, "plotAreaFill") && typeof chart.SetPlotAreaFill === "function") {
                chart.SetPlotAreaFill(createFill(args.plotAreaFill));
              }
              if (hasOwn(args, "plotAreaLine") && typeof chart.SetPlotAreaOutLine === "function") {
                chart.SetPlotAreaOutLine(createStroke(args.plotAreaLine));
              }
              if (hasOwn(args, "titleFill") && typeof chart.SetTitleFill === "function") {
                chart.SetTitleFill(createFill(args.titleFill));
              }
              if (hasOwn(args, "titleLine") && typeof chart.SetTitleOutLine === "function") {
                chart.SetTitleOutLine(createStroke(args.titleLine));
              }
              if (isObject(args.legend)) {
                if (hasOwn(args.legend, "position") && typeof chart.SetLegendPos === "function") {
                  chart.SetLegendPos(String(args.legend.position));
                }
                if (hasOwn(args.legend, "fontSize") && typeof chart.SetLegendFontSize === "function") {
                  chart.SetLegendFontSize(asFinite(args.legend.fontSize, 10));
                }
                if (hasOwn(args.legend, "fill") && typeof chart.SetLegendFill === "function") {
                  chart.SetLegendFill(createFill(args.legend.fill));
                }
                if (hasOwn(args.legend, "line") && typeof chart.SetLegendOutLine === "function") {
                  chart.SetLegendOutLine(createStroke(args.legend.line));
                }
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
              if (!spreadsheetMode && Array.isArray(args.categories) && typeof chart.SetCategoryName === "function") {
                for (var categoryIndex = 0; categoryIndex < args.categories.length; categoryIndex += 1) {
                  chart.SetCategoryName(String(args.categories[categoryIndex]), categoryIndex);
                }
              }
              applyChartSeries(chart, args.seriesUpdates, spreadsheetMode);
              if (Array.isArray(args.removeSeries) && typeof chart.RemoveSeria === "function") {
                var ordered = args.removeSeries.slice().map(Number).sort(function (left, right) { return right - left; });
                for (var removeIndex = 0; removeIndex < ordered.length; removeIndex += 1) {
                  chart.RemoveSeria(ordered[removeIndex]);
                }
              }
            }

            function createTextBox(slide, args) {
              var width = Math.max(10, hasOwn(args, "widthMm") ? Number(args.widthMm) : 120) * EMU_PER_MM;
              var height = Math.max(5, hasOwn(args, "heightMm") ? Number(args.heightMm) : 30) * EMU_PER_MM;
              var fill = createFill(args.fill, args.fillColor);
              var stroke = createStroke(args.line, args.lineColor, args.lineWidthPt);
              var shape = Api.CreateShape("rect", width, height, fill, stroke);
              var positionArgs = {
                xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                yMm: hasOwn(args, "yMm") ? args.yMm : 15,
              };
              applyDrawingFrame(shape, positionArgs);
              applyShapeOptions(shape, args);
              slide.AddObject(shape);
              return shape;
            }

            for (var callIndex = 0; callIndex < calls.length; callIndex += 1) {
              if (mutatingNames[calls[callIndex].name]) {
                mutating = true;
                break;
              }
            }
            if (mutating && typeof presentation.CreateNewHistoryPoint === "function") presentation.CreateNewHistoryPoint();

            for (var callIndex = 0; callIndex < calls.length; callIndex += 1) {
              var call = calls[callIndex];
              var args = getArgs(call);

              switch (call.name) {
                case "slides_inspect": {
                  var maxChars = Math.min(30000, Math.max(500, Number(args.maxChars) || 12000));
                  var slides = [];
                  var totalChars = 0;
                  for (var slideIndex = 0; slideIndex < slideCount(); slideIndex += 1) {
                    var inspectedSlide = presentation.GetSlideByIndex(slideIndex);
                    var shapes = slideShapes(inspectedSlide);
                    var pieces = [];
                    for (var s = 0; s < shapes.length; s += 1) {
                      var shapeText = drawingText(shapes[s]);
                      if (shapeText) pieces.push(shapeText);
                    }
                    var text = pieces.join("\n");
                    var remaining = Math.max(0, maxChars - totalChars);
                    slides.push({
                      slide: slideIndex + 1,
                      text: text.slice(0, remaining),
                      shapes: shapes.length,
                      charts: slideCharts(inspectedSlide).length,
                      objects: slideDrawings(inspectedSlide).length,
                    });
                    totalChars += Math.min(text.length, remaining);
                    if (totalChars >= maxChars) break;
                  }
                  results.push({
                    name: call.name,
                    slideCount: slideCount(),
                    currentSlide: presentation.GetCurSlideIndex() + 1,
                    slides: slides,
                    truncated: totalChars >= maxChars,
                  });
                  break;
                }

                case "slides_inspect_objects": {
                  var firstSlide = args.slide ? Number(args.slide) - 1 : 0;
                  var lastSlide = args.slide ? firstSlide + 1 : slideCount();
                  if (firstSlide < 0 || lastSlide > slideCount()) throw new Error("幻灯片页码超出范围");
                  var maxObjects = clamp(Math.round(asFinite(args.maxObjects, 200)), 1, 1000);
                  var kinds = Array.isArray(args.kinds) ? args.kinds.map(String) : null;
                  var objectSlides = [];
                  var objectTotal = 0;
                  for (var objectSlideIndex = firstSlide; objectSlideIndex < lastSlide && objectTotal < maxObjects; objectSlideIndex += 1) {
                    var objectSlide = presentation.GetSlideByIndex(objectSlideIndex);
                    var drawings = slideDrawings(objectSlide);
                    var objects = [];
                    for (var drawingIndex = 0; drawingIndex < drawings.length && objectTotal < maxObjects; drawingIndex += 1) {
                      var kind = drawingKind(drawings[drawingIndex]);
                      if (kinds && kinds.indexOf(kind) === -1) continue;
                      objects.push(describeDrawing(drawings[drawingIndex], drawingIndex, Boolean(args.includeRaw)));
                      objectTotal += 1;
                    }
                    objectSlides.push({
                      slide: objectSlideIndex + 1,
                      widthMm: emuToMm(safeCall(objectSlide, "GetWidth")),
                      heightMm: emuToMm(safeCall(objectSlide, "GetHeight")),
                      objects: objects,
                      raw: args.includeSlideRaw ? serialized(objectSlide) : undefined,
                    });
                  }
                  results.push({
                    name: call.name,
                    slides: objectSlides,
                    objectCount: objectTotal,
                    truncated: objectTotal >= maxObjects,
                  });
                  break;
                }

                case "slides_replace_text": {
                  if (!args.search) throw new Error("slides_replace_text.search 不能为空");
                  var replaceStart = args.slide ? Number(args.slide) - 1 : 0;
                  var replaceEnd = args.slide ? replaceStart + 1 : slideCount();
                  if (replaceStart < 0 || replaceEnd > slideCount()) throw new Error("幻灯片页码超出范围");
                  var replacedRuns = 0;
                  for (var replaceSlideIndex = replaceStart; replaceSlideIndex < replaceEnd; replaceSlideIndex += 1) {
                    var replaceShapes = slideShapes(presentation.GetSlideByIndex(replaceSlideIndex));
                    for (var replaceShapeIndex = 0; replaceShapeIndex < replaceShapes.length; replaceShapeIndex += 1) {
                      var replaceRuns = getRunsFromShape(replaceShapes[replaceShapeIndex]);
                      for (var replaceRunIndex = 0; replaceRunIndex < replaceRuns.length; replaceRunIndex += 1) {
                        var oldText = String(replaceRuns[replaceRunIndex].GetText() || "");
                        var newText = args.matchCase
                          ? oldText.split(String(args.search)).join(String(args.replace || ""))
                          : oldText.replace(new RegExp(String(args.search).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), String(args.replace || ""));
                        if (newText === oldText) continue;
                        replaceRuns[replaceRunIndex].RemoveAllElements();
                        replaceRuns[replaceRunIndex].AddText(newText);
                        replacedRuns += 1;
                      }
                    }
                  }
                  changed += replacedRuns;
                  results.push({ name: call.name, replacedRuns: replacedRuns, search: String(args.search) });
                  break;
                }

                case "slides_scale_font":
                case "slides_format_text": {
                  var formatStart = args.slide ? Number(args.slide) - 1 : 0;
                  var formatEnd = args.slide ? formatStart + 1 : slideCount();
                  if (formatStart < 0 || formatEnd > slideCount()) throw new Error("幻灯片页码超出范围");
                  var scale = args.scale === undefined ? null : Number(args.scale);
                  if (call.name === "slides_scale_font" && (!isFinite(scale) || scale <= 0 || scale > 3)) {
                    throw new Error("slides_scale_font.scale 必须大于 0 且不超过 3");
                  }
                  var changedRuns = 0;
                  for (var formatSlideIndex = formatStart; formatSlideIndex < formatEnd; formatSlideIndex += 1) {
                    var formatShapes = slideShapes(presentation.GetSlideByIndex(formatSlideIndex));
                    for (var formatShapeIndex = 0; formatShapeIndex < formatShapes.length; formatShapeIndex += 1) {
                      var formatRuns = getRunsFromShape(formatShapes[formatShapeIndex]);
                      for (var formatRunIndex = 0; formatRunIndex < formatRuns.length; formatRunIndex += 1) {
                        if (scale !== null) {
                          var oldSize = formatRuns[formatRunIndex].GetFontSize();
                          if (typeof oldSize !== "number" || oldSize <= 0) continue;
                          formatRuns[formatRunIndex].SetFontSize(Math.max(2, Math.round(oldSize * scale)));
                        } else {
                          applyRunFormat(formatRuns[formatRunIndex], args);
                        }
                        changedRuns += 1;
                      }
                    }
                  }
                  changed += changedRuns;
                  results.push({ name: call.name, changedRuns: changedRuns, slide: args.slide || "all" });
                  break;
                }

                case "slides_format_selection": {
                  var selection = Api.GetSelection();
                  var selectedShapes = selection && typeof selection.GetShapes === "function" ? selection.GetShapes() : [];
                  if (!selectedShapes.length) throw new Error("请先在 PPT 中选择一个或多个文本框/形状");
                  var selectedChangedRuns = 0;
                  for (var selectedIndex = 0; selectedIndex < selectedShapes.length; selectedIndex += 1) {
                    var selectedRuns = getRunsFromShape(selectedShapes[selectedIndex]);
                    for (var selectedRunIndex = 0; selectedRunIndex < selectedRuns.length; selectedRunIndex += 1) {
                      applyRunFormat(selectedRuns[selectedRunIndex], args);
                      selectedChangedRuns += 1;
                    }
                  }
                  changed += selectedChangedRuns;
                  results.push({ name: call.name, selectedShapes: selectedShapes.length, changedRuns: selectedChangedRuns });
                  break;
                }

                case "slides_add_slide": {
                  var newSlide = Api.CreateSlide();
                  if (args.backgroundColor) newSlide.SetBackground(Api.CreateSolidFill(apiColor(args.backgroundColor)));
                  presentation.AddSlide(newSlide, args.index ? Math.max(0, Number(args.index) - 1) : undefined);
                  if (args.title) createTextBox(newSlide, {
                    text: args.title,
                    xMm: 15,
                    yMm: 15,
                    widthMm: 220,
                    heightMm: 30,
                    fontSize: args.titleFontSize || 28,
                    bold: true,
                  });
                  changed += 1;
                  results.push({ name: call.name, slideCount: slideCount() });
                  break;
                }

                case "slides_duplicate_slide": {
                  var slideToDuplicate = getSlide(args.slide);
                  var duplicated = slideToDuplicate.Duplicate();
                  if (!duplicated) throw new Error("复制幻灯片失败");
                  changed += 1;
                  results.push({ name: call.name, slide: Number(args.slide), slideCount: slideCount() });
                  break;
                }

                case "slides_delete_slide": {
                  if (slideCount() <= 1) throw new Error("演示文稿至少需要保留一页");
                  var slideToDelete = getSlide(args.slide);
                  var deleted = slideToDelete.Delete();
                  if (!deleted) throw new Error("删除幻灯片失败");
                  changed += 1;
                  results.push({ name: call.name, deletedSlide: Number(args.slide), slideCount: slideCount() });
                  break;
                }

                case "slides_add_textbox": {
                  var textboxSlideNumber = Number(args.slide || presentation.GetCurSlideIndex() + 1);
                  var textboxSlide = getSlide(textboxSlideNumber);
                  var textbox = createTextBox(textboxSlide, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: textboxSlideNumber,
                    object: describeDrawing(textbox, slideDrawings(textboxSlide).indexOf(textbox), false),
                  });
                  break;
                }

                case "slides_set_background": {
                  var backgroundSlide = getSlide(args.slide);
                  var mode = String(args.mode || "custom");
                  if (mode === "clear") {
                    if (typeof backgroundSlide.ClearBackground !== "function") throw new Error("当前 ONLYOFFICE 版本不支持清除背景");
                    backgroundSlide.ClearBackground();
                  } else if (mode === "layout") {
                    if (typeof backgroundSlide.FollowLayoutBackground !== "function") throw new Error("当前 ONLYOFFICE 版本不支持跟随版式背景");
                    backgroundSlide.FollowLayoutBackground();
                  } else if (mode === "master") {
                    if (typeof backgroundSlide.FollowMasterBackground !== "function") throw new Error("当前 ONLYOFFICE 版本不支持跟随母版背景");
                    backgroundSlide.FollowMasterBackground();
                  } else {
                    if (!hasOwn(args, "fill")) throw new Error("自定义幻灯片背景需要 fill");
                    backgroundSlide.SetBackground(createFill(args.fill));
                  }
                  changed += 1;
                  results.push({ name: call.name, slide: Number(args.slide), mode: mode });
                  break;
                }

                case "slides_add_shape": {
                  var shapeSlide = getSlide(args.slide);
                  var shapeType = String(args.shapeType || "rect");
                  var shapeWidth = Math.max(1, asFinite(args.widthMm, 100)) * EMU_PER_MM;
                  var shapeHeight = Math.max(1, asFinite(args.heightMm, 60)) * EMU_PER_MM;
                  var newShape = Api.CreateShape(
                    shapeType,
                    shapeWidth,
                    shapeHeight,
                    createFill(args.fill, args.fillColor),
                    createStroke(args.line, args.lineColor, args.lineWidthPt)
                  );
                  applyDrawingFrame(newShape, {
                    xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                    yMm: hasOwn(args, "yMm") ? args.yMm : 15,
                  });
                  applyShapeOptions(newShape, args);
                  shapeSlide.AddObject(newShape);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    object: describeDrawing(newShape, slideDrawings(shapeSlide).indexOf(newShape), Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_update_shape": {
                  var resolvedShape = resolveDrawing(args, "shape");
                  applyShapeOptions(resolvedShape.drawing, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    object: describeDrawing(resolvedShape.drawing, resolvedShape.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_delete_object": {
                  var resolvedObject = resolveDrawing(args);
                  var objectKind = drawingKind(resolvedObject.drawing);
                  var objectDeleted = typeof resolvedObject.drawing.Delete === "function"
                    ? resolvedObject.drawing.Delete()
                    : resolvedObject.slide.RemoveObject(resolvedObject.drawing);
                  if (objectDeleted === false) throw new Error("删除对象失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    deletedKind: objectKind,
                    deletedObjectIndex: resolvedObject.index,
                  });
                  break;
                }

                case "slides_inspect_charts": {
                  var chartStart = args.slide ? Number(args.slide) - 1 : 0;
                  var chartEnd = args.slide ? chartStart + 1 : slideCount();
                  if (chartStart < 0 || chartEnd > slideCount()) throw new Error("幻灯片页码超出范围");
                  var maxCharts = clamp(Math.round(asFinite(args.maxCharts, 100)), 1, 500);
                  var chartSlides = [];
                  var chartTotal = 0;
                  for (var chartSlideIndex = chartStart; chartSlideIndex < chartEnd && chartTotal < maxCharts; chartSlideIndex += 1) {
                    var inspectedChartSlide = presentation.GetSlideByIndex(chartSlideIndex);
                    var charts = slideCharts(inspectedChartSlide);
                    var describedCharts = [];
                    for (var chartIndex = 0; chartIndex < charts.length && chartTotal < maxCharts; chartIndex += 1) {
                      describedCharts.push(describeDrawing(charts[chartIndex], chartIndex, args.includeRaw !== false));
                      chartTotal += 1;
                    }
                    chartSlides.push({ slide: chartSlideIndex + 1, charts: describedCharts });
                  }
                  results.push({
                    name: call.name,
                    slides: chartSlides,
                    chartCount: chartTotal,
                    truncated: chartTotal >= maxCharts,
                  });
                  break;
                }

                case "slides_add_chart": {
                  if (!Array.isArray(args.series) || !args.series.length) throw new Error("slides_add_chart.series 必须是非空二维数组");
                  if (!Array.isArray(args.seriesNames) || !args.seriesNames.length) throw new Error("slides_add_chart.seriesNames 不能为空");
                  if (!Array.isArray(args.categories) || !args.categories.length) throw new Error("slides_add_chart.categories 不能为空");
                  var addChartSlide = getSlide(args.slide);
                  var chartWidth = Math.max(20, asFinite(args.widthMm, 160)) * EMU_PER_MM;
                  var chartHeight = Math.max(20, asFinite(args.heightMm, 90)) * EMU_PER_MM;
                  var newChart = Api.CreateChart(
                    normalizeChartType(args.type, "bar"),
                    args.series.map(function (values) { return values.map(Number); }),
                    args.seriesNames.map(String),
                    args.categories.map(String),
                    chartWidth,
                    chartHeight,
                    clamp(Math.round(asFinite(args.style, 2)), 1, 48),
                    Array.isArray(args.numFormats) ? args.numFormats.map(String) : []
                  );
                  if (!newChart) throw new Error("创建 PPT 图表失败");
                  var chartArgs = {};
                  for (var chartArgName in args) chartArgs[chartArgName] = args[chartArgName];
                  if (!hasOwn(chartArgs, "xMm")) chartArgs.xMm = 15;
                  if (!hasOwn(chartArgs, "yMm")) chartArgs.yMm = 40;
                  applyChartOptions(newChart, chartArgs, false);
                  addChartSlide.AddObject(newChart);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    chart: describeDrawing(newChart, slideCharts(addChartSlide).indexOf(newChart), Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_update_chart": {
                  var resolvedChart = resolveDrawing(args, "chart");
                  applyChartOptions(resolvedChart.drawing, args, false);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    chart: describeDrawing(resolvedChart.drawing, resolvedChart.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_delete_chart": {
                  var chartToDelete = resolveDrawing(args, "chart");
                  var chartDeleted = typeof chartToDelete.drawing.Delete === "function"
                    ? chartToDelete.drawing.Delete()
                    : chartToDelete.slide.RemoveObject(chartToDelete.drawing);
                  if (chartDeleted === false) throw new Error("删除 PPT 图表失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    deletedChartIndex: chartToDelete.index,
                  });
                  break;
                }

                default:
                  throw new Error("Slides Bridge 不支持工具：" + call.name);
              }
            }

            return JSON.stringify({
              ok: true,
              editorType: "slide",
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

  window.AICopilotBridges.slide = {
    execute: execute,
    inspect: function () {
      return execute([{ name: "slides_inspect", arguments: { maxChars: 12000 } }]);
    },
  };
})();
