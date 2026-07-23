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
            var mutatingNames = {
              slides_replace_text: true,
              slides_scale_font: true,
              slides_format_text: true,
              slides_format_selection: true,
              slides_add_slide: true,
              slides_duplicate_slide: true,
              slides_delete_slide: true,
              slides_add_textbox: true,
            };

            function getArgs(call) {
              return (call && (call.arguments || call.args)) || {};
            }

            function slideCount() {
              if (typeof presentation.GetSlidesCount === "function") return presentation.GetSlidesCount();
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

            function createTextBox(slide, args) {
              var width = Math.max(10, Number(args.widthMm) || 120) * 36000;
              var height = Math.max(5, Number(args.heightMm) || 30) * 36000;
              var fill = args.fillColor ? Api.CreateSolidFill(Api.RGB(
                parseInt(String(args.fillColor).replace("#", "").slice(0, 2), 16),
                parseInt(String(args.fillColor).replace("#", "").slice(2, 4), 16),
                parseInt(String(args.fillColor).replace("#", "").slice(4, 6), 16)
              )) : Api.CreateNoFill();
              var stroke = Api.CreateStroke(0, Api.CreateNoFill());
              var shape = Api.CreateShape("rect", width, height, fill, stroke);
              shape.SetPosition(Math.max(0, Number(args.xMm) || 15) * 36000, Math.max(0, Number(args.yMm) || 15) * 36000);
              var content = shape.GetContent();
              if (typeof content.RemoveAllElements === "function") content.RemoveAllElements();
              var paragraph = Api.CreateParagraph();
              paragraph.SetJc(String(args.align || "left"));
              var run = paragraph.AddText(String(args.text || ""));
              applyRunFormat(run, args);
              content.Push(paragraph);
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
                    var slide = presentation.GetSlideByIndex(slideIndex);
                    var shapes = slideShapes(slide);
                    var pieces = [];
                    for (var s = 0; s < shapes.length; s += 1) {
                      var content;
                      try { content = shapes[s].GetContent(); } catch (error) { continue; }
                      if (content && typeof content.GetText === "function") pieces.push(String(content.GetText() || ""));
                    }
                    var text = pieces.join("\n");
                    var remaining = Math.max(0, maxChars - totalChars);
                    slides.push({ slide: slideIndex + 1, text: text.slice(0, remaining), shapes: shapes.length });
                    totalChars += Math.min(text.length, remaining);
                    if (totalChars >= maxChars) break;
                  }
                  results.push({ name: call.name, slideCount: slideCount(), currentSlide: presentation.GetCurSlideIndex() + 1, slides: slides, truncated: totalChars >= maxChars });
                  break;
                }

                case "slides_replace_text": {
                  if (!args.search) throw new Error("slides_replace_text.search 不能为空");
                  var start = args.slide ? Number(args.slide) - 1 : 0;
                  var end = args.slide ? start + 1 : slideCount();
                  if (start < 0 || end > slideCount()) throw new Error("幻灯片页码超出范围");
                  var replacedRuns = 0;
                  for (var slideIndex = start; slideIndex < end; slideIndex += 1) {
                    var shapes = slideShapes(presentation.GetSlideByIndex(slideIndex));
                    for (var s = 0; s < shapes.length; s += 1) {
                      var runs = getRunsFromShape(shapes[s]);
                      for (var r = 0; r < runs.length; r += 1) {
                        var oldText = String(runs[r].GetText() || "");
                        var newText = args.matchCase
                          ? oldText.split(String(args.search)).join(String(args.replace || ""))
                          : oldText.replace(new RegExp(String(args.search).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), String(args.replace || ""));
                        if (newText === oldText) continue;
                        runs[r].RemoveAllElements();
                        runs[r].AddText(newText);
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
                  var start = args.slide ? Number(args.slide) - 1 : 0;
                  var end = args.slide ? start + 1 : slideCount();
                  if (start < 0 || end > slideCount()) throw new Error("幻灯片页码超出范围");
                  var scale = args.scale === undefined ? null : Number(args.scale);
                  if (call.name === "slides_scale_font" && (!isFinite(scale) || scale <= 0 || scale > 3)) {
                    throw new Error("slides_scale_font.scale 必须大于 0 且不超过 3");
                  }
                  var changedRuns = 0;
                  for (var slideIndex = start; slideIndex < end; slideIndex += 1) {
                    var shapes = slideShapes(presentation.GetSlideByIndex(slideIndex));
                    for (var s = 0; s < shapes.length; s += 1) {
                      var runs = getRunsFromShape(shapes[s]);
                      for (var r = 0; r < runs.length; r += 1) {
                        if (scale !== null) {
                          var oldSize = runs[r].GetFontSize();
                          if (typeof oldSize !== "number" || oldSize <= 0) continue;
                          runs[r].SetFontSize(Math.max(2, Math.round(oldSize * scale)));
                        } else {
                          applyRunFormat(runs[r], args);
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
                  var changedRuns = 0;
                  for (var s = 0; s < selectedShapes.length; s += 1) {
                    var runs = getRunsFromShape(selectedShapes[s]);
                    for (var r = 0; r < runs.length; r += 1) {
                      applyRunFormat(runs[r], args);
                      changedRuns += 1;
                    }
                  }
                  changed += changedRuns;
                  results.push({ name: call.name, selectedShapes: selectedShapes.length, changedRuns: changedRuns });
                  break;
                }

                case "slides_add_slide": {
                  var newSlide = Api.CreateSlide();
                  if (args.backgroundColor) newSlide.SetBackground(Api.CreateSolidFill(Api.Color(String(args.backgroundColor))));
                  presentation.AddSlide(newSlide, args.index ? Math.max(0, Number(args.index) - 1) : undefined);
                  if (args.title) createTextBox(newSlide, { text: args.title, xMm: 15, yMm: 15, widthMm: 220, heightMm: 30, fontSize: args.titleFontSize || 28, bold: true });
                  changed += 1;
                  results.push({ name: call.name, slideCount: slideCount() });
                  break;
                }

                case "slides_duplicate_slide": {
                  var slide = getSlide(args.slide);
                  var duplicated = slide.Duplicate();
                  if (!duplicated) throw new Error("复制幻灯片失败");
                  changed += 1;
                  results.push({ name: call.name, slide: Number(args.slide), slideCount: slideCount() });
                  break;
                }

                case "slides_delete_slide": {
                  if (slideCount() <= 1) throw new Error("演示文稿至少需要保留一页");
                  var slide = getSlide(args.slide);
                  var deleted = slide.Delete();
                  if (!deleted) throw new Error("删除幻灯片失败");
                  changed += 1;
                  results.push({ name: call.name, deletedSlide: Number(args.slide), slideCount: slideCount() });
                  break;
                }

                case "slides_add_textbox": {
                  var slide = getSlide(args.slide || presentation.GetCurSlideIndex() + 1);
                  createTextBox(slide, args);
                  changed += 1;
                  results.push({ name: call.name, slide: Number(args.slide || presentation.GetCurSlideIndex() + 1) });
                  break;
                }

                default:
                  throw new Error("Slides Bridge 不支持工具：" + call.name);
              }
            }

            return JSON.stringify({ ok: true, editorType: "slide", changed: changed, needsSave: mutating, results: results });
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
