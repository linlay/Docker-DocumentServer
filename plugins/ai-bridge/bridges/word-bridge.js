(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) throw new Error((value && value.error) || "Word Bridge 执行失败");
    return value;
  }

  function execute(toolCalls) {
    return new Promise(function (resolve, reject) {
      Asc.scope.copilotWordCalls = toolCalls;
      Asc.plugin.info.recalculate = true;

      Asc.plugin.callCommand(
        function () {
          try {
            var calls = Asc.scope.copilotWordCalls || [];
            var doc = Api.GetDocument();
            var results = [];
            var changed = 0;
            var mutating = false;
            var mutatingNames = {
              word_replace_text: true,
              word_append_paragraph: true,
              word_insert_paragraph: true,
              word_format_document: true,
              word_format_selection: true,
              word_scale_font: true,
              word_add_table: true,
              word_set_document_text: true,
            };

            function getArgs(call) {
              return (call && (call.arguments || call.args)) || {};
            }

            function getRuns(paragraph) {
              var runs = [];
              if (!paragraph || typeof paragraph.GetElementsCount !== "function") return runs;
              var count = paragraph.GetElementsCount();
              for (var index = 0; index < count; index += 1) {
                var element = paragraph.GetElement(index);
                if (!element || typeof element.GetClassType !== "function") continue;
                if (element.GetClassType() === "run") runs.push(element);
              }
              return runs;
            }

            function allParagraphs() {
              return typeof doc.GetAllParagraphs === "function" ? doc.GetAllParagraphs() : [];
            }

            function applyRunFormat(run, format) {
              if (format.fontSize !== undefined) run.SetFontSize(Math.max(2, Math.round(Number(format.fontSize) * 2)));
              if (format.fontFamily) run.SetFontFamily(String(format.fontFamily));
              if (format.bold !== undefined) run.SetBold(Boolean(format.bold));
              if (format.italic !== undefined) run.SetItalic(Boolean(format.italic));
              if (format.underline !== undefined) run.SetUnderline(Boolean(format.underline));
              if (format.strikeout !== undefined) run.SetStrikeout(Boolean(format.strikeout));
              if (format.color) run.SetColor(Api.Color(String(format.color)));
            }

            for (var callIndex = 0; callIndex < calls.length; callIndex += 1) {
              if (mutatingNames[calls[callIndex].name]) {
                mutating = true;
                break;
              }
            }
            if (mutating && typeof doc.CreateNewHistoryPoint === "function") doc.CreateNewHistoryPoint();

            for (var callIndex = 0; callIndex < calls.length; callIndex += 1) {
              var call = calls[callIndex];
              var args = getArgs(call);

              switch (call.name) {
                case "word_inspect": {
                  var maxChars = Math.min(30000, Math.max(500, Number(args.maxChars) || 12000));
                  var text = doc.GetText({ ParaSeparator: "\n", TableCellSeparator: "\t", TableRowSeparator: "\n" });
                  results.push({
                    name: call.name,
                    text: String(text || "").slice(0, maxChars),
                    truncated: String(text || "").length > maxChars,
                    paragraphs: allParagraphs().length,
                  });
                  break;
                }

                case "word_replace_text": {
                  if (!args.search) throw new Error("word_replace_text.search 不能为空");
                  var replaced = doc.SearchAndReplace({
                    searchString: String(args.search),
                    replaceString: String(args.replace || ""),
                    matchCase: Boolean(args.matchCase),
                  });
                  changed += replaced ? 1 : 0;
                  results.push({ name: call.name, replaced: Boolean(replaced), search: String(args.search) });
                  break;
                }

                case "word_append_paragraph":
                case "word_insert_paragraph": {
                  if (args.text === undefined) throw new Error(call.name + ".text 不能为空");
                  var paragraph = Api.CreateParagraph();
                  var run = paragraph.AddText(String(args.text));
                  applyRunFormat(run, args);
                  if (args.align) paragraph.SetJc(String(args.align));
                  if (args.spacingBefore !== undefined) paragraph.SetSpacingBefore(Math.round(Number(args.spacingBefore) * 20));
                  if (args.spacingAfter !== undefined) paragraph.SetSpacingAfter(Math.round(Number(args.spacingAfter) * 20));
                  if (call.name === "word_insert_paragraph") doc.InsertContent([paragraph], Boolean(args.inline));
                  else doc.Push(paragraph);
                  changed += 1;
                  results.push({ name: call.name, inserted: true, characters: String(args.text).length });
                  break;
                }

                case "word_scale_font": {
                  var scale = Number(args.scale);
                  if (!isFinite(scale) || scale <= 0 || scale > 3) throw new Error("word_scale_font.scale 必须大于 0 且不超过 3");
                  var paragraphs = allParagraphs();
                  var changedRuns = 0;
                  for (var p = 0; p < paragraphs.length; p += 1) {
                    var runs = getRuns(paragraphs[p]);
                    for (var r = 0; r < runs.length; r += 1) {
                      var oldSize = runs[r].GetFontSize();
                      if (typeof oldSize !== "number" || oldSize <= 0) continue;
                      runs[r].SetFontSize(Math.max(2, Math.round(oldSize * scale)));
                      changedRuns += 1;
                    }
                  }
                  changed += changedRuns;
                  results.push({ name: call.name, scale: scale, changedRuns: changedRuns });
                  break;
                }

                case "word_format_document": {
                  var paragraphs = allParagraphs();
                  var changedRuns = 0;
                  for (var p = 0; p < paragraphs.length; p += 1) {
                    var paragraph = paragraphs[p];
                    if (args.align) paragraph.SetJc(String(args.align));
                    if (args.spacingBefore !== undefined) paragraph.SetSpacingBefore(Math.round(Number(args.spacingBefore) * 20));
                    if (args.spacingAfter !== undefined) paragraph.SetSpacingAfter(Math.round(Number(args.spacingAfter) * 20));
                    var runs = getRuns(paragraph);
                    for (var r = 0; r < runs.length; r += 1) {
                      applyRunFormat(runs[r], args);
                      changedRuns += 1;
                    }
                  }
                  changed += changedRuns;
                  results.push({ name: call.name, changedRuns: changedRuns, paragraphs: paragraphs.length });
                  break;
                }

                case "word_format_selection": {
                  var range = doc.GetRangeBySelect();
                  if (!range) throw new Error("请先在 Word 中选择要格式化的文本");
                  if (args.fontSize !== undefined) range.SetFontSize(Math.max(2, Math.round(Number(args.fontSize) * 2)));
                  if (args.fontFamily) range.SetFontFamily(String(args.fontFamily));
                  if (args.bold !== undefined) range.SetBold(Boolean(args.bold));
                  if (args.italic !== undefined) range.SetItalic(Boolean(args.italic));
                  if (args.underline !== undefined) range.SetUnderline(Boolean(args.underline));
                  if (args.strikeout !== undefined) range.SetStrikeout(Boolean(args.strikeout));
                  if (args.color) range.SetColor(Api.Color(String(args.color)));
                  changed += 1;
                  results.push({ name: call.name, formattedSelection: true });
                  break;
                }

                case "word_add_table": {
                  var rows = Math.min(100, Math.max(1, Math.floor(Number(args.rows) || 1)));
                  var cols = Math.min(50, Math.max(1, Math.floor(Number(args.cols) || 1)));
                  var table = Api.CreateTable(rows, cols);
                  table.SetWidth("percent", Math.min(100, Math.max(10, Number(args.widthPercent) || 100)));
                  var data = Array.isArray(args.data) ? args.data : [];
                  for (var rowIndex = 0; rowIndex < rows; rowIndex += 1) {
                    for (var colIndex = 0; colIndex < cols; colIndex += 1) {
                      if (!data[rowIndex] || data[rowIndex][colIndex] === undefined) continue;
                      table.GetRow(rowIndex).GetCell(colIndex).GetContent().GetElement(0).AddText(String(data[rowIndex][colIndex]));
                    }
                  }
                  doc.Push(table);
                  changed += rows * cols;
                  results.push({ name: call.name, rows: rows, cols: cols });
                  break;
                }

                case "word_set_document_text": {
                  if (args.text === undefined) throw new Error("word_set_document_text.text 不能为空");
                  doc.SetText(String(args.text));
                  changed += 1;
                  results.push({ name: call.name, replacedDocument: true, characters: String(args.text).length });
                  break;
                }

                default:
                  throw new Error("Word Bridge 不支持工具：" + call.name);
              }
            }

            return JSON.stringify({ ok: true, editorType: "word", changed: changed, needsSave: mutating, results: results });
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

  window.AICopilotBridges.word = {
    execute: execute,
    inspect: function () {
      return execute([{ name: "word_inspect", arguments: { maxChars: 12000 } }]);
    },
  };
})();
