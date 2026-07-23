(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) throw new Error((value && value.error) || "Word Bridge 执行失败");
    return value;
  }

  function getArgs(call) {
    return (call && (call.arguments || call.args)) || {};
  }

  function countSearchMatches(properties) {
    return new Promise(function (resolve, reject) {
      Asc.scope.copilotWordSearch = {
        search: properties.searchString,
        matchCase: properties.matchCase,
      };
      Asc.plugin.callCommand(
        function () {
          try {
            var search = String(Asc.scope.copilotWordSearch.search || "");
            var matchCase = Boolean(Asc.scope.copilotWordSearch.matchCase);
            var text = String(Api.GetDocument().GetText({
              ParaSeparator: "\n",
              TableCellSeparator: "\t",
              TableRowSeparator: "\n",
            }) || "");
            var haystack = matchCase ? text : text.toLowerCase();
            var needle = matchCase ? search : search.toLowerCase();
            var count = 0;
            var offset = 0;
            while (needle && (offset = haystack.indexOf(needle, offset)) !== -1) {
              count += 1;
              offset += needle.length;
            }
            return JSON.stringify({ ok: true, count: count });
          } catch (error) {
            return JSON.stringify({ ok: false, error: error && error.message ? error.message : String(error) });
          }
        },
        false,
        false,
        function (rawResult) {
          try {
            resolve(Number(parseResult(rawResult).count) || 0);
          } catch (error) {
            reject(error);
          }
        },
      );
    });
  }

  async function executeSearchAndReplace(call) {
    const args = getArgs(call);
    if (!args.search) throw new Error("word_replace_text.search 不能为空");
    const properties = {
      searchString: String(args.search),
      replaceString: String(args.replace || ""),
      matchCase: Boolean(args.matchCase),
    };
    const matches = await countSearchMatches(properties);
    if (!matches) {
      return {
        ok: true,
        editorType: "word",
        changed: 0,
        needsSave: false,
        results: [{
          name: call.name,
          replaced: false,
          replacedCount: 0,
          search: properties.searchString,
        }],
      };
    }

    await new Promise(function (resolve, reject) {
      try {
        const accepted = Asc.plugin.executeMethod("SearchAndReplace", [properties], function () {
          resolve();
        });
        if (accepted === false) reject(new Error("ONLYOFFICE 拒绝执行 SearchAndReplace"));
      } catch (error) {
        reject(error);
      }
    });

    return {
      ok: true,
      editorType: "word",
      changed: matches,
      needsSave: true,
      results: [{
        name: call.name,
        replaced: true,
        replacedCount: matches,
        search: properties.searchString,
      }],
    };
  }

  function mergeExecutionResult(target, source) {
    target.changed += Number(source && source.changed) || 0;
    target.needsSave = target.needsSave || Boolean(source && source.needsSave);
    if (source && Array.isArray(source.results)) target.results.push(...source.results);
  }

  function executeCallCommand(toolCalls) {
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
            var mutatingNames = {
              word_append_paragraph: true,
              word_insert_paragraph: true,
              word_format_document: true,
              word_format_selection: true,
              word_format_matches: true,
              word_delete_matches: true,
              word_add_hyperlink: true,
              word_add_comment: true,
              word_add_bookmark: true,
              word_format_paragraphs: true,
              word_set_paragraph_text: true,
              word_delete_paragraphs: true,
              word_set_list: true,
              word_insert_page_break: true,
              word_scale_font: true,
              word_add_table: true,
              word_set_table_cell: true,
              word_format_table: true,
              word_edit_table: true,
              word_set_page_layout: true,
              word_set_header_footer: true,
              word_set_document_text: true,
            };
            var textFormatKeys = [
              "fontSize", "fontFamily", "bold", "italic", "underline", "strikeout",
              "doubleStrikeout", "caps", "smallCaps", "color", "highlightColor",
              "characterSpacing", "vertAlign",
            ];
            var paragraphFormatKeys = textFormatKeys.concat([
              "align", "styleName", "headingLevel", "outlineLevel", "spacingBefore",
              "spacingAfter", "lineSpacing", "lineRule", "firstLineIndent", "leftIndent",
              "rightIndent", "keepLines", "keepNext", "widowControl",
              "contextualSpacing", "pageBreakBefore",
            ]);

            function hasDefined(args, keys) {
              for (var index = 0; index < keys.length; index += 1) {
                if (args[keys[index]] !== undefined) return true;
              }
              return false;
            }

            function commandArgs(call) {
              return (call && (call.arguments || call.args)) || {};
            }

            function finiteNumber(value, label) {
              var number = Number(value);
              if (!isFinite(number)) throw new Error(label + " 必须是有限数字");
              return number;
            }

            function pointsToTwips(value, label) {
              return Math.round(finiteNumber(value, label) * 20);
            }

            function mmToTwips(value, label) {
              return Math.round(finiteNumber(value, label) * 1440 / 25.4);
            }

            function allParagraphs() {
              return typeof doc.GetAllParagraphs === "function" ? doc.GetAllParagraphs() : [];
            }

            function allTables() {
              return typeof doc.GetAllTables === "function" ? doc.GetAllTables() : [];
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

            function applyTextFormat(target, format) {
              if (!target) return;
              if (format.fontSize !== undefined && typeof target.SetFontSize === "function") {
                target.SetFontSize(Math.max(2, Math.round(finiteNumber(format.fontSize, "fontSize") * 2)));
              }
              if (format.fontFamily && typeof target.SetFontFamily === "function") target.SetFontFamily(String(format.fontFamily));
              if (format.bold !== undefined && typeof target.SetBold === "function") target.SetBold(Boolean(format.bold));
              if (format.italic !== undefined && typeof target.SetItalic === "function") target.SetItalic(Boolean(format.italic));
              if (format.underline !== undefined && typeof target.SetUnderline === "function") target.SetUnderline(Boolean(format.underline));
              if (format.strikeout !== undefined && typeof target.SetStrikeout === "function") target.SetStrikeout(Boolean(format.strikeout));
              if (format.doubleStrikeout !== undefined && typeof target.SetDoubleStrikeout === "function") {
                target.SetDoubleStrikeout(Boolean(format.doubleStrikeout));
              }
              if (format.caps !== undefined && typeof target.SetCaps === "function") target.SetCaps(Boolean(format.caps));
              if (format.smallCaps !== undefined && typeof target.SetSmallCaps === "function") target.SetSmallCaps(Boolean(format.smallCaps));
              if (format.color && typeof target.SetColor === "function") target.SetColor(Api.Color(String(format.color)));
              if (format.highlightColor && typeof target.SetHighlight === "function") {
                target.SetHighlight(Api.Color(String(format.highlightColor)));
              }
              if (format.characterSpacing !== undefined && typeof target.SetSpacing === "function") {
                target.SetSpacing(pointsToTwips(format.characterSpacing, "characterSpacing"));
              }
              if (format.vertAlign && typeof target.SetVertAlign === "function") target.SetVertAlign(String(format.vertAlign));
            }

            function resolveStyle(styleName, styleType) {
              if (!styleName) return null;
              var style = typeof doc.GetStyle === "function" ? doc.GetStyle(String(styleName)) : null;
              if (!style && typeof doc.CreateStyle === "function" && String(styleName).indexOf("AI Bridge ") === 0) {
                style = doc.CreateStyle(String(styleName), styleType || "paragraph");
              }
              if (!style) throw new Error("Word 中不存在样式：" + styleName);
              return style;
            }

            function applyParagraphFormat(paragraph, format) {
              if (!paragraph) return;
              var styleName = format.styleName;
              if (!styleName && format.headingLevel !== undefined) {
                var headingLevel = Math.floor(finiteNumber(format.headingLevel, "headingLevel"));
                if (headingLevel < 1 || headingLevel > 9) throw new Error("headingLevel 必须在 1 到 9 之间");
                styleName = "Heading " + headingLevel;
              }
              if (styleName && typeof paragraph.SetStyle === "function") paragraph.SetStyle(resolveStyle(styleName, "paragraph"));
              if (format.align && typeof paragraph.SetJc === "function") paragraph.SetJc(String(format.align));
              if (format.spacingBefore !== undefined && typeof paragraph.SetSpacingBefore === "function") {
                paragraph.SetSpacingBefore(pointsToTwips(format.spacingBefore, "spacingBefore"));
              }
              if (format.spacingAfter !== undefined && typeof paragraph.SetSpacingAfter === "function") {
                paragraph.SetSpacingAfter(pointsToTwips(format.spacingAfter, "spacingAfter"));
              }
              if (format.lineSpacing !== undefined && typeof paragraph.SetSpacingLine === "function") {
                var lineRule = String(format.lineRule || "auto");
                var lineValue = lineRule === "auto"
                  ? Math.round(finiteNumber(format.lineSpacing, "lineSpacing") * 240)
                  : pointsToTwips(format.lineSpacing, "lineSpacing");
                paragraph.SetSpacingLine(lineValue, lineRule);
              }
              if (format.firstLineIndent !== undefined && typeof paragraph.SetIndFirstLine === "function") {
                paragraph.SetIndFirstLine(pointsToTwips(format.firstLineIndent, "firstLineIndent"));
              }
              if (format.leftIndent !== undefined && typeof paragraph.SetIndLeft === "function") {
                paragraph.SetIndLeft(pointsToTwips(format.leftIndent, "leftIndent"));
              }
              if (format.rightIndent !== undefined && typeof paragraph.SetIndRight === "function") {
                paragraph.SetIndRight(pointsToTwips(format.rightIndent, "rightIndent"));
              }
              if (format.keepLines !== undefined && typeof paragraph.SetKeepLines === "function") paragraph.SetKeepLines(Boolean(format.keepLines));
              if (format.keepNext !== undefined && typeof paragraph.SetKeepNext === "function") paragraph.SetKeepNext(Boolean(format.keepNext));
              if (format.widowControl !== undefined && typeof paragraph.SetWidowControl === "function") {
                paragraph.SetWidowControl(Boolean(format.widowControl));
              }
              if (format.contextualSpacing !== undefined && typeof paragraph.SetContextualSpacing === "function") {
                paragraph.SetContextualSpacing(Boolean(format.contextualSpacing));
              }
              if (format.pageBreakBefore !== undefined && typeof paragraph.SetPageBreakBefore === "function") {
                paragraph.SetPageBreakBefore(Boolean(format.pageBreakBefore));
              }
              if (format.outlineLevel !== undefined && typeof paragraph.SetOutlineLvl === "function") {
                var outlineLevel = Math.floor(finiteNumber(format.outlineLevel, "outlineLevel"));
                if (outlineLevel < 0 || outlineLevel > 8) throw new Error("outlineLevel 必须在 0 到 8 之间");
                paragraph.SetOutlineLvl(outlineLevel);
              }
              applyTextFormat(paragraph, format);
            }

            function paragraphText(paragraph) {
              return paragraph && typeof paragraph.GetText === "function" ? String(paragraph.GetText() || "") : "";
            }

            function selectParagraphEntries(args, allowCurrent) {
              var paragraphs = allParagraphs();
              var selected = [];
              var seen = {};
              var indexes = Array.isArray(args.paragraphIndexes) ? args.paragraphIndexes : [];

              function add(index) {
                if (index < 0 || index >= paragraphs.length) throw new Error("paragraphIndexes 包含超出范围的段落序号");
                if (seen[index]) return;
                seen[index] = true;
                selected.push({ paragraph: paragraphs[index], index: index });
              }

              if (args.all === true) {
                for (var allIndex = 0; allIndex < paragraphs.length; allIndex += 1) add(allIndex);
              }
              for (var indexPosition = 0; indexPosition < indexes.length; indexPosition += 1) {
                var oneBasedIndex = Math.floor(finiteNumber(indexes[indexPosition], "paragraphIndexes"));
                if (oneBasedIndex < 1) throw new Error("paragraphIndexes 必须从 1 开始");
                add(oneBasedIndex - 1);
              }
              if (args.search !== undefined && String(args.search) !== "") {
                var search = String(args.search);
                var matchCase = Boolean(args.matchCase);
                var matchMode = String(args.matchMode || "contains");
                var needle = matchCase ? search : search.toLowerCase();
                var matchingIndexes = [];
                for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
                  var text = paragraphText(paragraphs[paragraphIndex]);
                  var haystack = matchCase ? text : text.toLowerCase();
                  if ((matchMode === "exact" && haystack === needle) || (matchMode !== "exact" && haystack.indexOf(needle) !== -1)) {
                    matchingIndexes.push(paragraphIndex);
                  }
                }
                if (args.occurrence !== undefined) {
                  var occurrence = Math.floor(finiteNumber(args.occurrence, "occurrence"));
                  if (occurrence < 1) throw new Error("occurrence 必须从 1 开始");
                  if (matchingIndexes[occurrence - 1] !== undefined) add(matchingIndexes[occurrence - 1]);
                } else {
                  var maxParagraphs = Math.min(500, Math.max(1, Math.floor(Number(args.maxParagraphs) || 500)));
                  for (var matchIndex = 0; matchIndex < Math.min(maxParagraphs, matchingIndexes.length); matchIndex += 1) {
                    add(matchingIndexes[matchIndex]);
                  }
                }
              }
              if (!selected.length && allowCurrent && args.current === true && typeof doc.GetCurrentParagraph === "function") {
                var current = doc.GetCurrentParagraph();
                if (current) selected.push({ paragraph: current, index: -1 });
              }
              if (!selected.length) {
                throw new Error("必须使用 paragraphIndexes、search、all=true" + (allowCurrent ? " 或 current=true" : "") + " 指定段落");
              }
              return selected;
            }

            function selectSearchRanges(args) {
              if (!args.search) throw new Error("search 不能为空");
              var matches = typeof doc.Search === "function"
                ? doc.Search(String(args.search), Boolean(args.matchCase)) || []
                : [];
              var selected = [];
              if (args.occurrence !== undefined) {
                var occurrence = Math.floor(finiteNumber(args.occurrence, "occurrence"));
                if (occurrence < 1) throw new Error("occurrence 必须从 1 开始");
                if (matches[occurrence - 1]) selected.push(matches[occurrence - 1]);
              } else {
                var maxMatches = Math.min(500, Math.max(1, Math.floor(Number(args.maxMatches) || 500)));
                selected = matches.slice(0, maxMatches);
              }
              return { matches: matches, selected: selected };
            }

            function getTable(oneBasedIndex) {
              var tables = allTables();
              var tableIndex = Math.floor(finiteNumber(oneBasedIndex, "tableIndex"));
              if (tableIndex < 1 || tableIndex > tables.length) throw new Error("tableIndex 超出文档表格范围");
              return { table: tables[tableIndex - 1], index: tableIndex - 1, total: tables.length };
            }

            function getCell(table, oneBasedRow, oneBasedColumn) {
              var rowIndex = Math.floor(finiteNumber(oneBasedRow, "row")) - 1;
              var columnIndex = Math.floor(finiteNumber(oneBasedColumn, "column")) - 1;
              if (rowIndex < 0 || columnIndex < 0 || rowIndex >= table.GetRowsCount()) {
                throw new Error("表格行列序号超出范围");
              }
              var row = table.GetRow(rowIndex);
              if (!row || columnIndex >= row.GetCellsCount()) throw new Error("表格行列序号超出范围");
              return { row: row, cell: row.GetCell(columnIndex), rowIndex: rowIndex, columnIndex: columnIndex };
            }

            function applyCellFormat(cell, format) {
              if (format.backgroundColor && typeof cell.SetBackgroundColor === "function") {
                cell.SetBackgroundColor(Api.Color(String(format.backgroundColor)));
              }
              if (format.verticalAlign && typeof cell.SetVerticalAlign === "function") {
                cell.SetVerticalAlign(String(format.verticalAlign));
              }
              if (format.widthPercent !== undefined && typeof cell.SetWidth === "function") {
                cell.SetWidth("percent", Math.min(100, Math.max(1, finiteNumber(format.widthPercent, "widthPercent"))));
              }
              var content = cell.GetContent();
              var cellParagraphs = typeof content.GetAllParagraphs === "function" ? content.GetAllParagraphs() : [];
              for (var paragraphIndex = 0; paragraphIndex < cellParagraphs.length; paragraphIndex += 1) {
                applyParagraphFormat(cellParagraphs[paragraphIndex], format);
              }
            }

            function pageState() {
              var pageCount = typeof doc.GetPageCount === "function" ? Math.max(1, Number(doc.GetPageCount()) || 1) : 1;
              var currentIndex = typeof doc.GetCurrentPage === "function" ? Number(doc.GetCurrentPage()) : 0;
              if (!isFinite(currentIndex) || currentIndex < 0) currentIndex = 0;
              if (currentIndex >= pageCount) currentIndex = pageCount - 1;
              var visibleIndexes = typeof doc.GetCurrentVisiblePages === "function" ? doc.GetCurrentVisiblePages() || [] : [currentIndex];
              var visiblePages = [];
              for (var visibleIndex = 0; visibleIndex < visibleIndexes.length; visibleIndex += 1) {
                visiblePages.push(Number(visibleIndexes[visibleIndex]) + 1);
              }
              return {
                pageCount: pageCount,
                currentIndex: currentIndex,
                currentPage: currentIndex + 1,
                visiblePages: visiblePages,
              };
            }

            var hasMutatingCall = false;
            for (var mutatingIndex = 0; mutatingIndex < calls.length; mutatingIndex += 1) {
              if (mutatingNames[calls[mutatingIndex].name]) {
                hasMutatingCall = true;
                break;
              }
            }
            if (hasMutatingCall && typeof doc.CreateNewHistoryPoint === "function") doc.CreateNewHistoryPoint();

            for (var callIndex = 0; callIndex < calls.length; callIndex += 1) {
              var call = calls[callIndex];
              var args = commandArgs(call);

              switch (call.name) {
                case "word_inspect": {
                  var maxChars = Math.min(60000, Math.max(500, Number(args.maxChars) || 12000));
                  var text = String(doc.GetText({ ParaSeparator: "\n", TableCellSeparator: "\t", TableRowSeparator: "\n" }) || "");
                  var paragraphs = allParagraphs();
                  var tables = allTables();
                  var selection = typeof doc.GetRangeBySelect === "function" ? doc.GetRangeBySelect() : null;
                  var pages = pageState();
                  var maxParagraphs = Math.min(500, Math.max(0, Math.floor(Number(args.maxParagraphs) || 100)));
                  var paragraphDetails = [];
                  if (args.includeStructure !== false) {
                    for (var paragraphIndex = 0; paragraphIndex < Math.min(maxParagraphs, paragraphs.length); paragraphIndex += 1) {
                      var paragraph = paragraphs[paragraphIndex];
                      var style = typeof paragraph.GetStyle === "function" ? paragraph.GetStyle() : null;
                      paragraphDetails.push({
                        index: paragraphIndex + 1,
                        text: paragraphText(paragraph).slice(0, 500),
                        style: style && typeof style.GetName === "function" ? style.GetName() : null,
                        outlineLevel: typeof paragraph.GetOutlineLvl === "function" ? paragraph.GetOutlineLvl() : null,
                        align: typeof paragraph.GetJc === "function" ? paragraph.GetJc() : null,
                        inTable: typeof paragraph.GetParentTable === "function" ? Boolean(paragraph.GetParentTable()) : false,
                      });
                    }
                  }
                  var tableDetails = [];
                  for (var tableIndex = 0; tableIndex < Math.min(100, tables.length); tableIndex += 1) {
                    var table = tables[tableIndex];
                    var rows = typeof table.GetRowsCount === "function" ? table.GetRowsCount() : 0;
                    var columns = rows && table.GetRow(0) && typeof table.GetRow(0).GetCellsCount === "function"
                      ? table.GetRow(0).GetCellsCount()
                      : 0;
                    tableDetails.push({ index: tableIndex + 1, rows: rows, columns: columns });
                  }
                  var comments = [];
                  if (args.includeComments === true && typeof doc.GetAllComments === "function") {
                    var allComments = doc.GetAllComments() || [];
                    for (var commentIndex = 0; commentIndex < Math.min(200, allComments.length); commentIndex += 1) {
                      var comment = allComments[commentIndex];
                      comments.push({
                        id: typeof comment.GetId === "function" ? comment.GetId() : null,
                        text: typeof comment.GetText === "function" ? comment.GetText() : "",
                        author: typeof comment.GetAuthorName === "function" ? comment.GetAuthorName() : "",
                        solved: typeof comment.IsSolved === "function" ? comment.IsSolved() : false,
                      });
                    }
                  }
                  results.push({
                    name: call.name,
                    text: text.slice(0, maxChars),
                    truncated: text.length > maxChars,
                    characters: text.length,
                    paragraphs: paragraphs.length,
                    paragraphDetails: paragraphDetails,
                    paragraphDetailsTruncated: paragraphDetails.length < paragraphs.length,
                    tables: tables.length,
                    tableDetails: tableDetails,
                    comments: comments,
                    selectionText: selection && typeof selection.GetText === "function" ? String(selection.GetText() || "") : "",
                    pageCount: pages.pageCount,
                    currentPage: pages.currentPage,
                    visiblePages: pages.visiblePages,
                  });
                  break;
                }

                case "word_append_paragraph":
                case "word_insert_paragraph": {
                  if (args.text === undefined) throw new Error(call.name + ".text 不能为空");
                  var paragraph = Api.CreateParagraph();
                  var run = paragraph.AddText(String(args.text));
                  applyTextFormat(run, args);
                  applyParagraphFormat(paragraph, args);
                  if (args.listType) {
                    var numbering = doc.CreateNumbering(String(args.listType));
                    var listLevel = Math.min(8, Math.max(0, Math.floor(Number(args.listLevel) || 0)));
                    paragraph.SetNumbering(numbering.GetLevel(listLevel));
                  }
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
                  if (!hasDefined(args, paragraphFormatKeys)) {
                    throw new Error("word_format_document 至少需要一个格式属性");
                  }
                  var paragraphs = allParagraphs();
                  for (var p = 0; p < paragraphs.length; p += 1) applyParagraphFormat(paragraphs[p], args);
                  changed += paragraphs.length;
                  results.push({ name: call.name, formattedParagraphs: paragraphs.length });
                  break;
                }

                case "word_format_selection": {
                  if (!hasDefined(args, textFormatKeys)) {
                    throw new Error("word_format_selection 至少需要一个格式属性");
                  }
                  var range = typeof doc.GetRangeBySelect === "function" ? doc.GetRangeBySelect() : null;
                  if (!range) throw new Error("请先在 Word 中选择要格式化的文本");
                  applyTextFormat(range, args);
                  changed += 1;
                  results.push({ name: call.name, formattedSelection: true });
                  break;
                }

                case "word_format_matches": {
                  if (!hasDefined(args, textFormatKeys)) {
                    throw new Error("word_format_matches 至少需要一个格式属性");
                  }
                  var rangeSelection = selectSearchRanges(args);
                  for (var matchIndex = 0; matchIndex < rangeSelection.selected.length; matchIndex += 1) {
                    applyTextFormat(rangeSelection.selected[matchIndex], args);
                  }
                  changed += rangeSelection.selected.length;
                  results.push({
                    name: call.name,
                    search: String(args.search),
                    matches: rangeSelection.matches.length,
                    formattedMatches: rangeSelection.selected.length,
                    occurrence: args.occurrence === undefined ? null : Math.floor(Number(args.occurrence)),
                  });
                  break;
                }

                case "word_delete_matches": {
                  var deleteSelection = selectSearchRanges(args);
                  for (var deleteIndex = deleteSelection.selected.length - 1; deleteIndex >= 0; deleteIndex -= 1) {
                    deleteSelection.selected[deleteIndex].Delete();
                  }
                  changed += deleteSelection.selected.length;
                  results.push({
                    name: call.name,
                    search: String(args.search),
                    matches: deleteSelection.matches.length,
                    deletedMatches: deleteSelection.selected.length,
                  });
                  break;
                }

                case "word_add_hyperlink": {
                  if (!args.url) throw new Error("word_add_hyperlink.url 不能为空");
                  var hyperlinkSelection = selectSearchRanges(args);
                  for (var hyperlinkIndex = 0; hyperlinkIndex < hyperlinkSelection.selected.length; hyperlinkIndex += 1) {
                    hyperlinkSelection.selected[hyperlinkIndex].AddHyperlink(
                      String(args.url),
                      String(args.screenTip || args.url),
                      String(args.bookmarkName || ""),
                    );
                  }
                  changed += hyperlinkSelection.selected.length;
                  results.push({
                    name: call.name,
                    search: String(args.search),
                    linkedMatches: hyperlinkSelection.selected.length,
                    url: String(args.url),
                  });
                  break;
                }

                case "word_add_comment": {
                  if (!args.text) throw new Error("word_add_comment.text 不能为空");
                  var commentSelection = selectSearchRanges(args);
                  var commentIds = [];
                  for (var rangeIndex = 0; rangeIndex < commentSelection.selected.length; rangeIndex += 1) {
                    var addedComment = commentSelection.selected[rangeIndex].AddComment(
                      String(args.text),
                      args.author === undefined ? undefined : String(args.author),
                      args.userId === undefined ? undefined : String(args.userId),
                    );
                    commentIds.push(addedComment && typeof addedComment.GetId === "function" ? addedComment.GetId() : null);
                  }
                  changed += commentSelection.selected.length;
                  results.push({
                    name: call.name,
                    search: String(args.search),
                    commentedMatches: commentSelection.selected.length,
                    commentIds: commentIds,
                  });
                  break;
                }

                case "word_add_bookmark": {
                  if (!args.name) throw new Error("word_add_bookmark.name 不能为空");
                  if (args.occurrence === undefined) {
                    throw new Error("word_add_bookmark.occurrence 不能为空");
                  }
                  var bookmarkSelection = selectSearchRanges(args);
                  if (bookmarkSelection.selected.length !== 1) {
                    throw new Error("word_add_bookmark 必须通过 occurrence 精确指定一处文本");
                  }
                  var bookmarkAdded = bookmarkSelection.selected[0].AddBookmark(String(args.name));
                  if (bookmarkAdded === false) throw new Error("ONLYOFFICE 无法添加书签");
                  changed += 1;
                  results.push({ name: call.name, bookmark: String(args.name), added: true });
                  break;
                }

                case "word_format_paragraphs": {
                  if (!hasDefined(args, paragraphFormatKeys)) {
                    throw new Error("word_format_paragraphs 至少需要一个格式属性");
                  }
                  var paragraphEntries = selectParagraphEntries(args, true);
                  for (var entryIndex = 0; entryIndex < paragraphEntries.length; entryIndex += 1) {
                    applyParagraphFormat(paragraphEntries[entryIndex].paragraph, args);
                  }
                  changed += paragraphEntries.length;
                  results.push({
                    name: call.name,
                    formattedParagraphs: paragraphEntries.length,
                    paragraphIndexes: paragraphEntries.map(function (entry) { return entry.index < 0 ? null : entry.index + 1; }),
                  });
                  break;
                }

                case "word_set_paragraph_text": {
                  if (args.text === undefined) throw new Error("word_set_paragraph_text.text 不能为空");
                  var setTextEntries = selectParagraphEntries(args, true);
                  for (var setTextIndex = 0; setTextIndex < setTextEntries.length; setTextIndex += 1) {
                    var newRun = setTextEntries[setTextIndex].paragraph.SetText(String(args.text));
                    applyTextFormat(newRun, args);
                    applyParagraphFormat(setTextEntries[setTextIndex].paragraph, args);
                  }
                  changed += setTextEntries.length;
                  results.push({ name: call.name, updatedParagraphs: setTextEntries.length, characters: String(args.text).length });
                  break;
                }

                case "word_delete_paragraphs": {
                  var deleteEntries = selectParagraphEntries(args, true);
                  for (var paragraphDeleteIndex = deleteEntries.length - 1; paragraphDeleteIndex >= 0; paragraphDeleteIndex -= 1) {
                    deleteEntries[paragraphDeleteIndex].paragraph.Delete();
                  }
                  changed += deleteEntries.length;
                  results.push({ name: call.name, deletedParagraphs: deleteEntries.length });
                  break;
                }

                case "word_set_list": {
                  var listType = String(args.listType || "");
                  if (listType !== "bullet" && listType !== "numbered") {
                    throw new Error("word_set_list.listType 必须是 bullet 或 numbered");
                  }
                  var listEntries = selectParagraphEntries(args, true);
                  var numbering = doc.CreateNumbering(listType);
                  var listLevel = Math.min(8, Math.max(0, Math.floor(Number(args.level) || 0)));
                  var numberingLevel = numbering.GetLevel(listLevel);
                  for (var listIndex = 0; listIndex < listEntries.length; listIndex += 1) {
                    listEntries[listIndex].paragraph.SetNumbering(numberingLevel);
                    if (args.contextualSpacing !== undefined && typeof listEntries[listIndex].paragraph.SetContextualSpacing === "function") {
                      listEntries[listIndex].paragraph.SetContextualSpacing(Boolean(args.contextualSpacing));
                    }
                  }
                  changed += listEntries.length;
                  results.push({ name: call.name, listType: listType, level: listLevel, paragraphs: listEntries.length });
                  break;
                }

                case "word_insert_page_break": {
                  var breakEntries = selectParagraphEntries(args, true);
                  var position = String(args.position || "after");
                  for (var breakIndex = 0; breakIndex < breakEntries.length; breakIndex += 1) {
                    if (position === "before") breakEntries[breakIndex].paragraph.SetPageBreakBefore(true);
                    else if (position === "after") breakEntries[breakIndex].paragraph.AddPageBreak();
                    else throw new Error("word_insert_page_break.position 必须是 before 或 after");
                  }
                  changed += breakEntries.length;
                  results.push({ name: call.name, position: position, pageBreaks: breakEntries.length });
                  break;
                }

                case "word_navigate":
                case "word_scroll": {
                  var pages = pageState();
                  var target = call.name === "word_scroll"
                    ? String(args.direction || "down")
                    : String(args.target || (args.page !== undefined ? "page" : "end"));
                  var pageIndex = pages.currentIndex;
                  var selectedSearch = null;
                  if (call.name === "word_scroll") {
                    if (target !== "up" && target !== "down") throw new Error("word_scroll.direction 必须是 up 或 down");
                    var pageDelta = Math.min(100, Math.max(1, Math.floor(Number(args.pages) || 1)));
                    pageIndex += target === "up" ? -pageDelta : pageDelta;
                  } else if (target === "start") pageIndex = 0;
                  else if (target === "end") pageIndex = pages.pageCount - 1;
                  else if (target === "page") pageIndex = Math.floor(finiteNumber(args.page, "page")) - 1;
                  else if (target === "next") pageIndex += 1;
                  else if (target === "previous") pageIndex -= 1;
                  else if (target === "relative") pageIndex += Math.floor(finiteNumber(args.pageDelta, "pageDelta"));
                  else if (target === "current") pageIndex = pages.currentIndex;
                  else if (target === "search") {
                    selectedSearch = selectSearchRanges(args);
                    if (!selectedSearch.selected.length) {
                      results.push({
                        name: call.name,
                        target: target,
                        search: String(args.search),
                        matches: selectedSearch.matches.length,
                        moved: false,
                        currentPage: pages.currentPage,
                        pageCount: pages.pageCount,
                      });
                      break;
                    }
                    var searchRange = selectedSearch.selected[0];
                    searchRange.Select();
                    pageIndex = typeof searchRange.GetStartPage === "function" ? Number(searchRange.GetStartPage()) : pages.currentIndex;
                  } else {
                    throw new Error("word_navigate.target 必须是 start、end、page、next、previous、relative、current 或 search");
                  }
                  pageIndex = Math.min(pages.pageCount - 1, Math.max(0, pageIndex));
                  if (target !== "search") {
                    var moved = doc.GoToPage(pageIndex);
                    if (moved === false) throw new Error("ONLYOFFICE 无法跳转到指定页面");
                  }
                  results.push({
                    name: call.name,
                    target: target,
                    page: pageIndex + 1,
                    pageCount: pages.pageCount,
                    moved: true,
                    search: target === "search" ? String(args.search) : undefined,
                    matches: selectedSearch ? selectedSearch.matches.length : undefined,
                  });
                  break;
                }

                case "word_add_table": {
                  var rows = Math.min(100, Math.max(1, Math.floor(Number(args.rows) || 1)));
                  var cols = Math.min(50, Math.max(1, Math.floor(Number(args.cols) || 1)));
                  var table = Api.CreateTable(rows, cols);
                  table.SetWidth("percent", Math.min(100, Math.max(10, Number(args.widthPercent) || 100)));
                  if (args.styleName) table.SetStyle(resolveStyle(args.styleName, "table"));
                  else {
                    var borderedStyle = typeof doc.GetStyle === "function" ? doc.GetStyle("Bordered") : null;
                    if (borderedStyle) table.SetStyle(borderedStyle);
                  }
                  if (typeof table.SetTableLook === "function") {
                    table.SetTableLook(
                      Boolean(args.firstColumn),
                      args.firstRow !== false,
                      Boolean(args.lastColumn),
                      Boolean(args.lastRow),
                      args.horizontalBanding !== false,
                      Boolean(args.verticalBanding),
                    );
                  }
                  var data = Array.isArray(args.data) ? args.data : [];
                  for (var rowIndex = 0; rowIndex < rows; rowIndex += 1) {
                    for (var colIndex = 0; colIndex < cols; colIndex += 1) {
                      if (!data[rowIndex] || data[rowIndex][colIndex] === undefined) continue;
                      var cell = table.GetRow(rowIndex).GetCell(colIndex);
                      var cellRun = cell.SetText(String(data[rowIndex][colIndex]));
                      applyTextFormat(cellRun, args);
                    }
                  }
                  if (args.title && typeof table.SetTableTitle === "function") table.SetTableTitle(String(args.title));
                  if (args.description && typeof table.SetTableDescription === "function") table.SetTableDescription(String(args.description));
                  doc.Push(table);
                  changed += rows * cols;
                  results.push({ name: call.name, tableIndex: allTables().length, rows: rows, columns: cols });
                  break;
                }

                case "word_set_table_cell": {
                  if (!hasDefined(args, paragraphFormatKeys.concat(["text", "backgroundColor", "verticalAlign", "widthPercent"]))) {
                    throw new Error("word_set_table_cell 至少需要 text 或一个格式属性");
                  }
                  var tableEntry = getTable(args.tableIndex);
                  var cellEntry = getCell(tableEntry.table, args.row, args.column);
                  if (args.text !== undefined) {
                    var cellRun = cellEntry.cell.SetText(String(args.text));
                    applyTextFormat(cellRun, args);
                  }
                  applyCellFormat(cellEntry.cell, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    tableIndex: tableEntry.index + 1,
                    row: cellEntry.rowIndex + 1,
                    column: cellEntry.columnIndex + 1,
                    updated: true,
                  });
                  break;
                }

                case "word_format_table": {
                  if (!hasDefined(args, [
                    "widthPercent", "align", "styleName", "backgroundColor", "title",
                    "description", "firstRow", "lastRow", "firstColumn", "lastColumn",
                    "horizontalBanding", "verticalBanding", "fontSize", "fontFamily",
                    "bold", "italic", "color", "verticalAlign",
                  ])) {
                    throw new Error("word_format_table 至少需要一个格式属性");
                  }
                  var tableEntry = getTable(args.tableIndex);
                  var table = tableEntry.table;
                  if (args.widthPercent !== undefined) {
                    table.SetWidth("percent", Math.min(100, Math.max(1, finiteNumber(args.widthPercent, "widthPercent"))));
                  }
                  if (args.align && typeof table.SetJc === "function") table.SetJc(String(args.align));
                  if (args.styleName) table.SetStyle(resolveStyle(args.styleName, "table"));
                  if (args.backgroundColor && typeof table.SetBackgroundColor === "function") {
                    table.SetBackgroundColor(Api.Color(String(args.backgroundColor)));
                  }
                  if (args.title !== undefined && typeof table.SetTableTitle === "function") table.SetTableTitle(String(args.title));
                  if (args.description !== undefined && typeof table.SetTableDescription === "function") {
                    table.SetTableDescription(String(args.description));
                  }
                  if (typeof table.SetTableLook === "function" && (
                    args.firstColumn !== undefined || args.firstRow !== undefined || args.lastColumn !== undefined ||
                    args.lastRow !== undefined || args.horizontalBanding !== undefined || args.verticalBanding !== undefined
                  )) {
                    table.SetTableLook(
                      Boolean(args.firstColumn),
                      Boolean(args.firstRow),
                      Boolean(args.lastColumn),
                      Boolean(args.lastRow),
                      Boolean(args.horizontalBanding),
                      Boolean(args.verticalBanding),
                    );
                  }
                  var cells = typeof table.GetAllCells === "function" ? table.GetAllCells() : [];
                  for (var cellIndex = 0; cellIndex < cells.length; cellIndex += 1) applyCellFormat(cells[cellIndex], args);
                  changed += Math.max(1, cells.length);
                  results.push({ name: call.name, tableIndex: tableEntry.index + 1, formattedCells: cells.length });
                  break;
                }

                case "word_edit_table": {
                  var tableEntry = getTable(args.tableIndex);
                  var table = tableEntry.table;
                  var action = String(args.action || "");
                  if (action === "addRow") {
                    if (args.row === undefined) table.AddRow();
                    else {
                      var anchorRow = getCell(table, args.row, 1);
                      table.AddRow(anchorRow.cell, String(args.position || "after") === "after");
                    }
                  } else if (action === "addColumn") {
                    if (args.column === undefined) table.AddColumn();
                    else {
                      var anchorColumn = getCell(table, 1, args.column);
                      table.AddColumn(anchorColumn.cell, String(args.position || "after") === "after");
                    }
                  } else if (action === "removeRow") {
                    var removeRow = getCell(table, args.row, 1);
                    table.RemoveRow(removeRow.cell);
                  } else if (action === "removeColumn") {
                    var removeColumn = getCell(table, 1, args.column);
                    table.RemoveColumn(removeColumn.cell);
                  } else if (action === "mergeCells") {
                    var rowStart = Math.floor(finiteNumber(args.rowStart, "rowStart"));
                    var rowEnd = Math.floor(finiteNumber(args.rowEnd, "rowEnd"));
                    var columnStart = Math.floor(finiteNumber(args.columnStart, "columnStart"));
                    var columnEnd = Math.floor(finiteNumber(args.columnEnd, "columnEnd"));
                    if (rowStart > rowEnd || columnStart > columnEnd) throw new Error("合并区域起点不能大于终点");
                    var mergeCells = [];
                    for (var mergeRow = rowStart; mergeRow <= rowEnd; mergeRow += 1) {
                      for (var mergeColumn = columnStart; mergeColumn <= columnEnd; mergeColumn += 1) {
                        mergeCells.push(getCell(table, mergeRow, mergeColumn).cell);
                      }
                    }
                    if (mergeCells.length < 2 || !table.MergeCells(mergeCells)) throw new Error("ONLYOFFICE 无法合并指定单元格");
                  } else if (action === "splitCell") {
                    var splitCell = getCell(table, args.row, args.column);
                    var splitRows = Math.min(20, Math.max(1, Math.floor(Number(args.rows) || 1)));
                    var splitColumns = Math.min(20, Math.max(1, Math.floor(Number(args.columns) || 1)));
                    if (!splitCell.cell.Split(splitRows, splitColumns)) throw new Error("ONLYOFFICE 无法拆分指定单元格");
                  } else if (action === "clear") {
                    table.Clear();
                  } else if (action === "delete") {
                    table.Delete();
                  } else {
                    throw new Error("word_edit_table.action 不受支持");
                  }
                  changed += 1;
                  results.push({ name: call.name, tableIndex: tableEntry.index + 1, action: action, applied: true });
                  break;
                }

                case "word_set_page_layout": {
                  if (!hasDefined(args, [
                    "pageSize", "orientation", "widthMm", "heightMm", "marginLeftMm",
                    "marginTopMm", "marginRightMm", "marginBottomMm", "headerDistanceMm",
                    "footerDistanceMm", "titlePage",
                  ])) {
                    throw new Error("word_set_page_layout 至少需要一个页面属性");
                  }
                  var sections = typeof doc.GetSections === "function" ? doc.GetSections() : [];
                  if (!sections.length && typeof doc.GetFinalSection === "function") sections = [doc.GetFinalSection()];
                  var selectedSections = [];
                  if (args.sectionIndex !== undefined) {
                    var sectionIndex = Math.floor(finiteNumber(args.sectionIndex, "sectionIndex")) - 1;
                    if (sectionIndex < 0 || sectionIndex >= sections.length) throw new Error("sectionIndex 超出范围");
                    selectedSections.push(sections[sectionIndex]);
                  } else {
                    selectedSections = sections;
                  }
                  for (var sectionPosition = 0; sectionPosition < selectedSections.length; sectionPosition += 1) {
                    var section = selectedSections[sectionPosition];
                    if (
                      args.marginLeftMm !== undefined || args.marginTopMm !== undefined ||
                      args.marginRightMm !== undefined || args.marginBottomMm !== undefined
                    ) {
                      var leftMargin = args.marginLeftMm === undefined ? section.GetPageMarginLeft() : mmToTwips(args.marginLeftMm, "marginLeftMm");
                      var topMargin = args.marginTopMm === undefined ? section.GetPageMarginTop() : mmToTwips(args.marginTopMm, "marginTopMm");
                      var rightMargin = args.marginRightMm === undefined ? section.GetPageMarginRight() : mmToTwips(args.marginRightMm, "marginRightMm");
                      var bottomMargin = args.marginBottomMm === undefined ? section.GetPageMarginBottom() : mmToTwips(args.marginBottomMm, "marginBottomMm");
                      section.SetMargins(leftMargin, topMargin, rightMargin, bottomMargin);
                    }
                    if (args.pageSize || args.widthMm !== undefined || args.heightMm !== undefined || args.orientation) {
                      var widthMm;
                      var heightMm;
                      var pageSize = String(args.pageSize || "custom").toUpperCase();
                      if (pageSize === "A4") {
                        widthMm = 210;
                        heightMm = 297;
                      } else if (pageSize === "LETTER") {
                        widthMm = 215.9;
                        heightMm = 279.4;
                      } else if (pageSize === "LEGAL") {
                        widthMm = 215.9;
                        heightMm = 355.6;
                      } else {
                        widthMm = args.widthMm === undefined ? Number(section.GetPageWidth()) * 25.4 / 1440 : finiteNumber(args.widthMm, "widthMm");
                        heightMm = args.heightMm === undefined ? Number(section.GetPageHeight()) * 25.4 / 1440 : finiteNumber(args.heightMm, "heightMm");
                      }
                      var isPortrait = String(args.orientation || (heightMm >= widthMm ? "portrait" : "landscape")) !== "landscape";
                      var shortSide = Math.min(widthMm, heightMm);
                      var longSide = Math.max(widthMm, heightMm);
                      section.SetPageSize(
                        mmToTwips(isPortrait ? shortSide : longSide, "widthMm"),
                        mmToTwips(isPortrait ? longSide : shortSide, "heightMm"),
                        isPortrait,
                      );
                    }
                    if (args.headerDistanceMm !== undefined && typeof section.SetHeaderDistance === "function") {
                      section.SetHeaderDistance(mmToTwips(args.headerDistanceMm, "headerDistanceMm"));
                    }
                    if (args.footerDistanceMm !== undefined && typeof section.SetFooterDistance === "function") {
                      section.SetFooterDistance(mmToTwips(args.footerDistanceMm, "footerDistanceMm"));
                    }
                    if (args.titlePage !== undefined && typeof section.SetTitlePage === "function") {
                      section.SetTitlePage(Boolean(args.titlePage));
                    }
                  }
                  changed += selectedSections.length;
                  results.push({ name: call.name, updatedSections: selectedSections.length });
                  break;
                }

                case "word_set_header_footer": {
                  var kind = String(args.kind || "");
                  if (kind !== "header" && kind !== "footer") throw new Error("kind 必须是 header 或 footer");
                  var headerFooterType = String(args.type || "default");
                  var sections = typeof doc.GetSections === "function" ? doc.GetSections() : [];
                  if (!sections.length && typeof doc.GetFinalSection === "function") sections = [doc.GetFinalSection()];
                  var selectedSections = [];
                  if (args.sectionIndex !== undefined) {
                    var sectionIndex = Math.floor(finiteNumber(args.sectionIndex, "sectionIndex")) - 1;
                    if (sectionIndex < 0 || sectionIndex >= sections.length) throw new Error("sectionIndex 超出范围");
                    selectedSections.push(sections[sectionIndex]);
                  } else {
                    selectedSections = sections;
                  }
                  var action = String(args.action || "set");
                  if (action !== "set" && action !== "remove") {
                    throw new Error("word_set_header_footer.action 必须是 set 或 remove");
                  }
                  if (action === "set" && !hasDefined(args, paragraphFormatKeys.concat(["text", "pageNumber", "pagesCount"]))) {
                    throw new Error("word_set_header_footer 至少需要 text、页码字段或一个格式属性");
                  }
                  for (var sectionPosition = 0; sectionPosition < selectedSections.length; sectionPosition += 1) {
                    var section = selectedSections[sectionPosition];
                    if (action === "remove") {
                      if (kind === "header") section.RemoveHeader(headerFooterType);
                      else section.RemoveFooter(headerFooterType);
                      continue;
                    }
                    var content = kind === "header"
                      ? section.GetHeader(headerFooterType, true)
                      : section.GetFooter(headerFooterType, true);
                    if (args.text !== undefined && args.replace !== false && typeof content.SetText === "function") {
                      content.SetText(String(args.text));
                    }
                    var paragraph = typeof content.GetElement === "function" ? content.GetElement(0) : null;
                    if (!paragraph) {
                      paragraph = Api.CreateParagraph();
                      content.Push(paragraph);
                    }
                    if (args.replace === false && args.text !== undefined) paragraph.AddText(String(args.text));
                    if (args.pageNumber === true) paragraph.AddPageNumber();
                    if (args.pagesCount === true) paragraph.AddPagesCount();
                    applyParagraphFormat(paragraph, args);
                  }
                  changed += selectedSections.length;
                  results.push({
                    name: call.name,
                    kind: kind,
                    type: headerFooterType,
                    action: action,
                    updatedSections: selectedSections.length,
                  });
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

            return JSON.stringify({
              ok: true,
              editorType: "word",
              changed: changed,
              needsSave: changed > 0,
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

  async function execute(toolCalls) {
    const aggregate = {
      ok: true,
      editorType: "word",
      changed: 0,
      needsSave: false,
      results: [],
    };
    let callCommandBatch = [];

    async function flushCallCommandBatch() {
      if (!callCommandBatch.length) return;
      const batch = callCommandBatch;
      callCommandBatch = [];
      mergeExecutionResult(aggregate, await executeCallCommand(batch));
    }

    for (const call of toolCalls || []) {
      if (call && call.name === "word_replace_text") {
        await flushCallCommandBatch();
        mergeExecutionResult(aggregate, await executeSearchAndReplace(call));
      } else {
        callCommandBatch.push(call);
      }
    }
    await flushCallCommandBatch();
    return aggregate;
  }

  window.AICopilotBridges.word = {
    execute: execute,
    inspect: function () {
      return execute([{ name: "word_inspect", arguments: { maxChars: 12000, includeStructure: true } }]);
    },
  };
})();
