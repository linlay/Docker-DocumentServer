(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) {
      const error = new Error((value && value.error) || "Word Bridge 执行失败");
      error.code = value && value.code ? value.code : "EXECUTION_FAILED";
      const details = value && value.details && typeof value.details === "object"
        ? { ...value.details }
        : {};
      if (!details.phase) details.phase = "word-command";
      error.details = details;
      throw error;
    }
    return value;
  }

  function getArgs(call) {
    return (call && (call.arguments || call.args)) || {};
  }

  const WORD_PROPERTY_ALIASES = {
    characterSpacing: "characterSpacingPt",
    spacingBefore: "spacingBeforePt",
    spacingAfter: "spacingAfterPt",
    firstLineIndent: "firstLineIndentPt",
    leftIndent: "leftIndentPt",
    rightIndent: "rightIndentPt",
  };
  const WORD_POINT_ALIAS_TOOLS = new Set([
    "word_append_paragraph",
    "word_insert_paragraph",
    "word_format_document",
    "word_format_selection",
    "word_format_matches",
    "word_format_paragraphs",
    "word_set_paragraph_text",
    "word_set_table_cell",
    "word_manage_style",
  ]);
  const WORD_ENUM_FIELDS = new Set([
    "action",
    "align",
    "appearance",
    "contentMode",
    "direction",
    "format",
    "insertAt",
    "kind",
    "legendPosition",
    "lineRule",
    "listType",
    "lock",
    "matchMode",
    "mode",
    "orientation",
    "pageSize",
    "position",
    "referenceKind",
    "rowHeightRule",
    "sides",
    "style",
    "target",
    "type",
    "valueType",
    "vertAlign",
    "verticalAlign",
    "wrapping",
  ]);
  const WORD_CANONICAL_ENUM_VALUES = [
    "A4", "Legal", "Letter", "acceptAll", "add", "addCaption", "addColumn",
    "addCrossReference", "addEndnote", "addFootnote", "addRow",
    "addTableOfFigures", "addToc", "after", "any", "append", "around",
    "atLeast", "auto", "bar", "baseline", "before", "behind", "block",
    "bookmark", "boolean", "both", "bottom", "boundingBox", "bullet",
    "caption", "center", "character", "chart", "check", "checkbox", "clear",
    "clearForms", "comboBox", "comments", "configure", "contains", "content",
    "continuous", "control", "create", "current", "custom", "date",
    "datePicker", "decimal", "default", "delete", "deleteAttribute",
    "deleteBookmark", "deleteElement", "down", "dropDown", "edit", "end",
    "endnote", "even", "evenPage", "exact", "first", "footer", "footnote",
    "forms", "header", "heading", "hidden", "image", "inFront", "inline",
    "insertAttribute", "insertElement", "landscape", "latex", "left",
    "mathml", "mergeCells", "multilevel", "next", "nextPage", "none",
    "number", "numbered", "numbering", "oddPage", "oleObject", "onlyoffice",
    "page", "paragraph", "picture", "portrait", "previous", "protect",
    "readOnly", "rejectAll", "relative", "remove", "removeAll",
    "removeColumn", "removeRow", "reopen", "replace", "reply", "resolve",
    "right", "search", "set", "setImage", "setText", "shape", "smartArt",
    "splitCell", "square", "start", "stop", "string", "subscript",
    "superscript", "table", "through", "tight", "top", "topAndBottom",
    "unicode", "unprotect", "up", "update", "updateAll", "updateAttribute",
    "updateElement", "updateTableOfFigures", "updateToc", "vba",
  ];
  const WORD_ENUM_CANONICAL_BY_TOKEN = WORD_CANONICAL_ENUM_VALUES.reduce(
    function (map, value) {
      map[enumToken(value)] = value;
      return map;
    },
    {},
  );

  function enumToken(value) {
    return String(value).toLowerCase().replace(/[\s_-]+/g, "");
  }

  function normalizeWordToolCalls(toolCalls) {
    const calls = JSON.parse(JSON.stringify(Array.isArray(toolCalls) ? toolCalls : []));
    const argumentNormalizations = [];

    function notice(toolCallIndex, path, canonicalPath, kind) {
      argumentNormalizations.push({
        toolCallIndex,
        path,
        canonicalPath,
        kind,
      });
    }

    function aliasProperty(target, alias, canonical, path, toolCallIndex) {
      if (!Object.prototype.hasOwnProperty.call(target, alias)) return;
      const canonicalWins = Object.prototype.hasOwnProperty.call(target, canonical);
      if (!canonicalWins) target[canonical] = target[alias];
      delete target[alias];
      notice(
        toolCallIndex,
        path + "." + alias,
        path + "." + canonical,
        canonicalWins ? "canonicalWins" : "propertyAlias",
      );
    }

    function normalizeEnums(value, field, path, toolCallIndex, toolName) {
      if (Array.isArray(value)) {
        return value.map(function (item, index) {
          return normalizeEnums(item, field, path + "[" + index + "]", toolCallIndex, toolName);
        });
      }
      if (value && typeof value === "object") {
        for (const key of Object.keys(value)) {
          value[key] = normalizeEnums(
            value[key],
            key,
            path + "." + key,
            toolCallIndex,
            toolName,
          );
        }
        return value;
      }
      if (typeof value !== "string" || !WORD_ENUM_FIELDS.has(field)) return value;
      const token = enumToken(value);
      let canonical = null;
      let kind = "enumCanonicalization";
      if (field === "align" && token === "centre") {
        canonical = "center";
        kind = "enumAlias";
      } else if (
        field === "align"
        && (token === "justify" || token === "justified")
        && toolName !== "word_add_table"
        && toolName !== "word_add_nested_table"
      ) {
        canonical = "both";
        kind = "enumAlias";
      } else if (field === "lineRule" && token === "multiple") {
        canonical = "auto";
        kind = "enumAlias";
      } else {
        canonical = WORD_ENUM_CANONICAL_BY_TOKEN[token] || null;
      }
      if (canonical && canonical !== value) {
        notice(toolCallIndex, path, path, kind);
        return canonical;
      }
      return value;
    }

    function normalizeCell(cell, path, toolCallIndex) {
      if (!cell || typeof cell !== "object" || Array.isArray(cell)) return;
      aliasProperty(cell, "textColor", "color", path, toolCallIndex);
      for (const alias of Object.keys(WORD_PROPERTY_ALIASES)) {
        aliasProperty(cell, alias, WORD_PROPERTY_ALIASES[alias], path, toolCallIndex);
      }
    }

    for (let index = 0; index < calls.length; index += 1) {
      const call = calls[index] || {};
      const toolName = String(call.name || "");
      const argumentKey = call.arguments ? "arguments" : (call.args ? "args" : "arguments");
      const args = call[argumentKey] || {};
      call[argumentKey] = args;
      if (!args || typeof args !== "object" || Array.isArray(args)) continue;
      if (toolName === "word_set_document_properties") {
        aliasProperty(args, "author", "creator", "arguments", index);
      }
      if (toolName === "word_set_page_layout" || toolName === "word_manage_section") {
        aliasProperty(args, "differentFirstPage", "titlePage", "arguments", index);
      }
      if (toolName === "word_add_table" || toolName === "word_add_nested_table") {
        aliasProperty(args, "columns", "cols", "arguments", index);
        const data = Array.isArray(args.data) ? args.data : [];
        for (let row = 0; row < data.length; row += 1) {
          if (!Array.isArray(data[row])) continue;
          for (let column = 0; column < data[row].length; column += 1) {
            normalizeCell(
              data[row][column],
              "arguments.data[" + row + "][" + column + "]",
              index,
            );
          }
        }
      }
      if (WORD_POINT_ALIAS_TOOLS.has(toolName)) {
        for (const alias of Object.keys(WORD_PROPERTY_ALIASES)) {
          aliasProperty(args, alias, WORD_PROPERTY_ALIASES[alias], "arguments", index);
        }
      }
      normalizeEnums(args, "", "arguments", index, toolName);
    }
    return { toolCalls: calls, argumentNormalizations };
  }

  const WORD_TABLE_CELL_FIELD_TYPES = {
    text: "string",
    fontSize: "number",
    fontFamily: "string",
    bold: "boolean",
    italic: "boolean",
    underline: "boolean",
    color: "color",
    textColor: "color",
    highlightColor: "color",
    characterSpacing: "number",
    characterSpacingPt: "number",
    backgroundColor: "color",
    align: "string",
    spacingBefore: "number",
    spacingAfter: "number",
    spacingBeforePt: "number",
    spacingAfterPt: "number",
    lineSpacing: "number",
    lineRule: "string",
    firstLineIndent: "number",
    leftIndent: "number",
    rightIndent: "number",
    firstLineIndentPt: "number",
    leftIndentPt: "number",
    rightIndentPt: "number",
    verticalAlign: "string",
    widthPercent: "number",
  };
  const WORD_TABLE_CELL_ENUMS = {
    align: ["left", "center", "right", "both"],
    lineRule: ["auto", "exact", "atLeast"],
    verticalAlign: ["top", "center", "bottom"],
  };

  function isPlainObject(value) {
    return Boolean(value) && typeof value === "object" && !Array.isArray(value);
  }

  function isTableCellScalar(value) {
    return value === undefined
      || value === null
      || typeof value === "string"
      || typeof value === "boolean"
      || (typeof value === "number" && Number.isFinite(value));
  }

  function collectTableDataValidationErrors(index, name, data, add) {
    if (data === undefined) return;
    if (!Array.isArray(data)) {
      add(index, name, "data", "data must be an array of rows", "type");
      return;
    }
    for (let rowIndex = 0; rowIndex < data.length; rowIndex += 1) {
      const row = data[rowIndex];
      if (!Array.isArray(row)) {
        add(index, name, "data[" + rowIndex + "]", "table row must be an array", "type");
        continue;
      }
      for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
        const cell = row[columnIndex];
        const cellPath = "data[" + rowIndex + "][" + columnIndex + "]";
        if (isTableCellScalar(cell)) continue;
        if (!isPlainObject(cell)) {
          add(index, name, cellPath, "table cell must be a scalar or formatting object", "type");
          continue;
        }
        if (!Object.prototype.hasOwnProperty.call(cell, "text")) {
          add(index, name, cellPath + ".text", "formatted table cell requires text", "required");
        } else if (typeof cell.text !== "string") {
          add(index, name, cellPath + ".text", "formatted table cell text must be a string", "type");
        }
        for (const field of Object.keys(cell)) {
          if (field === "text") continue;
          const expected = WORD_TABLE_CELL_FIELD_TYPES[field];
          if (!expected) {
            add(index, name, cellPath + "." + field, field + " is not allowed", "additionalProperties");
            continue;
          }
          const value = cell[field];
          if (expected === "number" && (typeof value !== "number" || !Number.isFinite(value))) {
            add(index, name, cellPath + "." + field, field + " must be a finite number", "type");
          } else if (expected === "boolean" && typeof value !== "boolean") {
            add(index, name, cellPath + "." + field, field + " must be a boolean", "type");
          } else if (expected === "string" && typeof value !== "string") {
            add(index, name, cellPath + "." + field, field + " must be a string", "type");
          } else if (
            expected === "color"
            && (typeof value !== "string" || !/^#?[0-9A-Fa-f]{6}$/.test(value))
          ) {
            add(index, name, cellPath + "." + field, field + " must be a six-digit hex color", "pattern");
          }
          if (
            WORD_TABLE_CELL_ENUMS[field]
            && typeof value === "string"
            && WORD_TABLE_CELL_ENUMS[field].indexOf(value) === -1
          ) {
            add(index, name, cellPath + "." + field, field + " has an unsupported value", "enum");
          }
          if (
            field === "widthPercent"
            && typeof value === "number"
            && (value < 1 || value > 100)
          ) {
            add(index, name, cellPath + "." + field, field + " must be between 1 and 100", "range");
          }
        }
        const pointFieldPairs = [
          ["characterSpacingPt", "characterSpacing", false],
          ["spacingBeforePt", "spacingBefore", false],
          ["spacingAfterPt", "spacingAfter", false],
          ["firstLineIndentPt", "firstLineIndent", true],
          ["leftIndentPt", "leftIndent", true],
          ["rightIndentPt", "rightIndent", true],
        ];
        for (const pair of pointFieldPairs) {
          const explicitField = pair[0];
          const legacyField = pair[1];
          const legacyValue = cell[legacyField];
          if (
            pair[2]
            && typeof legacyValue === "number"
            && Number.isFinite(legacyValue)
            && Math.abs(legacyValue) >= 300
            && Math.abs(legacyValue) <= 2880
            && Math.abs(legacyValue % 20) < 0.000001
          ) {
            add(
              index,
              name,
              cellPath + "." + legacyField,
              legacyField + " is point-valued and appears to contain twips",
              "semantic"
            );
          }
        }
      }
    }
  }

  function collectStaticValidationErrors(toolCalls) {
    const errors = [];
    const calls = Array.isArray(toolCalls) ? toolCalls : [];

    function add(index, name, path, message, keyword) {
      errors.push({
        toolCallIndex: index,
        tool: String(name || ""),
        path: "arguments." + path,
        keyword: keyword || "semantic",
        message,
      });
    }

    for (let index = 0; index < calls.length; index += 1) {
      const call = calls[index] || {};
      const name = String(call.name || "");
      const args = getArgs(call);
      if (!args || typeof args !== "object" || Array.isArray(args)) continue;

      if (name === "word_add_table" || name === "word_add_nested_table") {
        collectTableDataValidationErrors(index, name, args.data, add);
      }

      if (name === "word_manage_section" && String(args.action || "configure") === "create") {
        if (args.paragraphIndex === undefined) {
          add(index, name, "paragraphIndex", "paragraphIndex is required when action is create");
        }
      }

      if (name === "word_format_table_advanced") {
        if (args.repeatHeader !== undefined && args.row === undefined) {
          add(index, name, "row", "row is required when repeatHeader is provided");
        }
      }

      if (name === "word_manage_fields" && String(args.action || "") === "add") {
        if (typeof args.instruction !== "string" || !args.instruction.trim()) {
          add(index, name, "instruction", "instruction is required when action is add");
        }
      }

      if (name === "word_insert_page_break") {
        const hasTarget = (
          args.all === true
          || args.current === true
          || (Array.isArray(args.paragraphIndexes) && args.paragraphIndexes.length > 0)
          || (typeof args.search === "string" && args.search.length > 0)
        );
        if (!hasTarget) {
          add(index, name, "target", "paragraphIndexes, search, all=true, or current=true is required");
        }
      }

      if (name === "word_edit_table") {
        const action = String(args.action || "");
        const requiredByAction = {
          removeRow: ["row"],
          removeColumn: ["column"],
          splitCell: ["row", "column"],
          mergeCells: ["rowStart", "rowEnd", "columnStart", "columnEnd"],
        };
        const required = requiredByAction[action] || [];
        for (const field of required) {
          if (args[field] === undefined) {
            add(index, name, field, field + " is required when action is " + action);
          }
        }
        if (
          action === "mergeCells"
          && required.every(function (field) { return Number.isInteger(args[field]); })
        ) {
          if (args.rowStart > args.rowEnd) {
            add(index, name, "rowEnd", "rowEnd must be greater than or equal to rowStart");
          }
          if (args.columnStart > args.columnEnd) {
            add(index, name, "columnEnd", "columnEnd must be greater than or equal to columnStart");
          }
        }
      }

      if (name === "word_set_header_footer" && String(args.action || "set") === "set") {
        const contentFields = [
          "text",
          "pageNumber",
          "pagesCount",
          "fields",
          "fontSize",
          "fontFamily",
          "bold",
          "italic",
          "color",
          "align",
        ];
        if (!contentFields.some(function (field) { return args[field] !== undefined; })) {
          add(index, name, "action", "set requires text, a page field, fields, or formatting");
        }
      }
    }
    return errors;
  }

  function requireStaticValidation(toolCalls, preliminaryErrors) {
    const validationErrors = collectStaticValidationErrors(toolCalls).concat(
      preliminaryErrors || [],
    );
    if (!validationErrors.length) return;
    const error = new Error("Word 工具参数校验失败");
    error.code = "INVALID_TOOL_ARGUMENTS";
    error.details = {
      validationErrors,
      completedToolCalls: 0,
      partialMutationPossible: false,
    };
    throw error;
  }

  function collectUnsafeLegacyPointUnits(toolCalls) {
    const validationErrors = [];

    function visit(value, path, toolCallIndex, tool) {
      if (Array.isArray(value)) {
        for (let index = 0; index < value.length; index += 1) {
          visit(value[index], path + "[" + index + "]", toolCallIndex, tool);
        }
        return;
      }
      if (!value || typeof value !== "object") return;
      for (const field of Object.keys(value)) {
        const childPath = path + "." + field;
        const child = value[field];
        if (
          (field === "firstLineIndent" || field === "leftIndent" || field === "rightIndent")
          && !Object.prototype.hasOwnProperty.call(value, WORD_PROPERTY_ALIASES[field])
          && typeof child === "number"
          && Number.isFinite(child)
          && Math.abs(child) >= 300
          && Math.abs(child) <= 2880
          && Math.abs(child % 20) < 0.000001
        ) {
          validationErrors.push({
            toolCallIndex,
            tool,
            path: childPath,
            keyword: "semantic",
            message: field + " is point-valued and appears to contain twips",
          });
        }
        visit(child, childPath, toolCallIndex, tool);
      }
    }

    const calls = Array.isArray(toolCalls) ? toolCalls : [];
    for (let index = 0; index < calls.length; index += 1) {
      visit(getArgs(calls[index]), "arguments", index, String((calls[index] || {}).name || ""));
    }
    return validationErrors;
  }

  let nativeSearchAndReplaceSupported = null;

  function detectNativeSearchAndReplace() {
    if (nativeSearchAndReplaceSupported !== null) {
      return Promise.resolve(nativeSearchAndReplaceSupported);
    }
    return new Promise(function (resolve, reject) {
      Asc.plugin.callCommand(
        function () {
          try {
            return JSON.stringify({
              ok: true,
              supported: typeof Api.GetDocument().SearchAndReplace === "function",
            });
          } catch (error) {
            return JSON.stringify({ ok: false, error: error && error.message ? error.message : String(error) });
          }
        },
        false,
        false,
        function (rawResult) {
          try {
            nativeSearchAndReplaceSupported = Boolean(parseResult(rawResult).supported);
            resolve(nativeSearchAndReplaceSupported);
          } catch (error) {
            reject(error);
          }
        },
      );
    });
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

  async function executeEditingRestrictions(call) {
    const args = getArgs(call);
    const action = String(args.action || "");
    if (action !== "protect" && action !== "unprotect") {
      throw new Error("word_set_protection.action 必须是 protect 或 unprotect");
    }
    if (args.password !== undefined) {
      throw new Error("ONLYOFFICE 插件 API 的 SetEditingRestrictions 不支持密码保护");
    }
    const mode = action === "unprotect" ? "none" : String(args.mode || "readOnly");
    const supportedModes = new Set(["none", "comments", "forms", "readOnly"]);
    if (!supportedModes.has(mode)) {
      throw new Error("word_set_protection.mode 必须是 readOnly、comments 或 forms");
    }
    await new Promise(function (resolve, reject) {
      try {
        const accepted = Asc.plugin.executeMethod("SetEditingRestrictions", [mode], function () {
          resolve();
        });
        if (accepted === false) reject(new Error("ONLYOFFICE 拒绝设置文档编辑限制"));
      } catch (error) {
        reject(error);
      }
    });
    return {
      ok: true,
      editorType: "word",
      changed: 1,
      needsSave: true,
      results: [{
        name: call.name,
        action,
        restriction: mode,
      }],
    };
  }

  function executeMacroTool(call) {
    return new Promise(function (resolve, reject) {
      const args = getArgs(call);
      let method;
      let params;
      let needsSave = false;
      if (call.name === "word_inspect_macros") {
        method = args.kind === "vba" ? "GetVBAMacros" : "GetMacros";
        params = null;
      } else if (call.name === "word_set_macros") {
        if (!Array.isArray(args.macros)) {
          reject(new Error("word_set_macros.macros 必须是数组"));
          return;
        }
        const macrosArray = args.macros.map(function (macro) {
          const normalized = {
            name: String((macro && macro.name) || ""),
            value: String((macro && macro.value) || ""),
          };
          if (!normalized.name) throw new Error("宏名称不能为空");
          if (macro && macro.guid !== undefined) normalized.guid = String(macro.guid);
          return normalized;
        });
        const current = args.current === undefined ? (macrosArray.length ? 0 : -1) : Math.floor(Number(args.current));
        if (!Number.isFinite(current) || current < -1 || (macrosArray.length && current >= macrosArray.length)) {
          reject(new Error("word_set_macros.current 超出宏数组范围"));
          return;
        }
        if (!macrosArray.length && current !== -1) {
          reject(new Error("空宏集合的 current 必须是 -1"));
          return;
        }
        method = "SetMacros";
        params = [JSON.stringify({ macrosArray: macrosArray, current: current })];
        needsSave = true;
      } else {
        reject(new Error("未知 Word 宏工具：" + call.name));
        return;
      }
      if (!window.Asc || !Asc.plugin || typeof Asc.plugin.executeMethod !== "function") {
        reject(new Error("当前 ONLYOFFICE 版本不支持宏插件方法"));
        return;
      }
      try {
        const accepted = Asc.plugin.executeMethod(method, params, function (data) {
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
            editorType: "word",
            changed: needsSave ? 1 : 0,
            needsSave: needsSave,
            results: [{
              name: call.name,
              kind: method === "GetVBAMacros" ? "vba" : "onlyoffice",
              content: content,
            }],
          });
        });
        if (accepted === false) reject(new Error("ONLYOFFICE 拒绝执行宏方法 " + method));
      } catch (error) {
        reject(error);
      }
    });
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
              word_replace_text: true,
              word_append_paragraph: true,
              word_insert_paragraph: true,
              word_format_document: true,
              word_format_selection: true,
              word_format_matches: true,
              word_delete_matches: true,
              word_add_hyperlink: true,
              word_add_comment: true,
              word_add_bookmark: true,
              word_add_image: true,
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
              word_set_document_properties: true,
              word_manage_section: true,
              word_manage_style: true,
              word_set_tabs: true,
              word_set_numbering: true,
              word_format_table_advanced: true,
              word_add_nested_table: true,
              word_manage_drawing: true,
              word_add_shape: true,
              word_add_chart: true,
              word_add_math: true,
              word_add_ole_object: true,
              word_manage_fields: true,
              word_manage_long_document: true,
              word_manage_comments: true,
              word_manage_revisions: true,
              word_manage_content_control: true,
              word_manage_custom_xml: true,
              word_set_watermark: true,
            };
            var textFormatKeys = [
              "fontSize", "fontFamily", "bold", "italic", "underline", "strikeout",
              "doubleStrikeout", "caps", "smallCaps", "color", "highlightColor",
              "characterSpacing", "characterSpacingPt", "vertAlign", "characterStyleName",
            ];
            var paragraphFormatKeys = textFormatKeys.concat([
              "align", "styleName", "headingLevel", "outlineLevel", "spacingBefore",
              "spacingAfter", "spacingBeforePt", "spacingAfterPt", "lineSpacing", "lineRule",
              "firstLineIndent", "leftIndent", "rightIndent", "firstLineIndentPt",
              "leftIndentPt", "rightIndentPt", "keepLines", "keepNext", "widowControl",
              "contextualSpacing", "pageBreakBefore",
            ]);

            function hasDefined(args, keys) {
              for (var index = 0; index < keys.length; index += 1) {
                if (args[keys[index]] !== undefined) return true;
              }
              return false;
            }

            function tableCellSpec(value, tableArgs) {
              var format = {};
              var defaultKeys = ["fontSize", "fontFamily", "bold", "italic", "color"];
              for (var defaultIndex = 0; defaultIndex < defaultKeys.length; defaultIndex += 1) {
                var defaultKey = defaultKeys[defaultIndex];
                if (tableArgs[defaultKey] !== undefined) format[defaultKey] = tableArgs[defaultKey];
              }
              if (value && typeof value === "object" && !Array.isArray(value)) {
                for (var cellKey in value) {
                  if (cellKey !== "text" && Object.prototype.hasOwnProperty.call(value, cellKey)) {
                    format[cellKey] = value[cellKey];
                  }
                }
                return { text: String(value.text), format: format };
              }
              return { text: value === null || value === undefined ? "" : String(value), format: format };
            }

            function setTableCellValue(cell, value, tableArgs) {
              var spec = tableCellSpec(value, tableArgs);
              var run = cell.SetText(spec.text);
              applyTextFormat(run, spec.format);
              applyCellFormat(cell, spec.format);
            }

            function commandArgs(call) {
              return (call && (call.arguments || call.args)) || {};
            }

            function finiteNumber(value, label) {
              var number = Number(value);
              if (!isFinite(number)) throw new Error(label + " 必须是有限数字");
              return number;
            }

            function validDate(value, label) {
              var date = new Date(String(value));
              if (!isFinite(date.getTime())) throw new Error(label + " 必须是有效日期");
              return date;
            }

            function commandError(code, message, details) {
              var error = new Error(message);
              error.code = code;
              error.details = details;
              throw error;
            }

            function resolveImageSize(args, maxWidthMm, maxHeightMm) {
              var metadata = args._image;
              if (!metadata || !isFinite(Number(metadata.widthPx)) || !isFinite(Number(metadata.heightPx))) {
                commandError("INVALID_IMAGE_SOURCE", "内部图片资源缺少有效尺寸");
              }
              var ratio = Number(metadata.widthPx) / Number(metadata.heightPx);
              if (!isFinite(ratio) || ratio <= 0) commandError("INVALID_IMAGE_SOURCE", "内部图片宽高比无效");
              var hasWidth = args.widthMm !== undefined;
              var hasHeight = args.heightMm !== undefined;
              var width = hasWidth ? finiteNumber(args.widthMm, "widthMm") : null;
              var height = hasHeight ? finiteNumber(args.heightMm, "heightMm") : null;
              if ((hasWidth && width <= 0) || (hasHeight && height <= 0)) {
                throw new Error("图片宽高必须大于 0");
              }
              if (!hasWidth && !hasHeight) {
                width = maxWidthMm;
                height = width / ratio;
                if (height > maxHeightMm) {
                  height = maxHeightMm;
                  width = height * ratio;
                }
              } else if (hasWidth && !hasHeight) {
                height = width / ratio;
              } else if (!hasWidth && hasHeight) {
                width = height * ratio;
              } else if (args.preserveAspectRatio !== false) {
                var containedWidth = Math.min(width, height * ratio);
                var containedHeight = containedWidth / ratio;
                width = containedWidth;
                height = containedHeight;
              }
              if (width > 1000 || height > 1000) throw new Error("图片宽高不能超过 1000 mm");
              return { widthMm: width, heightMm: height };
            }

            function resolveImageTarget(args) {
              var targetModes = 0;
              if (args.current === true) targetModes += 1;
              if (args.paragraphIndex !== undefined) targetModes += 1;
              if (args.search !== undefined) targetModes += 1;
              if (targetModes > 1) {
                throw new Error("current、paragraphIndex、search 最多只能指定一种图片插入目标");
              }
              var paragraphs = allParagraphs();
              if (args.paragraphIndex !== undefined) {
                var paragraphIndex = Math.floor(finiteNumber(args.paragraphIndex, "paragraphIndex"));
                if (paragraphIndex < 1 || paragraphIndex > paragraphs.length) {
                  throw new Error("paragraphIndex 超出文档段落范围");
                }
                return { mode: "paragraph", paragraph: paragraphs[paragraphIndex - 1], paragraphIndex: paragraphIndex };
              }
              if (args.search !== undefined) {
                var search = String(args.search || "");
                if (!search) throw new Error("search 不能为空");
                var matchCase = Boolean(args.matchCase);
                var needle = matchCase ? search : search.toLowerCase();
                var matches = [];
                for (var index = 0; index < paragraphs.length; index += 1) {
                  var text = paragraphText(paragraphs[index]);
                  var haystack = matchCase ? text : text.toLowerCase();
                  if (haystack.indexOf(needle) !== -1) matches.push(index);
                }
                var occurrence = args.occurrence === undefined
                  ? 1
                  : Math.floor(finiteNumber(args.occurrence, "occurrence"));
                if (occurrence < 1) throw new Error("occurrence 必须从 1 开始");
                if (matches[occurrence - 1] === undefined) throw new Error("未找到指定的图片插入目标段落");
                var matchedIndex = matches[occurrence - 1];
                return { mode: "paragraph", paragraph: paragraphs[matchedIndex], paragraphIndex: matchedIndex + 1 };
              }
              return { mode: "current", paragraph: null, paragraphIndex: null };
            }

            function pointsToTwips(value, label) {
              return Math.round(finiteNumber(value, label) * 20);
            }

            function pointArgument(format, explicitName, legacyName, rejectLikelyTwips) {
              var hasExplicit = format[explicitName] !== undefined;
              var hasLegacy = format[legacyName] !== undefined;
              if (hasExplicit && hasLegacy) {
                throw new Error(explicitName + " 与旧字段 " + legacyName + " 不能同时提供");
              }
              if (hasExplicit) return finiteNumber(format[explicitName], explicitName);
              if (!hasLegacy) return null;
              var legacyValue = finiteNumber(format[legacyName], legacyName);
              if (
                rejectLikelyTwips
                && Math.abs(legacyValue) >= 300
                && Math.abs(legacyValue) <= 2880
                && Math.abs(legacyValue % 20) < 0.000001
              ) {
                throw new Error(
                  legacyName + " 的单位是磅（pt），值 " + legacyValue
                  + " 疑似误传 twips；请改用 " + explicitName + ": "
                  + (legacyValue / 20) + "，不要传入 twips"
                );
              }
              return legacyValue;
            }

            function mmToTwips(value, label) {
              return Math.round(finiteNumber(value, label) * 1440 / 25.4);
            }

            function allParagraphs() {
              return typeof doc.GetAllParagraphs === "function" ? doc.GetAllParagraphs() : [];
            }

            var virtualTables = null;
            var virtualTopLevelContent = null;

            function ensureVirtualTableState() {
              if (virtualTables === null) {
                var documentTables = typeof doc.GetAllTables === "function" ? doc.GetAllTables() || [] : [];
                virtualTables = Array.prototype.slice.call(documentTables);
              }
              if (virtualTopLevelContent === null) {
                var documentContent = typeof doc.GetContent === "function" ? doc.GetContent() || [] : [];
                virtualTopLevelContent = Array.prototype.slice.call(documentContent);
                if (
                  !virtualTopLevelContent.length &&
                  typeof doc.GetElementsCount === "function" &&
                  typeof doc.GetElement === "function"
                ) {
                  var documentElementCount = Math.max(0, Number(doc.GetElementsCount()) || 0);
                  for (var documentElementIndex = 0; documentElementIndex < documentElementCount; documentElementIndex += 1) {
                    virtualTopLevelContent.push(doc.GetElement(documentElementIndex));
                  }
                }
              }
            }

            function allTables() {
              if (virtualTables !== null) return virtualTables;
              return typeof doc.GetAllTables === "function" ? doc.GetAllTables() : [];
            }

            function isTableElement(element) {
              return Boolean(
                element &&
                typeof element.GetRowsCount === "function" &&
                typeof element.GetRow === "function"
              );
            }

            function tableCollectionInsertionIndex(topLevelPosition) {
              ensureVirtualTableState();
              var topLevelTablesBefore = 0;
              for (
                var contentIndex = 0;
                contentIndex < Math.min(topLevelPosition, virtualTopLevelContent.length);
                contentIndex += 1
              ) {
                if (isTableElement(virtualTopLevelContent[contentIndex])) topLevelTablesBefore += 1;
              }
              if (!topLevelTablesBefore) return 0;
              var seenTopLevelTables = 0;
              for (var tablePosition = 0; tablePosition < virtualTables.length; tablePosition += 1) {
                var candidateTable = virtualTables[tablePosition];
                var parentTable = typeof candidateTable.GetParentTable === "function"
                  ? candidateTable.GetParentTable()
                  : null;
                var parentControl = typeof candidateTable.GetParentContentControl === "function"
                  ? candidateTable.GetParentContentControl()
                  : null;
                if (parentTable || parentControl) continue;
                if (seenTopLevelTables === topLevelTablesBefore) return tablePosition;
                seenTopLevelTables += 1;
              }
              return virtualTables.length;
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
              if (format.characterStyleName !== undefined) {
                var characterStyle = resolveStyle(format.characterStyleName, "run");
                var targetClass = typeof target.GetClassType === "function" ? target.GetClassType() : "";
                if (targetClass === "paragraph") {
                  var paragraphRuns = getRuns(target);
                  for (var runIndex = 0; runIndex < paragraphRuns.length; runIndex += 1) {
                    requireMethod(paragraphRuns[runIndex], "SetStyle", "应用字符样式").call(
                      paragraphRuns[runIndex],
                      characterStyle,
                    );
                  }
                } else {
                  requireMethod(target, "SetStyle", "应用字符样式").call(target, characterStyle);
                }
              }
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
              var characterSpacingPt = pointArgument(format, "characterSpacingPt", "characterSpacing", false);
              if (characterSpacingPt !== null && typeof target.SetSpacing === "function") {
                target.SetSpacing(pointsToTwips(characterSpacingPt, "characterSpacingPt"));
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
              var spacingBeforePt = pointArgument(format, "spacingBeforePt", "spacingBefore", false);
              if (spacingBeforePt !== null && typeof paragraph.SetSpacingBefore === "function") {
                paragraph.SetSpacingBefore(pointsToTwips(spacingBeforePt, "spacingBeforePt"));
              }
              var spacingAfterPt = pointArgument(format, "spacingAfterPt", "spacingAfter", false);
              if (spacingAfterPt !== null && typeof paragraph.SetSpacingAfter === "function") {
                paragraph.SetSpacingAfter(pointsToTwips(spacingAfterPt, "spacingAfterPt"));
              }
              if (
                format.lineRule !== undefined
                && ["auto", "exact", "atLeast"].indexOf(String(format.lineRule)) === -1
              ) {
                throw new Error("lineRule 必须是 auto、exact 或 atLeast；多倍行距使用 auto");
              }
              if (format.lineSpacing !== undefined && typeof paragraph.SetSpacingLine === "function") {
                var lineRule = String(format.lineRule || "auto");
                var lineValue = lineRule === "auto"
                  ? Math.round(finiteNumber(format.lineSpacing, "lineSpacing") * 240)
                  : pointsToTwips(format.lineSpacing, "lineSpacing");
                paragraph.SetSpacingLine(lineValue, lineRule);
              }
              var firstLineIndentPt = pointArgument(format, "firstLineIndentPt", "firstLineIndent", true);
              if (firstLineIndentPt !== null && typeof paragraph.SetIndFirstLine === "function") {
                paragraph.SetIndFirstLine(pointsToTwips(firstLineIndentPt, "firstLineIndentPt"));
              }
              var leftIndentPt = pointArgument(format, "leftIndentPt", "leftIndent", true);
              if (leftIndentPt !== null && typeof paragraph.SetIndLeft === "function") {
                paragraph.SetIndLeft(pointsToTwips(leftIndentPt, "leftIndentPt"));
              }
              var rightIndentPt = pointArgument(format, "rightIndentPt", "rightIndent", true);
              if (rightIndentPt !== null && typeof paragraph.SetIndRight === "function") {
                paragraph.SetIndRight(pointsToTwips(rightIndentPt, "rightIndentPt"));
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
                  if (matchMode === "exact") text = text.replace(/[\r\n]+$/, "");
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

            function topLevelElementPosition(element, label, collectionKind, collectionIndex) {
              if (!element) throw new Error("未找到" + label);
              if (typeof element.GetParentTable === "function" && element.GetParentTable()) {
                throw new Error(label + "位于嵌套表格中，只支持顶层锚点");
              }
              if (typeof element.GetParentContentControl === "function" && element.GetParentContentControl()) {
                throw new Error(label + "位于内容控件中，只支持顶层锚点");
              }
              if (
                typeof doc.GetContent !== "function" &&
                (typeof doc.GetElementsCount !== "function" || typeof doc.GetElement !== "function")
              ) {
                commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持解析顶层文档锚点");
              }
              var topLevelContent = virtualTopLevelContent !== null
                ? virtualTopLevelContent
                : (typeof doc.GetContent === "function" ? doc.GetContent() || [] : []);
              var elementCount = topLevelContent.length || Math.max(0, Number(doc.GetElementsCount()) || 0);

              function topLevelElementAt(position) {
                return topLevelContent.length ? topLevelContent[position] : doc.GetElement(position);
              }

              function documentElementKind(documentElement) {
                if (
                  documentElement &&
                  typeof documentElement.GetRowsCount === "function" &&
                  typeof documentElement.GetRow === "function"
                ) {
                  return "table";
                }
                if (
                  documentElement &&
                  typeof documentElement.GetText === "function" &&
                  typeof documentElement.GetElementsCount === "function"
                ) {
                  return "paragraph";
                }
                return documentElement && typeof documentElement.GetClassType === "function"
                  ? documentElement.GetClassType()
                  : "";
              }

              if (typeof element.GetPosInParent === "function") {
                var parentPosition = Number(element.GetPosInParent());
                if (
                  isFinite(parentPosition) &&
                  Math.floor(parentPosition) === parentPosition &&
                  parentPosition >= 0 &&
                  parentPosition < elementCount
                ) {
                  return parentPosition;
                }
              }
              if (collectionKind === "paragraph" && typeof element.GetParaId === "function") {
                var paragraphId = element.GetParaId();
                for (var paragraphPosition = 0; paragraphPosition < elementCount; paragraphPosition += 1) {
                  var topLevelParagraph = topLevelElementAt(paragraphPosition);
                  if (
                    documentElementKind(topLevelParagraph) === "paragraph" &&
                    typeof topLevelParagraph.GetParaId === "function" &&
                    topLevelParagraph.GetParaId() === paragraphId
                  ) {
                    return paragraphPosition;
                  }
                }
              }
              if (
                (collectionKind === "paragraph" || collectionKind === "table") &&
                isFinite(Number(collectionIndex))
              ) {
                var collection = collectionKind === "paragraph" ? allParagraphs() : allTables();
                var targetCollectionIndex = Math.floor(Number(collectionIndex));
                var topLevelOrdinal = -1;
                for (var collectionPosition = 0; collectionPosition <= targetCollectionIndex; collectionPosition += 1) {
                  var collectionElement = collection[collectionPosition];
                  if (!collectionElement) continue;
                  var nestedInTable = typeof collectionElement.GetParentTable === "function"
                    && collectionElement.GetParentTable();
                  var nestedInControl = typeof collectionElement.GetParentContentControl === "function"
                    && collectionElement.GetParentContentControl();
                  if (!nestedInTable && !nestedInControl) topLevelOrdinal += 1;
                }
                if (topLevelOrdinal >= 0) {
                  var seenTopLevel = -1;
                  for (var documentPosition = 0; documentPosition < elementCount; documentPosition += 1) {
                    var documentElement = topLevelElementAt(documentPosition);
                    var documentClass = documentElementKind(documentElement);
                    if (documentClass !== collectionKind) continue;
                    seenTopLevel += 1;
                    if (seenTopLevel === topLevelOrdinal) return documentPosition;
                  }
                }
              }
              for (var elementIndex = 0; elementIndex < elementCount; elementIndex += 1) {
                if (topLevelElementAt(elementIndex) === element) return elementIndex;
              }
              throw new Error(label + "不是顶层文档元素，只支持顶层段落或顶层表格锚点");
            }

            function resolveTableInsertion(args) {
              var insertAt = String(args.insertAt || "end");
              if (
                insertAt !== "end" && insertAt !== "current" &&
                insertAt !== "before" && insertAt !== "after"
              ) {
                throw new Error("word_add_table.insertAt 必须是 end、current、before 或 after");
              }
              var hasParagraphAnchor = args.paragraphIndex !== undefined;
              var hasTableAnchor = args.tableIndex !== undefined;
              var hasSearchAnchor = args.search !== undefined;
              var anchorCount = Number(hasParagraphAnchor) + Number(hasTableAnchor) + Number(hasSearchAnchor);
              if (insertAt === "before" || insertAt === "after") {
                if (anchorCount !== 1) {
                  throw new Error("word_add_table 的 before/after 必须且只能提供 paragraphIndex、tableIndex 或 search 一个锚点");
                }
              } else if (anchorCount) {
                throw new Error("word_add_table 的 end/current 不接受 paragraphIndex、tableIndex 或 search 锚点");
              }
              if (insertAt === "end" || insertAt === "current") {
                return { insertAt: insertAt, position: null, anchor: null };
              }

              var anchorElement;
              var anchor;
              var anchorCollectionKind;
              var anchorCollectionIndex;
              if (hasParagraphAnchor) {
                var paragraphAnchor = paragraphByIndex(args.paragraphIndex, "paragraphIndex");
                anchorElement = paragraphAnchor.paragraph;
                anchor = { type: "paragraph", paragraphIndex: paragraphAnchor.index + 1 };
                anchorCollectionKind = "paragraph";
                anchorCollectionIndex = paragraphAnchor.index;
              } else if (hasTableAnchor) {
                var tableAnchor = getTable(args.tableIndex);
                anchorElement = tableAnchor.table;
                anchor = { type: "table", tableIndex: tableAnchor.index + 1 };
                anchorCollectionKind = "table";
                anchorCollectionIndex = tableAnchor.index;
              } else {
                var search = String(args.search || "");
                if (!search) throw new Error("word_add_table.search 不能为空");
                var matchMode = String(args.matchMode || "contains");
                if (matchMode !== "contains" && matchMode !== "exact") {
                  throw new Error("word_add_table.matchMode 必须是 contains 或 exact");
                }
                var occurrence = args.occurrence === undefined
                  ? 1
                  : Math.floor(finiteNumber(args.occurrence, "occurrence"));
                if (occurrence < 1) throw new Error("word_add_table.occurrence 必须从 1 开始");
                var matchCase = Boolean(args.matchCase);
                var needle = matchCase ? search : search.toLowerCase();
                var paragraphs = allParagraphs();
                var matchedParagraphs = [];
                for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
                  var text = paragraphText(paragraphs[paragraphIndex]);
                  var haystack = matchCase ? text : text.toLowerCase();
                  if (
                    (matchMode === "exact" && haystack === needle) ||
                    (matchMode === "contains" && haystack.indexOf(needle) !== -1)
                  ) {
                    matchedParagraphs.push({ paragraph: paragraphs[paragraphIndex], index: paragraphIndex });
                  }
                }
                if (!matchedParagraphs[occurrence - 1]) {
                  throw new Error("未找到 word_add_table 指定的搜索锚点");
                }
                var searchAnchor = matchedParagraphs[occurrence - 1];
                anchorElement = searchAnchor.paragraph;
                anchorCollectionKind = "paragraph";
                anchorCollectionIndex = searchAnchor.index;
                anchor = {
                  type: "search",
                  search: search,
                  matchCase: matchCase,
                  matchMode: matchMode,
                  occurrence: occurrence,
                  paragraphIndex: searchAnchor.index + 1,
                };
              }
              var anchorPosition = topLevelElementPosition(
                anchorElement,
                "word_add_table 锚点",
                anchorCollectionKind,
                anchorCollectionIndex,
              );
              anchor.position = anchorPosition;
              return {
                insertAt: insertAt,
                position: insertAt === "before" ? anchorPosition : anchorPosition + 1,
                anchor: anchor,
              };
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

            function safeCall(target, methodName, args, fallback) {
              if (!target || typeof target[methodName] !== "function") return fallback;
              try {
                var value = target[methodName].apply(target, args || []);
                return value === undefined ? fallback : value;
              } catch (error) {
                return fallback;
              }
            }

            function documentSections() {
              var sections = typeof doc.GetSections === "function" ? doc.GetSections() || [] : [];
              if (!sections.length && typeof doc.GetFinalSection === "function") {
                var finalSection = doc.GetFinalSection();
                if (finalSection) sections = [finalSection];
              }
              return sections;
            }

            function getSection(oneBasedIndex) {
              var sections = documentSections();
              var index = oneBasedIndex === undefined
                ? sections.length - 1
                : Math.floor(finiteNumber(oneBasedIndex, "sectionIndex")) - 1;
              if (index < 0 || index >= sections.length) throw new Error("sectionIndex 超出范围");
              return { section: sections[index], index: index, total: sections.length };
            }

            function paragraphByIndex(oneBasedIndex, label) {
              var paragraphs = allParagraphs();
              var index = Math.floor(finiteNumber(oneBasedIndex, label || "paragraphIndex")) - 1;
              if (index < 0 || index >= paragraphs.length) throw new Error((label || "paragraphIndex") + " 超出文档段落范围");
              return { paragraph: paragraphs[index], index: index };
            }

            function resolveInsertionParagraph(args) {
              if (args.paragraphIndex !== undefined) return paragraphByIndex(args.paragraphIndex, "paragraphIndex");
              if (args.current === true && typeof doc.GetCurrentParagraph === "function") {
                var currentParagraph = doc.GetCurrentParagraph();
                if (currentParagraph) return { paragraph: currentParagraph, index: -1 };
              }
              var paragraphs = allParagraphs();
              if (!paragraphs.length) {
                var createdParagraph = Api.CreateParagraph();
                doc.Push(createdParagraph);
                return { paragraph: createdParagraph, index: 0 };
              }
              return { paragraph: paragraphs[paragraphs.length - 1], index: paragraphs.length - 1 };
            }

            function addDrawingToParagraph(drawing, args) {
              var target = resolveInsertionParagraph(args || {});
              if (!target.paragraph || typeof target.paragraph.AddDrawing !== "function") {
                commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持在段落中插入图形对象");
              }
              if (target.paragraph.AddDrawing(drawing) === false) throw new Error("ONLYOFFICE 拒绝插入图形对象");
              return target.index < 0 ? null : target.index + 1;
            }

            function allDrawings() {
              return typeof doc.GetAllDrawingObjects === "function" ? doc.GetAllDrawingObjects() || [] : [];
            }

            function drawingKind(drawing) {
              var classType = safeCall(drawing, "GetClassType", [], "");
              if (classType === "image") return "image";
              if (classType === "shape") return "shape";
              if (classType === "chart") return "chart";
              if (classType === "oleObject") return "oleObject";
              if (classType === "smartArt") return "smartArt";
              return classType || "drawing";
            }

            function getDrawing(args) {
              var drawings = allDrawings();
              var kind = String(args.kind || "any");
              var matches = [];
              for (var index = 0; index < drawings.length; index += 1) {
                var drawing = drawings[index];
                if (kind !== "any" && drawingKind(drawing) !== kind) continue;
                if (args.name !== undefined && String(safeCall(drawing, "GetName", [], "")) !== String(args.name)) continue;
                matches.push({ drawing: drawing, index: index });
              }
              if (args.drawingIndex !== undefined) {
                var oneBasedIndex = Math.floor(finiteNumber(args.drawingIndex, "drawingIndex"));
                if (oneBasedIndex < 1 || oneBasedIndex > drawings.length) throw new Error("drawingIndex 超出范围");
                var selectedDrawing = drawings[oneBasedIndex - 1];
                if (kind !== "any" && drawingKind(selectedDrawing) !== kind) throw new Error("drawingIndex 对应对象类型与 kind 不匹配");
                if (args.name !== undefined && String(safeCall(selectedDrawing, "GetName", [], "")) !== String(args.name)) {
                  throw new Error("drawingIndex 与 name 未指向同一对象");
                }
                return { drawing: selectedDrawing, index: oneBasedIndex - 1 };
              }
              if (!matches.length) throw new Error("未找到指定图形对象");
              if (matches.length > 1) throw new Error("name 匹配多个图形对象，请改用 drawingIndex");
              return matches[0];
            }

            function hexComponents(value) {
              var hex = String(value || "#000000").replace(/^#/, "");
              if (/^[0-9a-fA-F]{3}$/.test(hex)) {
                hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
              }
              if (!/^[0-9a-fA-F]{6}$/.test(hex)) throw new Error("颜色必须是 #RRGGBB");
              return [
                parseInt(hex.slice(0, 2), 16),
                parseInt(hex.slice(2, 4), 16),
                parseInt(hex.slice(4, 6), 16),
              ];
            }

            function createStroke(format, fallbackColor) {
              var line = format || {};
              var width = line.widthPt === undefined ? 1 : finiteNumber(line.widthPt, "lineWidthPt");
              if (line.style === "none" || width <= 0) return Api.CreateStroke(0, Api.CreateNoFill());
              return Api.CreateStroke(
                Math.round(width * 12700),
                Api.CreateSolidFill(Api.Color(String(line.color || fallbackColor || "#000000"))),
              );
            }

            function setDrawingCommon(drawing, args) {
              if (args.widthMm !== undefined || args.heightMm !== undefined) {
                var width = args.widthMm === undefined ? safeCall(drawing, "GetWidth", [], null) : Math.round(finiteNumber(args.widthMm, "widthMm") * 36000);
                var height = args.heightMm === undefined ? safeCall(drawing, "GetHeight", [], null) : Math.round(finiteNumber(args.heightMm, "heightMm") * 36000);
                if (!(width > 0) || !(height > 0) || typeof drawing.SetSize !== "function") {
                  commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持修改图形尺寸");
                }
                drawing.SetSize(width, height);
              }
              if (args.rotationDeg !== undefined) {
                if (typeof drawing.SetRotation !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持图形旋转");
                drawing.SetRotation(finiteNumber(args.rotationDeg, "rotationDeg"));
              }
              if (args.wrapping !== undefined) {
                if (typeof drawing.SetWrappingStyle !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持文字环绕");
                drawing.SetWrappingStyle(String(args.wrapping));
              }
              if (args.xMm !== undefined) {
                if (typeof drawing.SetHorPosition !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持水平浮动定位");
                drawing.SetHorPosition(String(args.relativeFromH || "page"), Math.round(finiteNumber(args.xMm, "xMm") * 36000));
              } else if (args.alignH !== undefined && typeof drawing.SetHorAlign === "function") {
                drawing.SetHorAlign(String(args.relativeFromH || "page"), String(args.alignH));
              }
              if (args.yMm !== undefined) {
                if (typeof drawing.SetVerPosition !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持垂直浮动定位");
                drawing.SetVerPosition(String(args.relativeFromV || "page"), Math.round(finiteNumber(args.yMm, "yMm") * 36000));
              } else if (args.alignV !== undefined && typeof drawing.SetVerAlign === "function") {
                drawing.SetVerAlign(String(args.relativeFromV || "page"), String(args.alignV));
              }
              if (args.nameUpdate !== undefined && typeof drawing.SetName === "function") drawing.SetName(String(args.nameUpdate));
              if (args.border && typeof drawing.SetOutLine === "function") drawing.SetOutLine(createStroke(args.border, "#000000"));
            }

            function requireMethod(target, methodName, featureName) {
              if (!target || typeof target[methodName] !== "function") {
                commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持" + featureName);
              }
              return target[methodName];
            }

            function contentControlByArgs(args) {
              var controls;
              if (args.index !== undefined) {
                requireMethod(doc, "GetAllContentControls", "读取内容控件");
                controls = doc.GetAllContentControls() || [];
              } else if (args.tag !== undefined) {
                requireMethod(doc, "GetContentControlsByTag", "按标签查找内容控件");
                controls = doc.GetContentControlsByTag(String(args.tag)) || [];
              } else {
                requireMethod(doc, "GetAllContentControls", "读取内容控件");
                controls = doc.GetAllContentControls() || [];
              }
              if (args.index !== undefined) {
                var index = Math.floor(finiteNumber(args.index, "index")) - 1;
                if (index < 0 || index >= controls.length) throw new Error("index 超出内容控件范围");
                return { control: controls[index], index: index, total: controls.length };
              }
              if (args.tag !== undefined) {
                if (!controls.length) throw new Error("未找到 tag 对应的内容控件");
                if (controls.length > 1) throw new Error("tag 匹配多个内容控件，请同时提供 index");
                return { control: controls[0], index: 0, total: controls.length };
              }
              if (typeof doc.GetCurrentContentControl === "function") {
                var currentControl = doc.GetCurrentContentControl();
                if (currentControl) return { control: currentControl, index: -1, total: controls.length };
              }
              throw new Error("必须提供 index、tag，或把光标置于目标内容控件中");
            }

            function contentControlLock(value) {
              var lockMap = {
                none: "unlocked",
                content: "contentLocked",
                control: "sdtLocked",
                both: "sdtContentLocked",
              };
              if (!lockMap[value]) throw new Error("不支持的内容控件锁定类型：" + value);
              return lockMap[value];
            }

            function applyContentControlProperties(control, args) {
              if (args.tag !== undefined) {
                requireMethod(control, "SetTag", "设置内容控件标签").call(control, String(args.tag));
              }
              if (args.title !== undefined) {
                requireMethod(control, "SetAlias", "设置内容控件标题").call(control, String(args.title));
              }
              if (args.placeholder !== undefined) {
                requireMethod(control, "SetPlaceholderText", "设置内容控件占位文字").call(control, String(args.placeholder));
              }
              if (args.lock !== undefined) {
                requireMethod(control, "SetLock", "设置内容控件锁定状态").call(control, contentControlLock(String(args.lock)));
              }
              if (args.color !== undefined) {
                var controlColor = Api.Color(String(args.color));
                if (typeof control.SetColor === "function") control.SetColor(controlColor);
                else if (typeof control.SetBorderColor === "function") control.SetBorderColor(controlColor);
                else commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持设置内容控件颜色");
              }
              if (args.appearance !== undefined) {
                requireMethod(control, "SetAppearance", "设置内容控件外观").call(
                  control,
                  String(args.appearance),
                );
              }
              if (args.dateFormat !== undefined) {
                requireMethod(control, "SetDateFormat", "设置日期内容控件格式").call(
                  control,
                  String(args.dateFormat),
                );
              }
              if (args.date !== undefined) {
                requireMethod(control, "SetDate", "设置日期内容控件日期").call(
                  control,
                  validDate(args.date, "date"),
                );
              }
              if (args.dataBinding !== undefined) {
                var binding = args.dataBinding || {};
                if (!String(binding.storeItemId || "")) {
                  throw new Error("dataBinding.storeItemId 不能为空");
                }
                if (!String(binding.xpath || "")) {
                  throw new Error("dataBinding.xpath 不能为空");
                }
                requireMethod(control, "SetDataBinding", "设置内容控件 XML 数据绑定").call(control, {
                  prefixMapping: binding.prefixMapping === undefined ? "" : String(binding.prefixMapping),
                  storeItemID: String(binding.storeItemId),
                  xpath: String(binding.xpath),
                });
              }
              if (args.updateFromXml === true) {
                requireMethod(control, "UpdateFromXmlMapping", "从自定义 XML 更新内容控件").call(control);
              }
              if (args.text !== undefined) {
                requireMethod(control, "RemoveAllElements", "替换内容控件文本").call(control);
                requireMethod(control, "AddText", "写入内容控件文本").call(control, String(args.text));
              }
            }

            function resolveInlineContentControlTarget(args, contentMode) {
              var hasTableIndex = args.tableIndex !== undefined;
              var hasRow = args.row !== undefined;
              var hasColumn = args.column !== undefined;
              var hasAnyCellCoordinate = hasTableIndex || hasRow || hasColumn;
              if (hasAnyCellCoordinate && !(hasTableIndex && hasRow && hasColumn)) {
                throw new Error("inline 内容控件的 tableIndex、row、column 必须同时提供");
              }
              var targetCount = Number(hasAnyCellCoordinate) +
                Number(args.paragraphIndex !== undefined) +
                Number(args.current === true);
              if (targetCount > 1) {
                throw new Error("inline 内容控件的段落、当前段落和表格单元格目标互斥");
              }
              if (contentMode === "replace" && targetCount !== 1) {
                throw new Error("inline 内容控件的 contentMode=replace 必须指定 paragraphIndex、current 或表格单元格目标");
              }

              if (hasAnyCellCoordinate) {
                var tableTarget = getTable(args.tableIndex);
                var cellTarget = getCell(tableTarget.table, args.row, args.column);
                if (contentMode === "replace") {
                  if (requireMethod(cellTarget.cell, "SetText", "清空内容控件目标单元格").call(
                    cellTarget.cell,
                    "",
                  ) === false) {
                    throw new Error("ONLYOFFICE 无法清空内容控件目标单元格");
                  }
                }
                var cellContent = requireMethod(
                  cellTarget.cell,
                  "GetContent",
                  "访问内容控件目标单元格",
                ).call(cellTarget.cell);
                var cellParagraphs = cellContent && typeof cellContent.GetAllParagraphs === "function"
                  ? cellContent.GetAllParagraphs() || []
                  : [];
                var cellParagraph = cellParagraphs.length
                  ? cellParagraphs[cellParagraphs.length - 1]
                  : (cellContent && typeof cellContent.GetElement === "function" ? cellContent.GetElement(0) : null);
                if (!cellParagraph) throw new Error("ONLYOFFICE 无法取得内容控件目标单元格段落");
                return {
                  paragraph: cellParagraph,
                  target: {
                    type: "cell",
                    tableIndex: tableTarget.index + 1,
                    row: cellTarget.rowIndex + 1,
                    column: cellTarget.columnIndex + 1,
                  },
                };
              }

              var paragraphTarget;
              var target;
              if (args.paragraphIndex !== undefined) {
                paragraphTarget = paragraphByIndex(args.paragraphIndex, "paragraphIndex");
                target = { type: "paragraph", paragraphIndex: paragraphTarget.index + 1 };
              } else if (args.current === true) {
                var currentTargetParagraph = typeof doc.GetCurrentParagraph === "function"
                  ? doc.GetCurrentParagraph()
                  : null;
                if (!currentTargetParagraph) {
                  throw new Error("ONLYOFFICE 无法取得当前段落");
                }
                paragraphTarget = { paragraph: currentTargetParagraph, index: -1 };
                target = { type: "current" };
              } else {
                paragraphTarget = resolveInsertionParagraph(args);
                target = {
                  type: "paragraph",
                  paragraphIndex: paragraphTarget.index < 0 ? null : paragraphTarget.index + 1,
                };
              }
              if (contentMode === "replace") {
                if (requireMethod(
                  paragraphTarget.paragraph,
                  "RemoveAllElements",
                  "清空内容控件目标段落",
                ).call(paragraphTarget.paragraph) === false) {
                  throw new Error("ONLYOFFICE 无法清空内容控件目标段落");
                }
              }
              return { paragraph: paragraphTarget.paragraph, target: target };
            }

            function createContentControl(args) {
              var kind = String(args.kind || "inline");
              var contentMode = String(args.contentMode || "append");
              if (contentMode !== "append" && contentMode !== "replace") {
                throw new Error("contentMode 必须是 append 或 replace");
              }
              var hasCellTarget = args.tableIndex !== undefined || args.row !== undefined || args.column !== undefined;
              if (kind !== "inline" && contentMode === "replace") {
                throw new Error("contentMode=replace 仅支持 inline 内容控件");
              }
              if (kind !== "inline" && hasCellTarget) {
                throw new Error("表格单元格目标仅支持 inline 内容控件");
              }
              var control;
              var resolvedTarget = null;
              var items = Array.isArray(args.items) ? args.items : [];
              var list = [];
              for (var itemIndex = 0; itemIndex < items.length; itemIndex += 1) {
                var item = items[itemIndex] || {};
                list.push({
                  display: String(item.display || ""),
                  value: item.value === undefined ? String(item.display || "") : String(item.value),
                });
              }
              if (kind === "block") {
                requireMethod(Api, "CreateBlockLvlSdt", "创建块级内容控件");
                control = Api.CreateBlockLvlSdt();
                var controlContent = requireMethod(control, "GetContent", "访问块级内容控件内容").call(control);
                var controlParagraph = controlContent && typeof controlContent.GetElement === "function"
                  ? controlContent.GetElement(0)
                  : null;
                if (args.text !== undefined && controlParagraph && typeof controlParagraph.SetText === "function") {
                  controlParagraph.SetText(String(args.text));
                }
                var blockInsertIndex = typeof doc.GetElementsCount === "function" ? doc.GetElementsCount() : 0;
                if (requireMethod(doc, "AddElement", "插入块级内容控件").call(doc, blockInsertIndex, control) === false) {
                  throw new Error("ONLYOFFICE 无法插入块级内容控件");
                }
              } else if (kind === "inline") {
                requireMethod(Api, "CreateInlineLvlSdt", "创建行内内容控件");
                resolvedTarget = resolveInlineContentControlTarget(args, contentMode);
                control = Api.CreateInlineLvlSdt();
                if (args.text !== undefined) control.AddText(String(args.text));
                requireMethod(resolvedTarget.paragraph, "AddInlineLvlSdt", "插入行内内容控件").call(
                  resolvedTarget.paragraph,
                  control,
                );
              } else {
                if (args.paragraphIndex !== undefined) {
                  var selectedTarget = paragraphByIndex(args.paragraphIndex, "paragraphIndex");
                  if (typeof selectedTarget.paragraph.Select === "function") selectedTarget.paragraph.Select();
                }
                if (kind === "checkbox") {
                  control = requireMethod(doc, "AddCheckBoxContentControl", "创建复选框内容控件").call(doc);
                } else if (kind === "comboBox") {
                  control = requireMethod(doc, "AddComboBoxContentControl", "创建组合框内容控件").call(
                    doc,
                    list,
                    args.selectedValue === undefined ? undefined : String(args.selectedValue),
                  );
                } else if (kind === "dropDown") {
                  control = requireMethod(doc, "AddDropDownListContentControl", "创建下拉列表内容控件").call(
                    doc,
                    list,
                    args.selectedValue === undefined ? undefined : String(args.selectedValue),
                  );
                } else if (kind === "datePicker") {
                  control = requireMethod(doc, "AddDatePickerContentControl", "创建日期内容控件").call(doc);
                } else if (kind === "picture") {
                  control = requireMethod(doc, "AddPictureContentControl", "创建图片内容控件").call(doc);
                } else {
                  throw new Error("不支持的内容控件类型：" + kind);
                }
              }
              if (!control) throw new Error("ONLYOFFICE 无法创建内容控件");
              if (kind !== "block" && kind !== "inline" && args.text !== undefined && typeof control.AddText === "function") {
                control.AddText(String(args.text));
              }
              applyContentControlProperties(control, {
                tag: args.tag,
                title: args.title,
                placeholder: args.placeholder,
                lock: args.lock,
                color: args.color,
                appearance: args.appearance,
                date: args.date,
                dateFormat: args.dateFormat,
                dataBinding: args.dataBinding,
                updateFromXml: args.updateFromXml,
              });
              if (args.checked !== undefined) {
                requireMethod(control, "SetCheckBoxChecked", "设置内容控件复选状态").call(control, Boolean(args.checked));
              }
              return {
                control: control,
                kind: kind,
                contentMode: kind === "inline" ? contentMode : undefined,
                target: resolvedTarget ? resolvedTarget.target : null,
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
                case "word_replace_text": {
                  if (!args.search) throw new Error("word_replace_text.search 不能为空");
                  var replacementProperties = {
                    searchString: String(args.search),
                    replaceString: args.replace === undefined ? "" : String(args.replace),
                    matchCase: Boolean(args.matchCase),
                  };
                  var replacementRanges = requireMethod(
                    doc,
                    "Search",
                    "统计 Word 替换匹配项",
                  ).call(doc, replacementProperties.searchString, replacementProperties.matchCase) || [];
                  var replacementCount = replacementRanges.length;
                  if (replacementCount) {
                    var replacementAccepted = requireMethod(
                      doc,
                      "SearchAndReplace",
                      "批内执行 Word 文本替换",
                    ).call(doc, replacementProperties);
                    if (replacementAccepted === false) {
                      throw new Error("ONLYOFFICE 拒绝执行 SearchAndReplace");
                    }
                  }
                  changed += replacementCount;
                  results.push({
                    name: call.name,
                    replaced: replacementCount > 0,
                    replacedCount: replacementCount,
                    search: replacementProperties.searchString,
                  });
                  break;
                }

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
                        spacingBeforePt: typeof paragraph.GetSpacingBefore === "function"
                          ? Number(paragraph.GetSpacingBefore()) / 20
                          : null,
                        spacingAfterPt: typeof paragraph.GetSpacingAfter === "function"
                          ? Number(paragraph.GetSpacingAfter()) / 20
                          : null,
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
                    tableDetails.push({
                      index: tableIndex + 1,
                      rows: rows,
                      columns: columns,
                      align: typeof table.GetJc === "function" ? table.GetJc() : null,
                    });
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

                case "word_inspect_advanced": {
                  var maxAdvancedItems = Math.min(1000, Math.max(1, Math.floor(Number(args.maxItems) || 200)));
                  var advanced = { name: call.name };
                  if (args.includeProperties !== false) {
                    var core = safeCall(doc, "GetCore", [], null);
                    var coreNames = [
                      "Title", "Subject", "Creator", "Description", "Keywords", "Category",
                      "Language", "Identifier", "LastModifiedBy", "Revision", "Created", "Modified",
                    ];
                    var coreProperties = {};
                    for (var coreIndex = 0; coreIndex < coreNames.length; coreIndex += 1) {
                      var coreName = coreNames[coreIndex];
                      coreProperties[coreName.charAt(0).toLowerCase() + coreName.slice(1)] =
                        safeCall(core, "Get" + coreName, [], null);
                    }
                    advanced.properties = {
                      core: coreProperties,
                      documentInfo: safeCall(doc, "GetDocumentInfo", [], null),
                    };
                    var requestedCustomProperties = Array.isArray(args.customPropertyNames)
                      ? args.customPropertyNames
                      : [];
                    if (requestedCustomProperties.length) {
                      var customPropertiesReader = safeCall(doc, "GetCustomProperties", [], null);
                      advanced.properties.custom = {};
                      for (var customPropertyIndex = 0; customPropertyIndex < Math.min(maxAdvancedItems, requestedCustomProperties.length); customPropertyIndex += 1) {
                        var requestedCustomPropertyName = String(requestedCustomProperties[customPropertyIndex] || "");
                        if (!requestedCustomPropertyName) throw new Error("customPropertyNames 不能包含空名称");
                        advanced.properties.custom[requestedCustomPropertyName] = safeCall(
                          customPropertiesReader,
                          "Get",
                          [requestedCustomPropertyName],
                          null,
                        );
                      }
                    }
                  }
                  if (args.includeSections !== false) {
                    var advancedSections = documentSections();
                    advanced.sections = [];
                    for (var advancedSectionIndex = 0; advancedSectionIndex < Math.min(maxAdvancedItems, advancedSections.length); advancedSectionIndex += 1) {
                      var advancedSection = advancedSections[advancedSectionIndex];
                      advanced.sections.push({
                        index: advancedSectionIndex + 1,
                        type: safeCall(advancedSection, "GetType", [], null),
                        pageWidthTwips: safeCall(advancedSection, "GetPageWidth", [], null),
                        pageHeightTwips: safeCall(advancedSection, "GetPageHeight", [], null),
                        marginLeftTwips: safeCall(advancedSection, "GetPageMarginLeft", [], null),
                        marginTopTwips: safeCall(advancedSection, "GetPageMarginTop", [], null),
                        marginRightTwips: safeCall(advancedSection, "GetPageMarginRight", [], null),
                        marginBottomTwips: safeCall(advancedSection, "GetPageMarginBottom", [], null),
                        startPageNumber: safeCall(advancedSection, "GetStartPageNumber", [], null),
                        titlePage: safeCall(advancedSection, "GetTitlePage", [], null),
                      });
                    }
                  }
                  if (args.includeStyles === true) {
                    var publicStyles = safeCall(doc, "GetAllStyles", [], []) || [];
                    var allStyles = [];
                    var seenStyleNames = {};

                    function collectionValuesForInspection(collection) {
                      if (!collection || (typeof collection !== "object" && typeof collection !== "function")) {
                        return [];
                      }
                      var values = [];
                      var seenKeys = {};
                      var collectionLength = Math.max(0, Math.floor(Number(collection.length) || 0));
                      for (var collectionIndex = 0; collectionIndex < collectionLength; collectionIndex += 1) {
                        var numericKey = String(collectionIndex);
                        seenKeys[numericKey] = true;
                        if (collection[collectionIndex]) values.push(collection[collectionIndex]);
                      }
                      var ownKeys = Object.keys(collection);
                      for (var ownKeyIndex = 0; ownKeyIndex < ownKeys.length; ownKeyIndex += 1) {
                        var ownKey = ownKeys[ownKeyIndex];
                        if (seenKeys[ownKey]) continue;
                        var ownValue = collection[ownKey];
                        if (ownValue) values.push(ownValue);
                      }
                      return values;
                    }

                    function addStyleForInspection(style) {
                      if (!style) return;
                      var styleName = safeCall(style, "GetName", [], null);
                      var styleKey = styleName === null || styleName === undefined
                        ? null
                        : "$" + String(styleName);
                      if (styleKey !== null && seenStyleNames[styleKey]) return;
                      if (styleKey !== null) seenStyleNames[styleKey] = true;
                      allStyles.push(style);
                    }

                    var publicStyleValues = collectionValuesForInspection(publicStyles);
                    for (var publicStyleIndex = 0; publicStyleIndex < publicStyleValues.length; publicStyleIndex += 1) {
                      addStyleForInspection(publicStyleValues[publicStyleIndex]);
                    }

                    // ONLYOFFICE 9.4 stores newly created styles under non-index
                    // keys such as "1_30" in an Array. ApiDocument.GetAllStyles()
                    // uses Array#forEach and therefore skips those styles. Merge
                    // the internal collection when available, then resolve each
                    // name back to the public ApiStyle wrapper.
                    var internalDocument = doc && doc.Document ? doc.Document : null;
                    var internalStyleCollection = safeCall(internalDocument, "Get_Styles", [], null);
                    var internalStyles = safeCall(internalStyleCollection, "GetAllStyles", [], []) || [];
                    var internalStyleValues = collectionValuesForInspection(internalStyles);
                    for (var internalStyleIndex = 0; internalStyleIndex < internalStyleValues.length; internalStyleIndex += 1) {
                      var internalStyle = internalStyleValues[internalStyleIndex];
                      var internalStyleName = safeCall(internalStyle, "GetName", [], null);
                      var publicStyle = internalStyleName !== null
                        && internalStyleName !== undefined
                        && typeof doc.GetStyle === "function"
                        ? doc.GetStyle(String(internalStyleName))
                        : null;
                      addStyleForInspection(publicStyle || internalStyle);
                    }

                    advanced.styles = [];
                    for (var styleIndex = 0; styleIndex < Math.min(maxAdvancedItems, allStyles.length); styleIndex += 1) {
                      var advancedStyle = allStyles[styleIndex];
                      var basedOn = safeCall(advancedStyle, "GetBasedOn", [], null);
                      advanced.styles.push({
                        index: styleIndex + 1,
                        name: safeCall(advancedStyle, "GetName", [], null),
                        type: safeCall(advancedStyle, "GetType", [], null),
                        basedOn: basedOn ? safeCall(basedOn, "GetName", [], null) : null,
                      });
                    }
                  }
                  if (args.includeNumbering === true) {
                    var numberedParagraphs = safeCall(doc, "GetAllNumberedParagraphs", [], []) || [];
                    advanced.numbering = [];
                    for (var numberedIndex = 0; numberedIndex < Math.min(maxAdvancedItems, numberedParagraphs.length); numberedIndex += 1) {
                      var numberedParagraph = numberedParagraphs[numberedIndex];
                      advanced.numbering.push({
                        paragraphIndex: allParagraphs().indexOf(numberedParagraph) + 1,
                        text: paragraphText(numberedParagraph).slice(0, 500),
                        level: safeCall(numberedParagraph, "GetNumPr", [], null),
                      });
                    }
                  }
                  if (args.includeDrawings === true) {
                    var drawings = allDrawings();
                    advanced.drawings = [];
                    for (var drawingIndex = 0; drawingIndex < Math.min(maxAdvancedItems, drawings.length); drawingIndex += 1) {
                      var advancedDrawing = drawings[drawingIndex];
                      advanced.drawings.push({
                        index: drawingIndex + 1,
                        kind: drawingKind(advancedDrawing),
                        name: safeCall(advancedDrawing, "GetName", [], null),
                        widthEmu: safeCall(advancedDrawing, "GetWidth", [], null),
                        heightEmu: safeCall(advancedDrawing, "GetHeight", [], null),
                        rotation: safeCall(advancedDrawing, "GetRotation", [], null),
                        wrapping: safeCall(advancedDrawing, "GetWrappingStyle", [], null),
                      });
                    }
                  }
                  if (args.includeBookmarks === true) {
                    advanced.bookmarks = (safeCall(doc, "GetAllBookmarksNames", [], []) || []).slice(0, maxAdvancedItems);
                  }
                  if (args.includeNotes === true) {
                    var footnoteParagraphs = safeCall(doc, "GetFootnotesFirstParagraphs", [], []) || [];
                    var endnoteParagraphs = safeCall(doc, "GetEndNotesFirstParagraphs", [], []) || [];
                    advanced.footnotes = footnoteParagraphs.slice(0, maxAdvancedItems).map(function (paragraph, index) {
                      return { index: index + 1, text: paragraphText(paragraph).slice(0, 1000) };
                    });
                    advanced.endnotes = endnoteParagraphs.slice(0, maxAdvancedItems).map(function (paragraph, index) {
                      return { index: index + 1, text: paragraphText(paragraph).slice(0, 1000) };
                    });
                  }
                  if (args.includeComments === true) {
                    var advancedComments = safeCall(doc, "GetAllComments", [], []) || [];
                    advanced.comments = advancedComments.slice(0, maxAdvancedItems).map(function (comment) {
                      return {
                        id: safeCall(comment, "GetId", [], null),
                        text: safeCall(comment, "GetText", [], ""),
                        author: safeCall(comment, "GetAuthorName", [], ""),
                        solved: safeCall(comment, "IsSolved", [], false),
                      };
                    });
                  }
                  if (args.includeRevisions === true) {
                    advanced.revisions = {
                      tracking: safeCall(doc, "IsTrackRevisions", [], false),
                      report: safeCall(doc, "GetReviewReport", [], null),
                    };
                  }
                  if (args.includeContentControls === true) {
                    var controls = safeCall(doc, "GetAllContentControls", [], []) || [];
                    advanced.contentControls = controls.slice(0, maxAdvancedItems).map(function (control, index) {
                      var controlDate = safeCall(control, "GetDate", [], null);
                      var controlList = safeCall(control, "GetDropdownList", [], null);
                      var controlListItems = safeCall(controlList, "GetAllItems", [], []) || [];
                      return {
                        index: index + 1,
                        type: safeCall(control, "GetClassType", [], null),
                        tag: safeCall(control, "GetTag", [], ""),
                        title: safeCall(control, "GetAlias", [], safeCall(control, "GetTitle", [], "")),
                        lock: safeCall(control, "GetLock", [], null),
                        appearance: safeCall(control, "GetAppearance", [], null),
                        placeholder: safeCall(control, "GetPlaceholderText", [], ""),
                        checked: safeCall(control, "IsCheckBox", [], false)
                          ? safeCall(control, "IsCheckBoxChecked", [], false)
                          : null,
                        date: controlDate && typeof controlDate.toISOString === "function"
                          ? controlDate.toISOString()
                          : controlDate,
                        dataBinding: safeCall(control, "GetDataBinding", [], null),
                        listItems: controlListItems.slice(0, maxAdvancedItems).map(function (item) {
                          return {
                            display: safeCall(item, "GetText", [], ""),
                            value: safeCall(item, "GetValue", [], ""),
                          };
                        }),
                      };
                    });
                  }
                  if (args.includeCustomXml === true) {
                    var xmlManager = safeCall(doc, "GetCustomXmlParts", [], null);
                    var xmlParts = safeCall(xmlManager, "GetAll", [], []) || [];
                    advanced.customXml = xmlParts.slice(0, maxAdvancedItems).map(function (part) {
                      return {
                        id: safeCall(part, "GetId", [], null),
                        xml: String(safeCall(part, "GetXml", [], "") || "").slice(0, 20000),
                      };
                    });
                  }
                  results.push(advanced);
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

                case "word_add_image": {
                  if (!args._image || !args._image.url) {
                    commandError("INVALID_IMAGE_SOURCE", "图片资源尚未安全导入");
                  }
                  if (typeof Api.CreateImage !== "function") {
                    commandError("IMAGE_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持 Api.CreateImage");
                  }
                  var imageTarget = resolveImageTarget(args);
                  var imageSize = resolveImageSize(args, 160, 180);
                  var image = Api.CreateImage(
                    String(args._image.url),
                    Math.round(imageSize.widthMm * 36000),
                    Math.round(imageSize.heightMm * 36000)
                  );
                  if (!image) commandError("IMAGE_FETCH_FAILED", "ONLYOFFICE 无法创建图片对象");
                  var wrapping = String(args.wrapping || "inline");
                  var wrappingModes = {
                    inline: true,
                    square: true,
                    tight: true,
                    through: true,
                    topAndBottom: true,
                    behind: true,
                    inFront: true,
                  };
                  if (!wrappingModes[wrapping]) throw new Error("不支持的图片环绕方式：" + wrapping);
                  if (typeof image.SetWrappingStyle === "function") image.SetWrappingStyle(wrapping);
                  else if (wrapping !== "inline") {
                    commandError("IMAGE_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持图片环绕方式");
                  }
                  if (args.name !== undefined && typeof image.SetName === "function") image.SetName(String(args.name));
                  if (imageTarget.mode === "current") {
                    if (typeof doc.InsertContent !== "function") {
                      commandError("IMAGE_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持在光标处插入图片");
                    }
                    var imageParagraph = Api.CreateParagraph();
                    if (!imageParagraph || typeof imageParagraph.AddDrawing !== "function") {
                      commandError("IMAGE_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持在 Word 段落中插图");
                    }
                    imageParagraph.AddDrawing(image);
                    if (doc.InsertContent([imageParagraph], true) === false) {
                      throw new Error("ONLYOFFICE 拒绝在当前光标处插入图片");
                    }
                  } else {
                    if (!imageTarget.paragraph || typeof imageTarget.paragraph.AddDrawing !== "function") {
                      commandError("IMAGE_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持在目标段落中插图");
                    }
                    imageTarget.paragraph.AddDrawing(image);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    inserted: true,
                    target: imageTarget.mode,
                    paragraphIndex: imageTarget.paragraphIndex,
                    widthMm: imageSize.widthMm,
                    heightMm: imageSize.heightMm,
                    wrapping: wrapping,
                    assetId: String(args._image.assetId || ""),
                  });
                  break;
                }

                case "word_set_document_properties": {
                  var propertyCore = typeof doc.GetCore === "function" ? doc.GetCore() : null;
                  if (!propertyCore) commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持文档核心属性");
                  var propertyMap = {
                    title: "SetTitle",
                    subject: "SetSubject",
                    creator: "SetCreator",
                    description: "SetDescription",
                    keywords: "SetKeywords",
                    category: "SetCategory",
                    language: "SetLanguage",
                    identifier: "SetIdentifier",
                    lastModifiedBy: "SetLastModifiedBy",
                    revision: "SetRevision",
                    created: "SetCreated",
                    modified: "SetModified",
                  };
                  var updatedProperties = [];
                  for (var propertyName in propertyMap) {
                    if (!Object.prototype.hasOwnProperty.call(propertyMap, propertyName) || args[propertyName] === undefined) continue;
                    var setterName = propertyMap[propertyName];
                    if (typeof propertyCore[setterName] !== "function") {
                      commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持核心属性：" + propertyName);
                    }
                    var propertyValue = args[propertyName];
                    if (propertyName === "created" || propertyName === "modified") {
                      propertyValue = validDate(propertyValue, propertyName);
                    }
                    propertyCore[setterName](propertyValue);
                    updatedProperties.push(propertyName);
                  }
                  var customOperations = Array.isArray(args.custom) ? args.custom : [];
                  if (customOperations.length) {
                    var customProperties = typeof doc.GetCustomProperties === "function" ? doc.GetCustomProperties() : null;
                    if (!customProperties) commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持自定义文档属性");
                    for (var customIndex = 0; customIndex < customOperations.length; customIndex += 1) {
                      var operation = customOperations[customIndex] || {};
                      var customName = String(operation.name || "");
                      if (!customName) throw new Error("custom.name 不能为空");
                      var customValue = operation.value;
                      if (customValue === undefined) throw new Error("custom.value 不能为空");
                      if (operation.valueType === "date") {
                        customValue = validDate(customValue, "custom.value");
                      } else if (operation.valueType === "number") {
                        customValue = finiteNumber(customValue, "custom.value");
                      } else if (operation.valueType === "boolean") {
                        if (customValue !== true && customValue !== false && customValue !== "true" && customValue !== "false") {
                          throw new Error("custom.value 必须是布尔值");
                        }
                        customValue = customValue === true || customValue === "true";
                      } else if (operation.valueType === "string") {
                        customValue = String(customValue);
                      }
                      if (typeof customProperties.Add !== "function" || customProperties.Add(customName, customValue) === false) {
                        throw new Error("ONLYOFFICE 无法写入自定义属性：" + customName);
                      }
                      updatedProperties.push("custom:" + customName);
                    }
                  }
                  if (!updatedProperties.length) throw new Error("word_set_document_properties 至少需要一个属性");
                  changed += updatedProperties.length;
                  results.push({ name: call.name, updated: updatedProperties });
                  break;
                }

                case "word_manage_section": {
                  var sectionAction = String(args.action || "configure");
                  var sectionEntry;
                  if (sectionAction === "create") {
                    if (args.paragraphIndex === undefined) throw new Error("创建分节时必须提供 paragraphIndex");
                    if (typeof doc.CreateSection !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持创建分节");
                    var sectionParagraph = paragraphByIndex(args.paragraphIndex, "paragraphIndex");
                    var createdSection = doc.CreateSection(sectionParagraph.paragraph);
                    if (!createdSection) throw new Error("ONLYOFFICE 无法在指定段落创建分节");
                    sectionEntry = { section: createdSection, index: documentSections().indexOf(createdSection) };
                  } else if (sectionAction === "configure") {
                    sectionEntry = getSection(args.sectionIndex);
                  } else {
                    throw new Error("word_manage_section.action 不受支持");
                  }
                  var managedSection = sectionEntry.section;
                  if (args.type !== undefined) {
                    if (typeof managedSection.SetType !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持设置分节类型");
                    managedSection.SetType(String(args.type));
                  }
                  if (args.startPageNumber !== undefined) {
                    if (typeof managedSection.SetStartPageNumber !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持设置起始页码");
                    managedSection.SetStartPageNumber(Math.floor(finiteNumber(args.startPageNumber, "startPageNumber")));
                  }
                  if (args.titlePage !== undefined) managedSection.SetTitlePage(Boolean(args.titlePage));
                  if (args.evenAndOddHeaders !== undefined) {
                    if (typeof doc.SetEvenAndOddHdrFtr !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持奇偶页独立页眉页脚");
                    doc.SetEvenAndOddHdrFtr(Boolean(args.evenAndOddHeaders));
                  }
                  if (args.columns) {
                    var columnConfig = args.columns;
                    if (Array.isArray(columnConfig.entries) && columnConfig.entries.length) {
                      var columnWidths = [];
                      var columnSpaces = [];
                      for (var columnIndex = 0; columnIndex < columnConfig.entries.length; columnIndex += 1) {
                        columnWidths.push(mmToTwips(columnConfig.entries[columnIndex].widthMm, "columns.entries.widthMm"));
                        if (columnIndex < columnConfig.entries.length - 1) {
                          columnSpaces.push(mmToTwips(columnConfig.entries[columnIndex].spaceMm || 0, "columns.entries.spaceMm"));
                        }
                      }
                      if (typeof managedSection.SetNotEqualColumns !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持非等宽分栏");
                      managedSection.SetNotEqualColumns(columnWidths, columnSpaces);
                    } else {
                      var columnCount = Math.min(20, Math.max(1, Math.floor(Number(columnConfig.count) || 1)));
                      if (typeof managedSection.SetEqualColumns !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持等宽分栏");
                      managedSection.SetEqualColumns(columnCount, mmToTwips(columnConfig.spaceMm || 0, "columns.spaceMm"));
                    }
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: sectionAction,
                    sectionIndex: sectionEntry.index >= 0 ? sectionEntry.index + 1 : null,
                    sections: documentSections().length,
                  });
                  break;
                }

                case "word_manage_style": {
                  var styleAction = String(args.action || "create");
                  var publicStyleType = String(args.type || "paragraph");
                  var styleType = publicStyleType === "character" ? "run" : publicStyleType;
                  var managedStyle = typeof doc.GetStyle === "function" ? doc.GetStyle(String(args.name)) : null;
                  if (!managedStyle && styleAction === "update") throw new Error("Word 中不存在样式：" + args.name);
                  if (!managedStyle) {
                    if (typeof doc.CreateStyle !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持创建样式");
                    managedStyle = doc.CreateStyle(String(args.name), styleType);
                  }
                  if (!managedStyle) throw new Error("ONLYOFFICE 无法创建或取得样式：" + args.name);
                  if (args.basedOn !== undefined) {
                    var baseStyle = typeof doc.GetStyle === "function" ? doc.GetStyle(String(args.basedOn)) : null;
                    if (!baseStyle) throw new Error("Word 中不存在父样式：" + args.basedOn);
                    if (typeof managedStyle.SetBasedOn !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持样式继承");
                    managedStyle.SetBasedOn(baseStyle);
                  }
                  var styleTextPr = safeCall(managedStyle, "GetTextPr", [], null);
                  var styleParaPr = safeCall(managedStyle, "GetParaPr", [], null);
                  applyTextFormat(styleTextPr, args);
                  applyParagraphFormat(styleParaPr, args);
                  changed += 1;
                  var returnedStyleType = safeCall(managedStyle, "GetType", [], styleType);
                  results.push({
                    name: call.name,
                    action: styleAction,
                    style: safeCall(managedStyle, "GetName", [], String(args.name)),
                    type: returnedStyleType === "run" ? "character" : returnedStyleType,
                    basedOn: args.basedOn === undefined ? null : String(args.basedOn),
                  });
                  break;
                }

                case "word_set_tabs": {
                  var tabEntries = selectParagraphEntries(args, true);
                  var configuredTabs = Array.isArray(args.tabs) ? args.tabs : [];
                  var tabPositions = [];
                  var tabAlignments = [];
                  if (!args.clearAll) {
                    for (var tabIndex = 0; tabIndex < configuredTabs.length; tabIndex += 1) {
                      var tab = configuredTabs[tabIndex] || {};
                      tabPositions.push(mmToTwips(tab.positionMm, "tabs.positionMm"));
                      tabAlignments.push(String(tab.align || "left"));
                    }
                  }
                  for (var tabParagraphIndex = 0; tabParagraphIndex < tabEntries.length; tabParagraphIndex += 1) {
                    if (typeof tabEntries[tabParagraphIndex].paragraph.SetTabs !== "function") {
                      commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持自定义制表位");
                    }
                    tabEntries[tabParagraphIndex].paragraph.SetTabs(tabPositions, tabAlignments);
                  }
                  changed += tabEntries.length;
                  results.push({ name: call.name, paragraphs: tabEntries.length, tabs: tabPositions.length });
                  break;
                }

                case "word_set_numbering": {
                  var numberingEntries = selectParagraphEntries(args, true);
                  var numberingKind = String(args.kind || "numbered");
                  var customNumbering = doc.CreateNumbering(numberingKind === "bullet" ? "bullet" : "numbered");
                  var configuredLevels = Array.isArray(args.levels) ? args.levels : [];
                  for (var configuredLevelIndex = 0; configuredLevelIndex < configuredLevels.length; configuredLevelIndex += 1) {
                    var configuredLevel = configuredLevels[configuredLevelIndex] || {};
                    var levelIndex = Math.min(8, Math.max(0, Math.floor(Number(configuredLevel.level) || 0)));
                    var levelObject = customNumbering.GetLevel(levelIndex);
                    if (configuredLevel.format !== undefined || configuredLevel.text !== undefined || configuredLevel.align !== undefined) {
                      if (typeof levelObject.SetCustomType !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持自定义编号格式");
                      var defaultNumberingText = "";
                      for (var defaultLevelIndex = 0; defaultLevelIndex <= levelIndex; defaultLevelIndex += 1) {
                        defaultNumberingText += "%" + defaultLevelIndex + ".";
                      }
                      levelObject.SetCustomType(
                        String(configuredLevel.format || "decimal"),
                        String(configuredLevel.text || defaultNumberingText),
                        String(configuredLevel.align || "left"),
                      );
                    }
                    if (configuredLevel.start !== undefined && typeof levelObject.SetStart === "function") {
                      levelObject.SetStart(Math.floor(finiteNumber(configuredLevel.start, "levels.start")));
                    }
                    if (configuredLevel.restart !== undefined && typeof levelObject.SetRestart === "function") {
                      levelObject.SetRestart(Math.floor(finiteNumber(configuredLevel.restart, "levels.restart")));
                    }
                  }
                  var selectedLevel = Math.min(8, Math.max(0, Math.floor(Number(args.level) || 0)));
                  var selectedNumberingLevel = customNumbering.GetLevel(selectedLevel);
                  if (args.restartAt !== undefined && typeof selectedNumberingLevel.SetStart === "function") {
                    selectedNumberingLevel.SetStart(Math.floor(finiteNumber(args.restartAt, "restartAt")));
                  }
                  for (var numberingParagraphIndex = 0; numberingParagraphIndex < numberingEntries.length; numberingParagraphIndex += 1) {
                    numberingEntries[numberingParagraphIndex].paragraph.SetNumbering(selectedNumberingLevel);
                  }
                  changed += numberingEntries.length;
                  results.push({
                    name: call.name,
                    kind: numberingKind,
                    level: selectedLevel,
                    paragraphs: numberingEntries.length,
                    configuredLevels: configuredLevels.length,
                  });
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
                  var position = String(args.position || "after");
                  var hasBreakTarget = (
                    args.all === true
                    || args.current === true
                    || (Array.isArray(args.paragraphIndexes) && args.paragraphIndexes.length > 0)
                    || (args.search !== undefined && String(args.search) !== "")
                  );
                  if (!hasBreakTarget) {
                    commandError(
                      "INVALID_ARGUMENTS",
                      "必须使用 paragraphIndexes、search、all=true 或 current=true 指定段落",
                      { partialMutationPossible: false },
                    );
                  }
                  if (position !== "before" && position !== "after") {
                    commandError(
                      "INVALID_ARGUMENTS",
                      "word_insert_page_break.position 必须是 before 或 after",
                      { partialMutationPossible: false },
                    );
                  }
                  var breakEntries = selectParagraphEntries(args, true);
                  for (var breakIndex = 0; breakIndex < breakEntries.length; breakIndex += 1) {
                    if (position === "before") breakEntries[breakIndex].paragraph.SetPageBreakBefore(true);
                    else breakEntries[breakIndex].paragraph.AddPageBreak();
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
                  ensureVirtualTableState();
                  var tableInsertion = resolveTableInsertion(args);
                  var rows = Math.min(100, Math.max(1, Math.floor(Number(args.rows) || 1)));
                  var cols = Math.min(50, Math.max(1, Math.floor(Number(args.cols) || 1)));
                  var table = Api.CreateTable(rows, cols);
                  table.SetWidth("percent", Math.min(100, Math.max(10, Number(args.widthPercent) || 100)));
                  if (args.align !== undefined) {
                    requireMethod(table, "SetJc", "设置 Word 表格整体对齐").call(
                      table,
                      String(args.align),
                    );
                  }
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
                      setTableCellValue(cell, data[rowIndex][colIndex], args);
                    }
                  }
                  if (args.title && typeof table.SetTableTitle === "function") table.SetTableTitle(String(args.title));
                  if (args.description && typeof table.SetTableDescription === "function") table.SetTableDescription(String(args.description));
                  var tableElements = [];
                  if (args.pageBreakBefore === true) {
                    var pageBreakParagraph = Api.CreateParagraph();
                    requireMethod(
                      pageBreakParagraph,
                      "SetPageBreakBefore",
                      "创建表格前分页段落",
                    ).call(pageBreakParagraph, true);
                    tableElements.push(pageBreakParagraph);
                  }
                  tableElements.push(table);
                  var tableDocumentPosition;
                  if (tableInsertion.insertAt === "end") {
                    tableDocumentPosition = virtualTopLevelContent.length + tableElements.length - 1;
                    for (var tableElementIndex = 0; tableElementIndex < tableElements.length; tableElementIndex += 1) {
                      if (requireMethod(doc, "Push", "追加 Word 表格").call(doc, tableElements[tableElementIndex]) === false) {
                        throw new Error("ONLYOFFICE 无法追加 Word 表格");
                      }
                    }
                    Array.prototype.push.apply(virtualTopLevelContent, tableElements);
                  } else if (tableInsertion.insertAt === "current") {
                    if (requireMethod(doc, "InsertContent", "在光标处插入 Word 表格").call(
                      doc,
                      tableElements,
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法在光标处插入 Word 表格");
                    }
                    var currentTablePosition = typeof table.GetPosInParent === "function"
                      ? Number(table.GetPosInParent())
                      : NaN;
                    var currentInsertionPosition = isFinite(currentTablePosition) && currentTablePosition >= 0
                      ? Math.floor(currentTablePosition) - tableElements.length + 1
                      : virtualTopLevelContent.length;
                    Array.prototype.splice.apply(
                      virtualTopLevelContent,
                      [currentInsertionPosition, 0].concat(tableElements),
                    );
                    tableDocumentPosition = currentInsertionPosition + tableElements.length - 1;
                  } else {
                    tableDocumentPosition = tableInsertion.position + tableElements.length - 1;
                    for (var relativeElementIndex = 0; relativeElementIndex < tableElements.length; relativeElementIndex += 1) {
                      if (requireMethod(doc, "AddElement", "定点插入 Word 表格").call(
                        doc,
                        tableInsertion.position + relativeElementIndex,
                        tableElements[relativeElementIndex],
                      ) === false) {
                        throw new Error("ONLYOFFICE 无法定点插入 Word 表格");
                      }
                    }
                    Array.prototype.splice.apply(
                      virtualTopLevelContent,
                      [tableInsertion.position, 0].concat(tableElements),
                    );
                  }
                  var insertedTableIndex = tableCollectionInsertionIndex(tableDocumentPosition);
                  virtualTables.splice(insertedTableIndex, 0, table);
                  changed += rows * cols;
                  results.push({
                    name: call.name,
                    insertAt: tableInsertion.insertAt,
                    tableIndex: insertedTableIndex + 1,
                    anchor: tableInsertion.anchor,
                    pageBreakBefore: args.pageBreakBefore === true,
                    rows: rows,
                    columns: cols,
                  });
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
                    if (
                      args.rowStart === undefined
                      || args.rowEnd === undefined
                      || args.columnStart === undefined
                      || args.columnEnd === undefined
                    ) {
                      commandError(
                        "INVALID_ARGUMENTS",
                        "word_edit_table mergeCells 需要 rowStart、rowEnd、columnStart、columnEnd",
                        { partialMutationPossible: false },
                      );
                    }
                    var rowStart = Math.floor(finiteNumber(args.rowStart, "rowStart"));
                    var rowEnd = Math.floor(finiteNumber(args.rowEnd, "rowEnd"));
                    var columnStart = Math.floor(finiteNumber(args.columnStart, "columnStart"));
                    var columnEnd = Math.floor(finiteNumber(args.columnEnd, "columnEnd"));
                    if (rowStart > rowEnd || columnStart > columnEnd) {
                      commandError(
                        "INVALID_ARGUMENTS",
                        "合并区域起点不能大于终点",
                        { partialMutationPossible: false },
                      );
                    }
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

                case "word_format_table_advanced": {
                  var advancedTableEntry = getTable(args.tableIndex);
                  var advancedTable = advancedTableEntry.table;
                  if (args.repeatHeader !== undefined && args.row === undefined) {
                    commandError(
                      "INVALID_ARGUMENTS",
                      "设置重复表头时必须提供 row",
                      { partialMutationPossible: false },
                    );
                  }
                  var targetRows = [];
                  var targetCells = [];
                  if (args.row !== undefined) {
                    var targetRowCell = getCell(advancedTable, args.row, args.column || 1);
                    targetRows = [targetRowCell.row];
                    if (args.column !== undefined) {
                      targetCells = [targetRowCell.cell];
                    } else {
                      for (var rowCellIndex = 0; rowCellIndex < targetRowCell.row.GetCellsCount(); rowCellIndex += 1) {
                        targetCells.push(targetRowCell.row.GetCell(rowCellIndex));
                      }
                    }
                  } else {
                    for (var tableRowIndex = 0; tableRowIndex < advancedTable.GetRowsCount(); tableRowIndex += 1) {
                      targetRows.push(advancedTable.GetRow(tableRowIndex));
                    }
                  }
                  if (!targetCells.length && args.column !== undefined) {
                    for (var columnRowIndex = 0; columnRowIndex < advancedTable.GetRowsCount(); columnRowIndex += 1) {
                      targetCells.push(getCell(advancedTable, columnRowIndex + 1, args.column).cell);
                    }
                  }
                  if (!targetCells.length) targetCells = typeof advancedTable.GetAllCells === "function" ? advancedTable.GetAllCells() || [] : [];
                  if (args.rowHeightMm !== undefined) {
                    var rowHeightRule = String(args.rowHeightRule || "atLeast");
                    for (var heightRowIndex = 0; heightRowIndex < targetRows.length; heightRowIndex += 1) {
                      if (typeof targetRows[heightRowIndex].SetHeight !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持设置表格行高");
                      targetRows[heightRowIndex].SetHeight(rowHeightRule, mmToTwips(args.rowHeightMm, "rowHeightMm"));
                    }
                  }
                  if (args.columnWidthMm !== undefined) {
                    for (var widthCellIndex = 0; widthCellIndex < targetCells.length; widthCellIndex += 1) {
                      if (typeof targetCells[widthCellIndex].SetWidth !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持设置单元格宽度");
                      targetCells[widthCellIndex].SetWidth("twips", mmToTwips(args.columnWidthMm, "columnWidthMm"));
                    }
                  }
                  if (args.repeatHeader !== undefined) {
                    for (var headerRowIndex = 0; headerRowIndex < targetRows.length; headerRowIndex += 1) {
                      if (typeof targetRows[headerRowIndex].SetTableHeader !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持跨页重复表头");
                      targetRows[headerRowIndex].SetTableHeader(Boolean(args.repeatHeader));
                    }
                  }
                  if (args.cellMarginsMm) {
                    var margins = args.cellMarginsMm;
                    var marginNames = ["Top", "Right", "Bottom", "Left"];
                    var marginValues = [margins.top, margins.right, margins.bottom, margins.left];
                    if (args.row !== undefined || args.column !== undefined) {
                      for (var marginCellIndex = 0; marginCellIndex < targetCells.length; marginCellIndex += 1) {
                        for (var cellMarginIndex = 0; cellMarginIndex < marginNames.length; cellMarginIndex += 1) {
                          if (marginValues[cellMarginIndex] === undefined) continue;
                          var cellMarginMethod = "SetCellMargin" + marginNames[cellMarginIndex];
                          requireMethod(targetCells[marginCellIndex], cellMarginMethod, "设置单元格边距").call(
                            targetCells[marginCellIndex],
                            mmToTwips(marginValues[cellMarginIndex], "cellMarginsMm." + marginNames[cellMarginIndex].toLowerCase()),
                          );
                        }
                      }
                    } else {
                      for (var tableMarginIndex = 0; tableMarginIndex < marginNames.length; tableMarginIndex += 1) {
                        if (marginValues[tableMarginIndex] === undefined) continue;
                        var tableMarginMethod = "SetTableCellMargin" + marginNames[tableMarginIndex];
                        requireMethod(advancedTable, tableMarginMethod, "设置表格默认单元格边距").call(
                          advancedTable,
                          mmToTwips(marginValues[tableMarginIndex], "cellMarginsMm." + marginNames[tableMarginIndex].toLowerCase()),
                        );
                      }
                    }
                  }
                  if (args.borders) {
                    var tableBorderMethods = {
                      top: "SetTableBorderTop",
                      right: "SetTableBorderRight",
                      bottom: "SetTableBorderBottom",
                      left: "SetTableBorderLeft",
                      insideHorizontal: "SetTableBorderInsideH",
                      insideVertical: "SetTableBorderInsideV",
                    };
                    var cellBorderMethods = {
                      top: "SetCellBorderTop",
                      right: "SetCellBorderRight",
                      bottom: "SetCellBorderBottom",
                      left: "SetCellBorderLeft",
                    };
                    for (var borderName in tableBorderMethods) {
                      if (!Object.prototype.hasOwnProperty.call(tableBorderMethods, borderName) || !args.borders[borderName]) continue;
                      var border = args.borders[borderName];
                      var borderRgb = hexComponents(border.color || "#000000");
                      var borderArguments = [
                        String(border.style || "single"),
                        Math.max(0, Math.round(finiteNumber(border.widthPt === undefined ? 1 : border.widthPt, "borders.widthPt") * 8)),
                        Math.max(0, finiteNumber(border.spacePt || 0, "borders.spacePt")),
                        borderRgb[0],
                        borderRgb[1],
                        borderRgb[2],
                      ];
                      if (args.row !== undefined || args.column !== undefined) {
                        if (!cellBorderMethods[borderName]) {
                          throw new Error("单元格目标不支持 insideHorizontal 或 insideVertical 边框");
                        }
                        for (var borderCellIndex = 0; borderCellIndex < targetCells.length; borderCellIndex += 1) {
                          requireMethod(targetCells[borderCellIndex], cellBorderMethods[borderName], "设置单元格边框").apply(
                            targetCells[borderCellIndex],
                            borderArguments,
                          );
                        }
                      } else {
                        requireMethod(advancedTable, tableBorderMethods[borderName], "设置表格边框").apply(
                          advancedTable,
                          borderArguments,
                        );
                      }
                    }
                  }
                  if (args.wrapping !== undefined) {
                    if (typeof advancedTable.SetWrappingStyle !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持表格文字环绕");
                    advancedTable.SetWrappingStyle(String(args.wrapping));
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    tableIndex: advancedTableEntry.index + 1,
                    rows: targetRows.length,
                    cells: targetCells.length,
                  });
                  break;
                }

                case "word_add_nested_table": {
                  var parentTableEntry = getTable(args.tableIndex);
                  var parentCellEntry = getCell(parentTableEntry.table, args.row, args.column);
                  var nestedRows = Math.min(50, Math.max(1, Math.floor(finiteNumber(args.rows, "rows"))));
                  var nestedCols = Math.min(20, Math.max(1, Math.floor(finiteNumber(args.cols, "cols"))));
                  var nestedTable = Api.CreateTable(nestedRows, nestedCols);
                  if (args.widthPercent !== undefined) nestedTable.SetWidth("percent", Math.min(100, Math.max(1, finiteNumber(args.widthPercent, "widthPercent"))));
                  if (args.align !== undefined) {
                    requireMethod(nestedTable, "SetJc", "设置 Word 嵌套表格整体对齐").call(
                      nestedTable,
                      String(args.align),
                    );
                  }
                  if (args.styleName) nestedTable.SetStyle(resolveStyle(args.styleName, "table"));
                  var nestedData = Array.isArray(args.data) ? args.data : [];
                  for (var nestedRowIndex = 0; nestedRowIndex < nestedRows; nestedRowIndex += 1) {
                    for (var nestedColumnIndex = 0; nestedColumnIndex < nestedCols; nestedColumnIndex += 1) {
                      if (!nestedData[nestedRowIndex] || nestedData[nestedRowIndex][nestedColumnIndex] === undefined) continue;
                      var nestedCell = nestedTable.GetRow(nestedRowIndex).GetCell(nestedColumnIndex);
                      setTableCellValue(nestedCell, nestedData[nestedRowIndex][nestedColumnIndex], args);
                    }
                  }
                  var parentContent = parentCellEntry.cell.GetContent();
                  var insertedNested = false;
                  if (typeof parentContent.Push === "function") insertedNested = parentContent.Push(nestedTable) !== false;
                  else if (typeof parentTableEntry.table.AddElement === "function") {
                    insertedNested = parentTableEntry.table.AddElement(
                      parentCellEntry.cell,
                      safeCall(parentContent, "GetElementsCount", [], 0),
                      nestedTable,
                    ) !== false;
                  }
                  if (!insertedNested) throw new Error("ONLYOFFICE 无法插入嵌套表格");
                  changed += 1;
                  results.push({
                    name: call.name,
                    tableIndex: parentTableEntry.index + 1,
                    row: parentCellEntry.rowIndex + 1,
                    column: parentCellEntry.columnIndex + 1,
                    rows: nestedRows,
                    columns: nestedCols,
                  });
                  break;
                }

                case "word_manage_drawing": {
                  var drawingEntry = getDrawing(args);
                  var managedDrawing = drawingEntry.drawing;
                  var drawingAction = String(args.action || "");
                  if (drawingAction === "delete") {
                    if (typeof managedDrawing.Delete !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持删除该图形对象");
                    if (managedDrawing.Delete() === false) throw new Error("ONLYOFFICE 无法删除图形对象");
                  } else if (drawingAction === "update") {
                    setDrawingCommon(managedDrawing, args);
                  } else {
                    throw new Error("word_manage_drawing.action 必须是 update 或 delete");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: drawingAction,
                    drawingIndex: drawingEntry.index + 1,
                    kind: drawingKind(managedDrawing),
                  });
                  break;
                }

                case "word_add_shape": {
                  if (typeof Api.CreateShape !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持创建形状");
                  var shapeWidth = Math.round(finiteNumber(args.widthMm === undefined ? 60 : args.widthMm, "widthMm") * 36000);
                  var shapeHeight = Math.round(finiteNumber(args.heightMm === undefined ? 30 : args.heightMm, "heightMm") * 36000);
                  var shapeFill = args.fillColor === undefined
                    ? Api.CreateSolidFill(Api.Color("#D9EAF7"))
                    : Api.CreateSolidFill(Api.Color(String(args.fillColor)));
                  var shapeStroke = createStroke({
                    widthPt: args.lineWidthPt === undefined ? 1 : args.lineWidthPt,
                    color: args.lineColor || "#4472C4",
                  }, "#4472C4");
                  var shape = Api.CreateShape(String(args.shapeType), shapeWidth, shapeHeight, shapeFill, shapeStroke);
                  if (!shape) throw new Error("ONLYOFFICE 无法创建形状：" + args.shapeType);
                  if (args.text !== undefined) {
                    var shapeContent = safeCall(shape, "GetDocContent", [], safeCall(shape, "GetContent", [], null));
                    var shapeParagraph = shapeContent && typeof shapeContent.GetElement === "function" ? shapeContent.GetElement(0) : null;
                    if (!shapeParagraph) {
                      shapeParagraph = Api.CreateParagraph();
                      if (shapeContent && typeof shapeContent.Push === "function") shapeContent.Push(shapeParagraph);
                    }
                    if (!shapeParagraph) commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持形状内文本");
                    var shapeTextRun = shapeParagraph.SetText(String(args.text));
                    applyTextFormat(shapeTextRun, args);
                    applyParagraphFormat(shapeParagraph, args);
                  }
                  if (args.name !== undefined && typeof shape.SetName === "function") shape.SetName(String(args.name));
                  if (args.wrapping !== undefined && typeof shape.SetWrappingStyle === "function") shape.SetWrappingStyle(String(args.wrapping));
                  if (args.rotationDeg !== undefined && typeof shape.SetRotation === "function") shape.SetRotation(finiteNumber(args.rotationDeg, "rotationDeg"));
                  var shapeParagraphIndex = addDrawingToParagraph(shape, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    inserted: true,
                    paragraphIndex: shapeParagraphIndex,
                    kind: "shape",
                    shapeType: String(args.shapeType),
                  });
                  break;
                }

                case "word_add_chart": {
                  if (typeof Api.CreateChart !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持创建图表");
                  if (!Array.isArray(args.data) || !args.data.length) throw new Error("word_add_chart.data 不能为空");
                  var chartWidth = Math.round(finiteNumber(args.widthMm === undefined ? 120 : args.widthMm, "widthMm") * 36000);
                  var chartHeight = Math.round(finiteNumber(args.heightMm === undefined ? 75 : args.heightMm, "heightMm") * 36000);
                  var chartNumberFormats = Array.isArray(args.numberFormats)
                    ? args.numberFormats
                    : args.data.map(function () { return "General"; });
                  if (chartNumberFormats.length !== args.data.length) {
                    throw new Error("numberFormats 长度必须与 data 的系列数一致");
                  }
                  var chart = Api.CreateChart(
                    String(args.chartType || "bar"),
                    args.data,
                    Array.isArray(args.seriesNames) ? args.seriesNames : [],
                    Array.isArray(args.categories) ? args.categories : [],
                    chartWidth,
                    chartHeight,
                    Math.floor(Number(args.style) || 1),
                    chartNumberFormats,
                  );
                  if (!chart) throw new Error("ONLYOFFICE 无法创建图表");
                  if (args.title !== undefined && typeof chart.SetTitle === "function") chart.SetTitle(String(args.title), 13);
                  if (args.legendPosition !== undefined && typeof chart.SetLegendPos === "function") {
                    chart.SetLegendPos(String(args.legendPosition));
                  }
                  if (
                    args.showSeriesNames !== undefined || args.showCategories !== undefined ||
                    args.showValues !== undefined
                  ) {
                    if (typeof chart.SetShowDataLabels !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持图表数据标签");
                    chart.SetShowDataLabels(
                      Boolean(args.showSeriesNames),
                      Boolean(args.showCategories),
                      Boolean(args.showValues),
                      false,
                    );
                  }
                  if (args.wrapping !== undefined && typeof chart.SetWrappingStyle === "function") chart.SetWrappingStyle(String(args.wrapping));
                  var chartParagraphIndex = addDrawingToParagraph(chart, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    inserted: true,
                    paragraphIndex: chartParagraphIndex,
                    kind: "chart",
                    chartType: String(args.chartType || "bar"),
                    series: args.data.length,
                  });
                  break;
                }

                case "word_add_math": {
                  if (typeof doc.AddMathEquation !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持数学公式");
                  if (!args.equation) throw new Error("word_add_math.equation 不能为空");
                  var mathTarget = resolveInsertionParagraph(args);
                  if (mathTarget.paragraph && typeof mathTarget.paragraph.Select === "function") mathTarget.paragraph.Select();
                  if (doc.AddMathEquation(String(args.equation), String(args.format || "unicode")) === false) {
                    throw new Error("ONLYOFFICE 无法插入数学公式");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    inserted: true,
                    format: String(args.format || "unicode"),
                    paragraphIndex: mathTarget.index < 0 ? null : mathTarget.index + 1,
                  });
                  break;
                }

                case "word_add_ole_object": {
                  if (!args._image || !args._image.url) commandError("INVALID_IMAGE_SOURCE", "OLE 预览图尚未安全导入");
                  if (typeof Api.CreateOleObject !== "function") commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本或许可证不支持 OLE 对象");
                  var oleSize = resolveImageSize(args, 120, 90);
                  var oleObject = Api.CreateOleObject(
                    String(args._image.url),
                    Math.round(oleSize.widthMm * 36000),
                    Math.round(oleSize.heightMm * 36000),
                    String(args.data || ""),
                    String(args.applicationId || ""),
                  );
                  if (!oleObject) throw new Error("ONLYOFFICE 无法创建 OLE 对象");
                  if (args.name !== undefined && typeof oleObject.SetName === "function") oleObject.SetName(String(args.name));
                  var oleParagraphIndex = addDrawingToParagraph(oleObject, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    inserted: true,
                    paragraphIndex: oleParagraphIndex,
                    kind: "oleObject",
                    assetId: String(args._image.assetId || ""),
                  });
                  break;
                }

                case "word_manage_fields": {
                  var fieldAction = String(args.action || "");
                  if (fieldAction === "add") {
                    if (!args.instruction) throw new Error("word_manage_fields.instruction 不能为空");
                    var fieldTargets = [];
                    if (args.search !== undefined) {
                      fieldTargets = selectSearchRanges(args).selected;
                      if (!fieldTargets.length) throw new Error("未找到字段插入目标文本");
                    } else {
                      var fieldParagraph = resolveInsertionParagraph(args);
                      var fieldPlaceholder = requireMethod(
                        fieldParagraph.paragraph,
                        "AddText",
                        "创建字段占位范围",
                      ).call(fieldParagraph.paragraph, "\u200B");
                      var fieldRange = requireMethod(fieldPlaceholder, "GetRange", "取得字段插入范围").call(
                        fieldPlaceholder,
                      );
                      if (!fieldRange) throw new Error("ONLYOFFICE 无法取得字段插入范围");
                      fieldTargets = [fieldRange];
                    }
                    for (var fieldIndex = 0; fieldIndex < fieldTargets.length; fieldIndex += 1) {
                      if (requireMethod(fieldTargets[fieldIndex], "AddField", "插入动态字段").call(
                        fieldTargets[fieldIndex],
                        String(args.instruction),
                      ) === false) {
                        throw new Error("ONLYOFFICE 无法插入字段：" + args.instruction);
                      }
                    }
                    changed += fieldTargets.length;
                    results.push({
                      name: call.name,
                      action: fieldAction,
                      instruction: String(args.instruction),
                      insertedFields: fieldTargets.length,
                    });
                  } else if (fieldAction === "updateAll") {
                    if (requireMethod(doc, "UpdateAllFields", "更新全部字段").call(doc) === false) {
                      throw new Error("ONLYOFFICE 无法更新全部字段");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: fieldAction, updated: true });
                  } else if (fieldAction === "clearForms") {
                    if (requireMethod(doc, "ClearAllFields", "清除全部表单字段").call(doc) === false) {
                      throw new Error("ONLYOFFICE 无法清除全部表单字段");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: fieldAction, cleared: true });
                  } else {
                    throw new Error("word_manage_fields.action 不受支持");
                  }
                  break;
                }

                case "word_manage_long_document": {
                  var longAction = String(args.action || "");
                  if (args.paragraphIndex !== undefined) {
                    var longTarget = paragraphByIndex(args.paragraphIndex, "paragraphIndex");
                    if (typeof longTarget.paragraph.Select === "function") longTarget.paragraph.Select();
                  }
                  if (longAction === "addToc") {
                    var tocProperties = {
                      ShowPageNums: args.showPageNumbers !== false,
                      RightAlgn: args.rightAlignPageNumbers !== false,
                      LeaderType: String(args.leaderType || "dot"),
                      FormatAsLinks: args.formatAsLinks !== false,
                      BuildFrom: { OutlineLvls: Math.min(9, Math.max(1, Math.floor(Number(args.outlineLevels) || 9))) },
                      TocStyle: "standard",
                    };
                    if (requireMethod(doc, "AddTableOfContents", "插入自动目录").call(doc, tocProperties) === false) {
                      throw new Error("ONLYOFFICE 无法插入自动目录");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: longAction, properties: tocProperties });
                  } else if (longAction === "updateToc") {
                    if (requireMethod(doc, "UpdateAllTOC", "更新自动目录").call(doc) === false) {
                      throw new Error("ONLYOFFICE 无法更新自动目录");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: longAction, updated: true });
                  } else if (longAction === "addCaption") {
                    if (args.paragraphIndex === undefined) throw new Error("addCaption 必须提供 paragraphIndex");
                    if (!args.label) throw new Error("addCaption.label 不能为空");
                    var captionTarget = paragraphByIndex(args.paragraphIndex, "paragraphIndex").paragraph;
                    if (requireMethod(captionTarget, "AddCaption", "插入题注").call(
                      captionTarget,
                      String(args.text || ""),
                      String(args.label),
                      Boolean(args.excludeLabel),
                      String(args.numberFormat || "Arabic"),
                      Boolean(args.before),
                      args.headingLevel === undefined ? undefined : Math.floor(finiteNumber(args.headingLevel, "headingLevel")),
                      String(args.separator || "hyphen"),
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法插入题注");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: longAction, label: String(args.label) });
                  } else if (longAction === "addTableOfFigures") {
                    if (!args.label) throw new Error("addTableOfFigures.label 不能为空");
                    var figureProperties = {
                      ShowPageNums: args.showPageNumbers !== false,
                      RightAlgn: args.rightAlignPageNumbers !== false,
                      LeaderType: String(args.leaderType || "dot"),
                      FormatAsLinks: args.formatAsLinks !== false,
                      BuildFrom: String(args.label),
                      LabelNumber: args.excludeLabel !== true,
                      TofStyle: "standard",
                    };
                    if (requireMethod(doc, "AddTableOfFigures", "插入图表目录").call(doc, figureProperties, false) === false) {
                      throw new Error("ONLYOFFICE 无法插入图表目录");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: longAction, properties: figureProperties });
                  } else if (longAction === "updateTableOfFigures") {
                    if (requireMethod(doc, "UpdateAllTOF", "更新图表目录").call(doc) === false) {
                      throw new Error("ONLYOFFICE 无法更新图表目录");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: longAction, updated: true });
                  } else if (longAction === "addCrossReference") {
                    var referenceKind = String(args.referenceKind || "caption");
                    var referenceTarget = resolveInsertionParagraph(args).paragraph;
                    var referenceSucceeded = false;
                    var referencedIndex = null;
                    if (referenceKind === "caption") {
                      if (!args.label) throw new Error("题注交叉引用的 label 不能为空");
                      var captionTargetIndexValue = args.targetIndex === undefined ? args.captionIndex : args.targetIndex;
                      if (captionTargetIndexValue === undefined) throw new Error("题注交叉引用的 targetIndex 不能为空");
                      var captionParagraphs = requireMethod(doc, "GetAllCaptionParagraphs", "读取题注").call(
                        doc,
                        String(args.label),
                      ) || [];
                      referencedIndex = Math.floor(finiteNumber(captionTargetIndexValue, "targetIndex")) - 1;
                      if (referencedIndex < 0 || referencedIndex >= captionParagraphs.length) throw new Error("targetIndex 超出题注范围");
                      referenceSucceeded = requireMethod(referenceTarget, "AddCaptionCrossRef", "插入题注交叉引用").call(
                        referenceTarget,
                        String(args.label),
                        String(args.referenceType || "entireCaption"),
                        captionParagraphs[referencedIndex],
                        args.hyperlink !== false,
                        Boolean(args.aboveBelow),
                      );
                    } else if (referenceKind === "bookmark") {
                      if (!args.bookmarkName) throw new Error("书签交叉引用的 bookmarkName 不能为空");
                      referenceSucceeded = requireMethod(referenceTarget, "AddBookmarkCrossRef", "插入书签交叉引用").call(
                        referenceTarget,
                        String(args.referenceType || "text"),
                        String(args.bookmarkName),
                        args.hyperlink !== false,
                        Boolean(args.aboveBelow),
                        String(args.numberSeparator || ""),
                      );
                    } else {
                      if (args.targetIndex === undefined) throw new Error(referenceKind + " 交叉引用的 targetIndex 不能为空");
                      referencedIndex = Math.floor(finiteNumber(args.targetIndex, "targetIndex")) - 1;
                      var referenceParagraphs;
                      var referenceMethod;
                      var defaultReferenceType;
                      if (referenceKind === "heading") {
                        referenceParagraphs = requireMethod(doc, "GetAllHeadingParagraphs", "读取标题段落").call(doc) || [];
                        referenceMethod = "AddHeadingCrossRef";
                        defaultReferenceType = "text";
                      } else if (referenceKind === "numbered") {
                        referenceParagraphs = requireMethod(doc, "GetAllNumberedParagraphs", "读取编号段落").call(doc) || [];
                        referenceMethod = "AddNumberedCrossRef";
                        defaultReferenceType = "paraNum";
                      } else if (referenceKind === "footnote") {
                        referenceParagraphs = requireMethod(doc, "GetFootnotesFirstParagraphs", "读取脚注").call(doc) || [];
                        referenceMethod = "AddFootnoteCrossRef";
                        defaultReferenceType = "formFootnoteNum";
                      } else if (referenceKind === "endnote") {
                        referenceParagraphs = requireMethod(doc, "GetEndNotesFirstParagraphs", "读取尾注").call(doc) || [];
                        referenceMethod = "AddEndnoteCrossRef";
                        defaultReferenceType = "formEndnoteNum";
                      } else {
                        throw new Error("referenceKind 不受支持");
                      }
                      if (referencedIndex < 0 || referencedIndex >= referenceParagraphs.length) {
                        throw new Error("targetIndex 超出 " + referenceKind + " 目标范围");
                      }
                      var referenceArguments = [
                        String(args.referenceType || defaultReferenceType),
                        referenceParagraphs[referencedIndex],
                        args.hyperlink !== false,
                        Boolean(args.aboveBelow),
                      ];
                      if (referenceKind === "numbered") referenceArguments.push(String(args.numberSeparator || ""));
                      referenceSucceeded = requireMethod(referenceTarget, referenceMethod, "插入交叉引用").apply(
                        referenceTarget,
                        referenceArguments,
                      );
                    }
                    if (referenceSucceeded === false) {
                      throw new Error("ONLYOFFICE 无法插入交叉引用");
                    }
                    changed += 1;
                    results.push({
                      name: call.name,
                      action: longAction,
                      referenceKind: referenceKind,
                      targetIndex: referencedIndex === null ? null : referencedIndex + 1,
                      bookmarkName: referenceKind === "bookmark" ? String(args.bookmarkName) : null,
                    });
                  } else if (longAction === "addFootnote" || longAction === "addEndnote") {
                    if (!args.noteText) throw new Error(longAction + ".noteText 不能为空");
                    var noteTarget = resolveInsertionParagraph(args);
                    if (noteTarget.paragraph && typeof noteTarget.paragraph.Select === "function") noteTarget.paragraph.Select();
                    var noteMethod = longAction === "addFootnote" ? "AddFootnote" : "AddEndnote";
                    var note = requireMethod(doc, noteMethod, longAction === "addFootnote" ? "插入脚注" : "插入尾注").call(doc);
                    if (!note) throw new Error("ONLYOFFICE 无法插入脚注或尾注");
                    if (typeof note.AddText === "function") note.AddText(String(args.noteText));
                    else {
                      var noteParagraph = typeof note.GetElement === "function" ? note.GetElement(0) : null;
                      if (!noteParagraph || typeof noteParagraph.AddText !== "function") {
                        commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本无法写入脚注或尾注内容");
                      }
                      noteParagraph.AddText(String(args.noteText));
                    }
                    changed += 1;
                    results.push({ name: call.name, action: longAction, inserted: true });
                  } else if (longAction === "deleteBookmark") {
                    if (!args.bookmarkName) throw new Error("deleteBookmark.bookmarkName 不能为空");
                    if (requireMethod(doc, "DeleteBookmark", "删除书签").call(doc, String(args.bookmarkName)) === false) {
                      throw new Error("ONLYOFFICE 未找到或无法删除书签：" + args.bookmarkName);
                    }
                    changed += 1;
                    results.push({
                      name: call.name,
                      action: longAction,
                      bookmarkName: String(args.bookmarkName),
                      deleted: true,
                    });
                  } else {
                    throw new Error("word_manage_long_document.action 不受支持");
                  }
                  break;
                }

                case "word_manage_comments": {
                  var commentAction = String(args.action || "");
                  if (commentAction === "removeAll") {
                    var removableComments = requireMethod(doc, "GetAllComments", "读取批注").call(doc) || [];
                    for (var removeCommentIndex = removableComments.length - 1; removeCommentIndex >= 0; removeCommentIndex -= 1) {
                      requireMethod(removableComments[removeCommentIndex], "Delete", "删除批注").call(
                        removableComments[removeCommentIndex],
                      );
                    }
                    changed += removableComments.length;
                    results.push({ name: call.name, action: commentAction, removed: removableComments.length });
                    break;
                  }
                  if (!args.commentId) throw new Error("commentId 不能为空");
                  var managedComment = requireMethod(doc, "GetCommentById", "按 ID 读取批注").call(
                    doc,
                    String(args.commentId),
                  );
                  if (!managedComment) throw new Error("未找到批注：" + args.commentId);
                  if (commentAction === "reply") {
                    if (!args.text) throw new Error("reply.text 不能为空");
                    requireMethod(managedComment, "AddReply", "回复批注").call(
                      managedComment,
                      String(args.text),
                      args.author === undefined ? undefined : String(args.author),
                      args.userId === undefined ? undefined : String(args.userId),
                    );
                  } else if (commentAction === "edit") {
                    if (args.text === undefined) throw new Error("edit.text 不能为空");
                    requireMethod(managedComment, "SetText", "编辑批注").call(managedComment, String(args.text));
                  } else if (commentAction === "remove") {
                    requireMethod(managedComment, "Delete", "删除批注").call(managedComment);
                  } else if (commentAction === "resolve" || commentAction === "reopen") {
                    requireMethod(managedComment, "SetSolved", "设置批注解决状态").call(
                      managedComment,
                      commentAction === "resolve",
                    );
                  } else {
                    throw new Error("word_manage_comments.action 不受支持");
                  }
                  changed += 1;
                  results.push({ name: call.name, action: commentAction, commentId: String(args.commentId) });
                  break;
                }

                case "word_manage_revisions": {
                  var revisionAction = String(args.action || "");
                  if (revisionAction === "start" || revisionAction === "stop") {
                    var revisionEnabled = revisionAction === "start";
                    var revisionSetter = typeof doc.SetAssistantTrackRevisions === "function"
                      ? "SetAssistantTrackRevisions"
                      : "SetTrackRevisions";
                    if (requireMethod(doc, revisionSetter, "设置修订模式").call(doc, revisionEnabled) === false) {
                      throw new Error("ONLYOFFICE 无法设置修订模式");
                    }
                  } else if (revisionAction === "acceptAll") {
                    if (requireMethod(doc, "AcceptAllRevisionChanges", "接受全部修订").call(doc) === false) {
                      throw new Error("ONLYOFFICE 无法接受全部修订");
                    }
                  } else if (revisionAction === "rejectAll") {
                    if (requireMethod(doc, "RejectAllRevisionChanges", "拒绝全部修订").call(doc) === false) {
                      throw new Error("ONLYOFFICE 无法拒绝全部修订");
                    }
                  } else {
                    throw new Error("word_manage_revisions.action 不受支持");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: revisionAction,
                    tracking: typeof doc.IsTrackRevisions === "function" ? doc.IsTrackRevisions() : undefined,
                  });
                  break;
                }

                case "word_manage_content_control": {
                  var controlAction = String(args.action || "");
                  if (controlAction === "add") {
                    var createdControl = createContentControl(args);
                    changed += 1;
                    results.push({
                      name: call.name,
                      action: controlAction,
                      kind: createdControl.kind,
                      id: safeCall(createdControl.control, "GetId", [], null),
                      tag: safeCall(createdControl.control, "GetTag", [], args.tag || ""),
                      contentMode: createdControl.contentMode,
                      target: createdControl.target,
                    });
                    break;
                  }
                  var controlEntry = contentControlByArgs(args);
                  var managedControl = controlEntry.control;
                  if (controlAction === "update") {
                    if (!hasDefined(args, [
                      "tag", "title", "text", "placeholder", "items", "selectedValue", "checked",
                      "lock", "color", "appearance", "date", "dateFormat", "dataBinding", "updateFromXml",
                    ])) {
                      throw new Error("update 至少需要一个内容控件属性");
                    }
                    applyContentControlProperties(managedControl, args);
                    if (Array.isArray(args.items)) {
                      var dropdownList = requireMethod(managedControl, "GetDropdownList", "更新内容控件列表").call(
                        managedControl,
                      );
                      requireMethod(dropdownList, "Clear", "清空内容控件列表").call(dropdownList);
                      for (var listItemIndex = 0; listItemIndex < args.items.length; listItemIndex += 1) {
                        var listItem = args.items[listItemIndex] || {};
                        var display = String(listItem.display || "");
                        requireMethod(dropdownList, "Add", "添加内容控件列表项").call(
                          dropdownList,
                          display,
                          listItem.value === undefined ? display : String(listItem.value),
                        );
                      }
                    }
                    if (args.selectedValue !== undefined) {
                      requireMethod(managedControl, "SelectListItem", "选择内容控件列表项").call(
                        managedControl,
                        String(args.selectedValue),
                      );
                    }
                    if (args.checked !== undefined) {
                      requireMethod(managedControl, "SetCheckBoxChecked", "设置内容控件复选状态").call(
                        managedControl,
                        Boolean(args.checked),
                      );
                    }
                  } else if (controlAction === "remove") {
                    if (requireMethod(managedControl, "Delete", "删除内容控件").call(managedControl, true) === false) {
                      throw new Error("ONLYOFFICE 无法删除内容控件");
                    }
                  } else if (controlAction === "clear") {
                    if (requireMethod(managedControl, "RemoveAllElements", "清空内容控件").call(managedControl) === false) {
                      throw new Error("ONLYOFFICE 无法清空内容控件");
                    }
                  } else if (controlAction === "check") {
                    if (args.checked === undefined) throw new Error("check.checked 不能为空");
                    if (requireMethod(managedControl, "SetCheckBoxChecked", "设置内容控件复选状态").call(
                      managedControl,
                      Boolean(args.checked),
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法设置内容控件复选状态");
                    }
                  } else {
                    throw new Error("word_manage_content_control.action 不受支持");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: controlAction,
                    index: controlEntry.index < 0 ? null : controlEntry.index + 1,
                    id: safeCall(managedControl, "GetId", [], null),
                    tag: safeCall(managedControl, "GetTag", [], ""),
                  });
                  break;
                }

                case "word_manage_custom_xml": {
                  var xmlAction = String(args.action || "");
                  var xmlManager = requireMethod(doc, "GetCustomXmlParts", "访问自定义 XML").call(doc);
                  if (!xmlManager) commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持自定义 XML");
                  if (xmlAction === "add") {
                    if (!args.xml) throw new Error("add.xml 不能为空");
                    var addedXmlPart = requireMethod(xmlManager, "Add", "新增自定义 XML").call(
                      xmlManager,
                      String(args.xml),
                    );
                    if (!addedXmlPart) throw new Error("ONLYOFFICE 无法新增自定义 XML");
                    changed += 1;
                    results.push({
                      name: call.name,
                      action: xmlAction,
                      partId: safeCall(addedXmlPart, "GetId", [], null),
                      xml: safeCall(addedXmlPart, "GetXml", [], String(args.xml)),
                    });
                    break;
                  }
                  if (!args.partId) throw new Error("partId 不能为空");
                  var xmlPart = requireMethod(xmlManager, "GetById", "按 ID 读取自定义 XML").call(
                    xmlManager,
                    String(args.partId),
                  );
                  if (!xmlPart) throw new Error("未找到自定义 XML partId：" + args.partId);
                  var xmlResultPart = xmlPart;
                  if (xmlAction === "replace") {
                    if (!args.xml) throw new Error("replace.xml 不能为空");
                    if (requireMethod(xmlPart, "Delete", "替换自定义 XML").call(xmlPart) === false) {
                      throw new Error("ONLYOFFICE 无法删除旧的自定义 XML");
                    }
                    xmlResultPart = requireMethod(xmlManager, "Add", "替换自定义 XML").call(xmlManager, String(args.xml));
                    if (!xmlResultPart) throw new Error("ONLYOFFICE 无法写入新的自定义 XML");
                  } else if (xmlAction === "remove") {
                    if (requireMethod(xmlPart, "Delete", "删除自定义 XML").call(xmlPart) === false) {
                      throw new Error("ONLYOFFICE 无法删除自定义 XML");
                    }
                    xmlResultPart = null;
                  } else if (xmlAction === "insertElement") {
                    if (!args.xpath || !args.xml) throw new Error("insertElement 需要 xpath 和 xml");
                    if (requireMethod(xmlPart, "InsertElement", "插入 XML 元素").call(
                      xmlPart,
                      String(args.xpath),
                      String(args.xml),
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法插入 XML 元素");
                    }
                  } else if (xmlAction === "updateElement") {
                    if (!args.xpath || !args.xml) throw new Error("updateElement 需要 xpath 和 xml");
                    if (requireMethod(xmlPart, "UpdateElement", "更新 XML 元素").call(
                      xmlPart,
                      String(args.xpath),
                      String(args.xml),
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法更新 XML 元素");
                    }
                  } else if (xmlAction === "deleteElement") {
                    if (!args.xpath) throw new Error("deleteElement.xpath 不能为空");
                    if (requireMethod(xmlPart, "DeleteElement", "删除 XML 元素").call(
                      xmlPart,
                      String(args.xpath),
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法删除 XML 元素");
                    }
                  } else if (xmlAction === "insertAttribute" || xmlAction === "updateAttribute") {
                    if (!args.xpath || !args.name || args.value === undefined) {
                      throw new Error(xmlAction + " 需要 xpath、name 和 value");
                    }
                    var attributeMethod = xmlAction === "insertAttribute" ? "InsertAttribute" : "UpdateAttribute";
                    if (requireMethod(xmlPart, attributeMethod, "写入 XML 属性").call(
                      xmlPart,
                      String(args.xpath),
                      String(args.name),
                      String(args.value),
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法写入 XML 属性");
                    }
                  } else if (xmlAction === "deleteAttribute") {
                    if (!args.xpath || !args.name) throw new Error("deleteAttribute 需要 xpath 和 name");
                    if (requireMethod(xmlPart, "DeleteAttribute", "删除 XML 属性").call(
                      xmlPart,
                      String(args.xpath),
                      String(args.name),
                    ) === false) {
                      throw new Error("ONLYOFFICE 无法删除 XML 属性");
                    }
                  } else {
                    throw new Error("word_manage_custom_xml.action 不受支持");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: xmlAction,
                    previousPartId: String(args.partId),
                    partId: xmlResultPart ? safeCall(xmlResultPart, "GetId", [], String(args.partId)) : null,
                    xml: xmlResultPart ? safeCall(xmlResultPart, "GetXml", [], null) : null,
                  });
                  break;
                }

                case "word_set_watermark": {
                  var watermarkAction = String(args.action || "setText");
                  if (watermarkAction === "remove") {
                    if (requireMethod(doc, "RemoveWatermark", "删除水印").call(doc) === false) {
                      throw new Error("ONLYOFFICE 无法删除水印");
                    }
                    changed += 1;
                    results.push({ name: call.name, action: watermarkAction, removed: true });
                    break;
                  }
                  var watermarkSettings = requireMethod(doc, "GetWatermarkSettings", "读取水印设置").call(doc);
                  if (!watermarkSettings) commandError("WORD_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持水印设置");
                  if (watermarkAction === "setText") {
                    if (!args.text) throw new Error("setText.text 不能为空");
                    requireMethod(watermarkSettings, "SetType", "设置文字水印").call(watermarkSettings, "text");
                    requireMethod(watermarkSettings, "SetText", "设置文字水印").call(watermarkSettings, String(args.text));
                    var watermarkTextPr = requireMethod(watermarkSettings, "GetTextPr", "读取水印文本格式").call(
                      watermarkSettings,
                    );
                    applyTextFormat(watermarkTextPr, args);
                    requireMethod(watermarkSettings, "SetTextPr", "设置水印文本格式").call(
                      watermarkSettings,
                      watermarkTextPr,
                    );
                  } else if (watermarkAction === "setImage") {
                    if (!args._image || !args._image.url) commandError("INVALID_IMAGE_SOURCE", "水印图片尚未安全导入");
                    requireMethod(watermarkSettings, "SetType", "设置图片水印").call(watermarkSettings, "image");
                    requireMethod(watermarkSettings, "SetImageURL", "设置图片水印").call(
                      watermarkSettings,
                      String(args._image.url),
                    );
                    var watermarkScale = args.scale === undefined ? 1 : finiteNumber(args.scale, "scale");
                    if (!(watermarkScale > 0)) throw new Error("scale 必须大于 0");
                    var watermarkWidth = Math.round(Number(args._image.widthPx) * 9525 * watermarkScale);
                    var watermarkHeight = Math.round(Number(args._image.heightPx) * 9525 * watermarkScale);
                    requireMethod(watermarkSettings, "SetImageSize", "设置图片水印尺寸").call(
                      watermarkSettings,
                      watermarkWidth,
                      watermarkHeight,
                    );
                  } else {
                    throw new Error("word_set_watermark.action 不受支持");
                  }
                  if (args.opacity !== undefined) {
                    requireMethod(watermarkSettings, "SetOpacity", "设置水印透明度").call(
                      watermarkSettings,
                      Math.round(finiteNumber(args.opacity, "opacity") * 2.55),
                    );
                  }
                  if (args.diagonal !== undefined) {
                    requireMethod(watermarkSettings, "SetDirection", "设置水印方向").call(
                      watermarkSettings,
                      args.diagonal ? "clockwise45" : "horizontal",
                    );
                  }
                  if (!requireMethod(doc, "SetWatermarkSettings", "应用水印设置").call(doc, watermarkSettings)) {
                    throw new Error("ONLYOFFICE 无法应用水印设置");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: watermarkAction,
                    type: watermarkAction === "setText" ? "text" : "image",
                    assetId: args._image ? String(args._image.assetId || "") : undefined,
                  });
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
                  if (action === "set" && !hasDefined(args, paragraphFormatKeys.concat(["text", "pageNumber", "pagesCount", "fields"]))) {
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
                    var headerFooterFields = Array.isArray(args.fields) ? args.fields : [];
                    for (var headerFooterFieldIndex = 0; headerFooterFieldIndex < headerFooterFields.length; headerFooterFieldIndex += 1) {
                      var headerFooterInstruction = String(headerFooterFields[headerFooterFieldIndex] || "");
                      if (!headerFooterInstruction) throw new Error("fields 不能包含空字段指令");
                      var headerFooterPlaceholder = requireMethod(
                        paragraph,
                        "AddText",
                        "创建页眉页脚字段占位范围",
                      ).call(paragraph, "\u200B");
                      var headerFooterRange = requireMethod(
                        headerFooterPlaceholder,
                        "GetRange",
                        "取得页眉页脚字段范围",
                      ).call(headerFooterPlaceholder);
                      if (requireMethod(headerFooterRange, "AddField", "插入页眉页脚动态字段").call(
                        headerFooterRange,
                        headerFooterInstruction,
                      ) === false) {
                        throw new Error("ONLYOFFICE 无法插入页眉页脚字段：" + headerFooterInstruction);
                      }
                    }
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
                  var documentText = String(args.text);
                  var documentLines = documentText.replace(/\r\n?/g, "\n").split("\n");
                  if (typeof doc.RemoveAllElements !== "function" || doc.RemoveAllElements() === false) {
                    throw new Error("ONLYOFFICE 无法清空 Word 正文");
                  }
                  var firstDocumentParagraph = typeof doc.GetElement === "function" ? doc.GetElement(0) : null;
                  if (!firstDocumentParagraph || typeof firstDocumentParagraph.SetText !== "function") {
                    throw new Error("ONLYOFFICE 无法取得清空后的首段落");
                  }
                  firstDocumentParagraph.SetText(documentLines[0]);
                  for (var documentLineIndex = 1; documentLineIndex < documentLines.length; documentLineIndex += 1) {
                    var documentParagraph = Api.CreateParagraph();
                    if (documentLines[documentLineIndex]) documentParagraph.AddText(documentLines[documentLineIndex]);
                    if (doc.Push(documentParagraph) === false) {
                      throw new Error("ONLYOFFICE 无法追加 Word 段落");
                    }
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    replacedDocument: true,
                    characters: documentText.length,
                    paragraphs: documentLines.length,
                  });
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
            var failedCallIndex = typeof callIndex === "number" && callIndex >= 0
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
            if (
              error
              && error.details
              && typeof error.details.partialMutationPossible === "boolean"
            ) {
              currentMutationPossible = error.details.partialMutationPossible;
            }
            var details = {
              phase: "word-command",
              completedToolCalls: failedCallIndex === null ? 0 : failedCallIndex,
              partialMutationPossible: priorMutationPossible || currentMutationPossible,
            };
            if (failedCall && failedCall.name) details.tool = String(failedCall.name);
            if (failedCallIndex !== null) details.toolCallIndex = failedCallIndex;
            return JSON.stringify({
              ok: false,
              code: error && error.code ? error.code : undefined,
              error: error && error.message ? error.message : String(error),
              details: details,
            });
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
    const originalCalls = Array.isArray(toolCalls) ? toolCalls : [];
    const unsafeUnitErrors = collectUnsafeLegacyPointUnits(originalCalls);
    const normalizedInput = normalizeWordToolCalls(originalCalls);
    const aggregate = {
      ok: true,
      editorType: "word",
      changed: 0,
      needsSave: false,
      results: [],
    };
    if (normalizedInput.argumentNormalizations.length) {
      aggregate.argumentNormalizations = normalizedInput.argumentNormalizations;
    }
    const calls = normalizedInput.toolCalls;
    requireStaticValidation(calls, unsafeUnitErrors);
    const hasReplacement = calls.some(function (call) {
      return call && call.name === "word_replace_text";
    });
    const useNativeSearchAndReplace = hasReplacement
      ? await detectNativeSearchAndReplace()
      : false;
    let callCommandBatch = [];
    let callCommandBatchStartIndex = 0;

    function addToolContext(error, call, toolCallIndex, completedToolCalls, priorMutationPossible) {
      const contextualError = error && typeof error === "object"
        ? error
        : new Error(String(error));
      const existingDetails = contextualError.details && typeof contextualError.details === "object"
        ? contextualError.details
        : {};
      contextualError.details = {
        phase: existingDetails.phase || "word-command",
        tool: call && call.name ? String(call.name) : undefined,
        toolCallIndex,
        completedToolCalls,
        partialMutationPossible: Boolean(
          existingDetails.partialMutationPossible
          || priorMutationPossible
        ),
      };
      return contextualError;
    }

    async function flushCallCommandBatch() {
      if (!callCommandBatch.length) return;
      const batch = callCommandBatch;
      const batchStartIndex = callCommandBatchStartIndex;
      callCommandBatch = [];
      try {
        mergeExecutionResult(aggregate, await executeCallCommand(batch));
      } catch (error) {
        const details = error && error.details && typeof error.details === "object"
          ? error.details
          : {};
        const relativeIndex = Number.isInteger(details.toolCallIndex)
          ? details.toolCallIndex
          : 0;
        const completedInBatch = Number.isInteger(details.completedToolCalls)
          ? details.completedToolCalls
          : relativeIndex;
        const absoluteIndex = batchStartIndex + relativeIndex;
        throw addToolContext(
          error,
          calls[absoluteIndex],
          absoluteIndex,
          batchStartIndex + completedInBatch,
          aggregate.changed > 0,
        );
      }
    }

    for (let toolCallIndex = 0; toolCallIndex < calls.length; toolCallIndex += 1) {
      const call = calls[toolCallIndex];
      try {
        if (call && call.name === "word_replace_text" && !useNativeSearchAndReplace) {
          await flushCallCommandBatch();
          mergeExecutionResult(aggregate, await executeSearchAndReplace(call));
        } else if (call && call.name === "word_set_protection") {
          await flushCallCommandBatch();
          mergeExecutionResult(aggregate, await executeEditingRestrictions(call));
        } else if (call && (call.name === "word_inspect_macros" || call.name === "word_set_macros")) {
          await flushCallCommandBatch();
          mergeExecutionResult(aggregate, await executeMacroTool(call));
        } else {
          if (!callCommandBatch.length) callCommandBatchStartIndex = toolCallIndex;
          callCommandBatch.push(call);
        }
      } catch (error) {
        if (
          error
          && error.details
          && Number.isInteger(error.details.toolCallIndex)
          && error.details.toolCallIndex < toolCallIndex
        ) {
          throw error;
        }
        const currentMutationPossible = Boolean(
          call
          && (
            call.name === "word_replace_text"
            || call.name === "word_set_protection"
            || call.name === "word_set_macros"
          )
        );
        throw addToolContext(
          error,
          call,
          toolCallIndex,
          toolCallIndex,
          aggregate.changed > 0 || currentMutationPossible,
        );
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
