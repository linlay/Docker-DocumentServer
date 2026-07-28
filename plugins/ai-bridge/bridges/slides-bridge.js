(function () {
  "use strict";

  window.AICopilotBridges = window.AICopilotBridges || {};

  function parseResult(rawResult) {
    const value = typeof rawResult === "string" ? JSON.parse(rawResult || "{}") : rawResult;
    if (!value || !value.ok) {
      const error = new Error((value && value.error) || "Slides Bridge 执行失败");
      if (value && value.code) error.code = value.code;
      if (value && value.details && typeof value.details === "object") error.details = { ...value.details };
      throw error;
    }
    return value;
  }

  const SLIDE_TABLE_CELL_FIELD_TYPES = {
    text: "string",
    fontSize: "number",
    fontFamily: "string",
    bold: "boolean",
    italic: "boolean",
    underline: "boolean",
    color: "color",
    align: "string",
    firstLineIndentMm: "number",
    leftIndentMm: "number",
    rightIndentMm: "number",
    spacingBeforePt: "number",
    spacingAfterPt: "number",
    lineSpacing: "number",
    lineRule: "string",
    level: "number",
    listType: "string",
    bulletSymbol: "string",
    numberingType: "string",
    startAt: "number",
    fill: "object",
    backgroundColor: "color",
    verticalAlign: "string",
    border: "object",
  };
  const SLIDE_TABLE_CELL_ENUMS = {
    align: ["left", "center", "right", "both"],
    lineRule: ["auto", "exact", "atLeast"],
    listType: ["none", "bullet", "number"],
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

  function isHexColor(value) {
    return typeof value === "string" && /^#?[0-9A-Fa-f]{6}$/.test(value);
  }

  function collectSlideFillErrors(index, name, path, fill, add) {
    if (!isPlainObject(fill)) {
      add(index, name, path, "fill must be an object", "type");
      return;
    }
    const type = fill.raw !== undefined ? "raw" : String(fill.type || "");
    const fieldsByType = {
      none: ["type"],
      solid: ["type", "color"],
      linearGradient: ["type", "stops", "angleDeg"],
      radialGradient: ["type", "stops"],
      pattern: ["type", "pattern", "backgroundColor", "foregroundColor"],
      raw: ["type", "raw"],
    };
    const requiredByType = {
      none: ["type"],
      solid: ["type", "color"],
      linearGradient: ["type", "stops"],
      radialGradient: ["type", "stops"],
      pattern: ["type", "pattern", "backgroundColor", "foregroundColor"],
      raw: ["raw"],
    };
    const allowedFields = fieldsByType[type];
    if (!allowedFields) {
      add(index, name, path + ".type", "fill type is not supported", "enum");
      return;
    }
    for (const field of Object.keys(fill)) {
      if (allowedFields.indexOf(field) === -1) {
        add(index, name, path + "." + field, field + " is not allowed", "additionalProperties");
      }
    }
    for (const field of requiredByType[type]) {
      if (!Object.prototype.hasOwnProperty.call(fill, field)) {
        add(index, name, path + "." + field, field + " is required", "required");
      }
    }
    for (const field of ["color", "backgroundColor", "foregroundColor"]) {
      if (Object.prototype.hasOwnProperty.call(fill, field) && !isHexColor(fill[field])) {
        add(index, name, path + "." + field, field + " must be a six-digit hex color", "pattern");
      }
    }
    if (Object.prototype.hasOwnProperty.call(fill, "pattern") && typeof fill.pattern !== "string") {
      add(index, name, path + ".pattern", "pattern must be a string", "type");
    }
    if (
      Object.prototype.hasOwnProperty.call(fill, "angleDeg")
      && (typeof fill.angleDeg !== "number" || !Number.isFinite(fill.angleDeg))
    ) {
      add(index, name, path + ".angleDeg", "angleDeg must be a finite number", "type");
    }
    if (type !== "linearGradient" && type !== "radialGradient") return;
    if (!Array.isArray(fill.stops)) {
      if (Object.prototype.hasOwnProperty.call(fill, "stops")) {
        add(index, name, path + ".stops", "stops must be an array", "type");
      }
      return;
    }
    if (fill.stops.length < 2 || fill.stops.length > 16) {
      add(index, name, path + ".stops", "stops must contain 2 to 16 items", "range");
    }
    for (let stopIndex = 0; stopIndex < fill.stops.length; stopIndex += 1) {
      const stop = fill.stops[stopIndex];
      const stopPath = path + ".stops[" + stopIndex + "]";
      if (!isPlainObject(stop)) {
        add(index, name, stopPath, "gradient stop must be an object", "type");
        continue;
      }
      for (const field of Object.keys(stop)) {
        if (field !== "position" && field !== "color") {
          add(index, name, stopPath + "." + field, field + " is not allowed", "additionalProperties");
        }
      }
      if (!Object.prototype.hasOwnProperty.call(stop, "position")) {
        add(index, name, stopPath + ".position", "position is required", "required");
      } else if (
        typeof stop.position !== "number"
        || !Number.isFinite(stop.position)
        || stop.position < 0
        || stop.position > 100
      ) {
        add(index, name, stopPath + ".position", "position must be between 0 and 100", "range");
      }
      if (!Object.prototype.hasOwnProperty.call(stop, "color")) {
        add(index, name, stopPath + ".color", "color is required", "required");
      } else if (!isHexColor(stop.color)) {
        add(index, name, stopPath + ".color", "color must be a six-digit hex color", "pattern");
      }
    }
  }

  function collectSlideBorderErrors(index, name, path, border, add) {
    if (!isPlainObject(border)) {
      add(index, name, path, "border must be an object", "type");
      return;
    }
    const allowedFields = ["widthMm", "color", "fill", "sides"];
    for (const field of Object.keys(border)) {
      if (allowedFields.indexOf(field) === -1) {
        add(index, name, path + "." + field, field + " is not allowed", "additionalProperties");
      }
    }
    if (
      Object.prototype.hasOwnProperty.call(border, "widthMm")
      && (
        typeof border.widthMm !== "number"
        || !Number.isFinite(border.widthMm)
        || border.widthMm < 0
        || border.widthMm > 20
      )
    ) {
      add(index, name, path + ".widthMm", "widthMm must be between 0 and 20", "range");
    }
    if (Object.prototype.hasOwnProperty.call(border, "color") && !isHexColor(border.color)) {
      add(index, name, path + ".color", "color must be a six-digit hex color", "pattern");
    }
    if (Object.prototype.hasOwnProperty.call(border, "fill")) {
      collectSlideFillErrors(index, name, path + ".fill", border.fill, add);
    }
    if (Object.prototype.hasOwnProperty.call(border, "sides")) {
      if (!Array.isArray(border.sides)) {
        add(index, name, path + ".sides", "sides must be an array", "type");
      } else {
        const supportedSides = ["top", "right", "bottom", "left"];
        const seenSides = {};
        for (let sideIndex = 0; sideIndex < border.sides.length; sideIndex += 1) {
          const side = border.sides[sideIndex];
          if (supportedSides.indexOf(side) === -1) {
            add(index, name, path + ".sides[" + sideIndex + "]", "side is not supported", "enum");
          } else if (seenSides[side]) {
            add(index, name, path + ".sides[" + sideIndex + "]", "side must not be duplicated", "uniqueItems");
          }
          seenSides[side] = true;
        }
      }
    }
  }

  function collectSlideTableDataErrors(index, name, data, add) {
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
          const expected = SLIDE_TABLE_CELL_FIELD_TYPES[field];
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
          } else if (expected === "object" && !isPlainObject(value)) {
            add(index, name, cellPath + "." + field, field + " must be an object", "type");
          } else if (
            expected === "color"
            && (typeof value !== "string" || !/^#?[0-9A-Fa-f]{6}$/.test(value))
          ) {
            add(index, name, cellPath + "." + field, field + " must be a six-digit hex color", "pattern");
          }
          if (
            SLIDE_TABLE_CELL_ENUMS[field]
            && typeof value === "string"
            && SLIDE_TABLE_CELL_ENUMS[field].indexOf(value) === -1
          ) {
            add(index, name, cellPath + "." + field, field + " has an unsupported value", "enum");
          }
          if (
            (field === "fontSize" || field === "lineSpacing")
            && typeof value === "number"
            && value <= 0
          ) {
            add(index, name, cellPath + "." + field, field + " must be greater than 0", "range");
          } else if (
            (field === "spacingBeforePt" || field === "spacingAfterPt")
            && typeof value === "number"
            && value < 0
          ) {
            add(index, name, cellPath + "." + field, field + " must be at least 0", "range");
          } else if (
            field === "level"
            && typeof value === "number"
            && (!Number.isInteger(value) || value < 0 || value > 8)
          ) {
            add(index, name, cellPath + "." + field, field + " must be an integer from 0 to 8", "range");
          } else if (
            field === "startAt"
            && typeof value === "number"
            && (!Number.isInteger(value) || value < 1)
          ) {
            add(index, name, cellPath + "." + field, field + " must be an integer of at least 1", "range");
          } else if (
            field === "bulletSymbol"
            && typeof value === "string"
            && (value.length < 1 || value.length > 8)
          ) {
            add(index, name, cellPath + "." + field, field + " must contain 1 to 8 characters", "range");
          }
          if (field === "fill" && isPlainObject(value)) {
            collectSlideFillErrors(index, name, cellPath + ".fill", value, add);
          } else if (field === "border" && isPlainObject(value)) {
            collectSlideBorderErrors(index, name, cellPath + ".border", value, add);
          }
        }
      }
    }
  }

  function requireStaticValidation(toolCalls) {
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
      if (String(call.name || "") !== "slides_add_table") continue;
      const args = (call && (call.arguments || call.args)) || {};
      if (!args || typeof args !== "object" || Array.isArray(args)) continue;
      collectSlideTableDataErrors(index, call.name, args.data, add);
    }
    if (!errors.length) return;
    const error = new Error("Slides 工具参数校验失败");
    error.code = "INVALID_TOOL_ARGUMENTS";
    error.details = {
      validationErrors: errors,
      completedToolCalls: 0,
      partialMutationPossible: false,
    };
    throw error;
  }

  function executeOfficeCommands(toolCalls) {
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
            var TWIPS_PER_MM = 1440 / 25.4;
            var mutatingNames = {
              slides_replace_text: true,
              slides_scale_font: true,
              slides_format_text: true,
              slides_format_selection: true,
              slides_add_slide: true,
              slides_duplicate_slide: true,
              slides_delete_slide: true,
              slides_move_slide: true,
              slides_set_visibility: true,
              slides_set_size: true,
              slides_apply_layout: true,
              slides_set_show_settings: true,
              slides_apply_theme: true,
              slides_set_theme: true,
              slides_create_layout: true,
              slides_add_template_shape: true,
              slides_manage_template_object: true,
              slides_set_template_background: true,
              slides_set_text_content: true,
              slides_format_paragraphs: true,
              slides_update_object: true,
              slides_set_hyperlink: true,
              slides_set_notes: true,
              slides_add_comment: true,
              slides_set_transition: true,
              slides_add_table: true,
              slides_set_table_cell: true,
              slides_edit_table: true,
              slides_format_table: true,
              slides_align_objects: true,
              slides_group_objects: true,
              slides_reorder_object: true,
              slides_add_connector: true,
              slides_add_freeform: true,
              slides_add_word_art: true,
              slides_add_math: true,
              slides_add_ole_object: true,
              slides_add_image_shape: true,
              slides_manage_animation: true,
              slides_manage_comment: true,
              slides_add_textbox: true,
              slides_add_image: true,
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

            function commandError(code, message) {
              var error = new Error(message);
              error.code = code;
              throw error;
            }

            function resolveImageSize(args, maxWidthMm, maxHeightMm) {
              var metadata = args._image;
              if (!metadata || !isFinite(Number(metadata.widthPx)) || !isFinite(Number(metadata.heightPx))) {
                commandError("INVALID_IMAGE_SOURCE", "内部图片资源缺少有效尺寸");
              }
              var ratio = Number(metadata.widthPx) / Number(metadata.heightPx);
              if (!isFinite(ratio) || ratio <= 0) commandError("INVALID_IMAGE_SOURCE", "内部图片宽高比无效");
              var hasWidth = hasOwn(args, "widthMm");
              var hasHeight = hasOwn(args, "heightMm");
              var width = hasWidth ? Number(args.widthMm) : null;
              var height = hasHeight ? Number(args.heightMm) : null;
              if ((hasWidth && (!isFinite(width) || width <= 0)) || (hasHeight && (!isFinite(height) || height <= 0))) {
                throw new Error("图片宽高必须是大于 0 的有限数字");
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

            function mmToTwips(value) {
              return Math.round(asFinite(value, 0) * TWIPS_PER_MM);
            }

            function pointsToTwips(value) {
              return Math.round(asFinite(value, 0) * 20);
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

            function masterCount() {
              if (typeof presentation.GetMastersCount === "function") return presentation.GetMastersCount();
              if (typeof presentation.GetAllSlideMasters === "function") return presentation.GetAllSlideMasters().length;
              return 0;
            }

            function getMaster(oneBasedIndex) {
              var count = masterCount();
              var index = Number(oneBasedIndex);
              if (!Number.isInteger(index) || index < 1 || index > count) {
                throw new Error("幻灯片母版序号超出范围，当前共 " + count + " 个母版");
              }
              return presentation.GetMaster(index - 1);
            }

            function layoutCount(master) {
              if (master && typeof master.GetLayoutsCount === "function") return master.GetLayoutsCount();
              if (master && typeof master.GetAllLayouts === "function") return master.GetAllLayouts().length;
              return 0;
            }

            function getLayout(master, oneBasedIndex) {
              var count = layoutCount(master);
              var index = Number(oneBasedIndex);
              if (!Number.isInteger(index) || index < 1 || index > count) {
                throw new Error("幻灯片版式序号超出范围，当前母版共 " + count + " 个版式");
              }
              return master.GetLayout(index - 1);
            }

            function getSlide(oneBasedIndex) {
              var index = Number(oneBasedIndex);
              if (!Number.isInteger(index) || index < 1 || index > slideCount()) {
                throw new Error("幻灯片序号无效：页码超出范围，当前共 " + slideCount() + " 页");
              }
              return presentation.GetSlideByIndex(index - 1);
            }

            function slideShapes(slide) {
              return typeof slide.GetAllShapes === "function" ? slide.GetAllShapes() : [];
            }

            function slideCharts(slide) {
              return typeof slide.GetAllCharts === "function" ? slide.GetAllCharts() : [];
            }

            function slideTables(slide) {
              return typeof slide.GetAllTables === "function" ? slide.GetAllTables() : [];
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

            function describeLayout(layout, masterIndex, layoutIndex, includeRaw, includeObjects) {
              if (!layout) return null;
              var layoutDrawings = safeCall(layout, "GetAllDrawings") || [];
              var result = {
                masterIndex: masterIndex,
                layoutIndex: layoutIndex,
                name: safeCall(layout, "GetName"),
                layoutType: safeCall(layout, "GetLayoutType"),
                objects: typeof layout.GetAllDrawings === "function" ? layoutDrawings.length : null,
                drawings: includeObjects ? layoutDrawings.map(function (drawing, index) {
                  return describeDrawing(drawing, index, includeRaw);
                }) : undefined,
              };
              if (includeRaw) result.raw = serialized(layout);
              return result;
            }

            function locateLayout(layout, includeRaw) {
              if (!layout) return null;
              for (var masterIndex = 0; masterIndex < masterCount(); masterIndex += 1) {
                var master = presentation.GetMaster(masterIndex);
                for (var layoutIndex = 0; layoutIndex < layoutCount(master); layoutIndex += 1) {
                  var candidate = master.GetLayout(layoutIndex);
                  if (candidate === layout) {
                    return describeLayout(candidate, masterIndex + 1, layoutIndex + 1, includeRaw);
                  }
                }
              }
              return describeLayout(layout, null, null, includeRaw);
            }

            function describeTheme(theme, includeRaw) {
              if (!theme) return null;
              var colorScheme = safeCall(theme, "GetColorScheme");
              var fontScheme = safeCall(theme, "GetFontScheme");
              var formatScheme = safeCall(theme, "GetFormatScheme");
              return {
                colorScheme: serialized(colorScheme),
                fontScheme: serialized(fontScheme),
                formatScheme: serialized(formatScheme),
                raw: includeRaw ? serialized(theme) : undefined,
              };
            }

            function getTemplateContainer(args) {
              var masterIndex = Number(args.masterIndex || 1);
              var master = getMaster(masterIndex);
              var scope = String(args.scope || "master");
              if (scope === "master") {
                return { scope: scope, master: master, container: master, masterIndex: masterIndex, layoutIndex: null };
              }
              if (scope === "layout") {
                var layoutIndex = Number(args.layoutIndex);
                return {
                  scope: scope,
                  master: master,
                  container: getLayout(master, layoutIndex),
                  masterIndex: masterIndex,
                  layoutIndex: layoutIndex,
                };
              }
              throw new Error("模板作用域必须是 master 或 layout");
            }

            function resolveTemplateDrawing(args) {
              var target = getTemplateContainer(args);
              var drawings = safeCall(target.container, "GetAllDrawings") || [];
              for (var index = 0; index < drawings.length; index += 1) {
                if (!drawingMatches(drawings[index], args, index)) continue;
                target.drawing = drawings[index];
                target.index = index;
                return target;
              }
              throw new Error("找不到母版或版式对象；请提供有效的 objectId、name 或对象序号");
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

            function getParagraphsFromShape(shape) {
              var content = safeCall(shape, "GetContent") || safeCall(shape, "GetDocContent");
              return content && typeof content.GetAllParagraphs === "function" ? content.GetAllParagraphs() : [];
            }

            function applyRunFormat(run, format) {
              if (format.fontSize !== undefined) run.SetFontSize(Math.max(2, Math.round(Number(format.fontSize) * 2)));
              if (format.fontFamily) run.SetFontFamily(String(format.fontFamily));
              if (format.bold !== undefined) run.SetBold(Boolean(format.bold));
              if (format.italic !== undefined) run.SetItalic(Boolean(format.italic));
              if (format.underline !== undefined) run.SetUnderline(Boolean(format.underline));
              if (format.color) run.SetColor(Api.Color(String(format.color)));
            }

            function paragraphProperties(paragraph) {
              var paraPr = safeCall(paragraph, "GetParaPr");
              return paraPr || paragraph;
            }

            function setParagraphBullet(paragraph, spec) {
              if (!hasOwn(spec, "listType")) return;
              var target = paragraphProperties(paragraph);
              if (!target || typeof target.SetBullet !== "function") {
                throw new Error("当前 ONLYOFFICE 版本不支持设置项目符号或编号");
              }
              var listType = String(spec.listType || "none");
              if (listType === "none") {
                target.SetBullet(null);
                return;
              }
              var bullet;
              if (listType === "bullet") {
                if (typeof Api.CreateBullet !== "function") throw new Error("当前 ONLYOFFICE 版本不支持项目符号");
                bullet = Api.CreateBullet(String(spec.bulletSymbol || "•"));
              } else if (listType === "number") {
                if (typeof Api.CreateNumbering !== "function") throw new Error("当前 ONLYOFFICE 版本不支持编号列表");
                bullet = Api.CreateNumbering(
                  String(spec.numberingType || "ArabicPeriod"),
                  Math.max(1, Math.round(asFinite(spec.startAt, 1)))
                );
              } else {
                throw new Error("不支持的列表类型：" + listType);
              }
              target.SetBullet(bullet);
            }

            function applyParagraphFormat(paragraph, spec) {
              var target = paragraphProperties(paragraph);
              if (hasOwn(spec, "align") && typeof target.SetJc === "function") target.SetJc(String(spec.align));
              if (hasOwn(spec, "firstLineIndentMm") && typeof target.SetIndFirstLine === "function") {
                target.SetIndFirstLine(mmToTwips(spec.firstLineIndentMm));
              }
              if (hasOwn(spec, "leftIndentMm") && typeof target.SetIndLeft === "function") {
                target.SetIndLeft(mmToTwips(spec.leftIndentMm));
              }
              if (hasOwn(spec, "rightIndentMm") && typeof target.SetIndRight === "function") {
                target.SetIndRight(mmToTwips(spec.rightIndentMm));
              }
              if (hasOwn(spec, "spacingBeforePt") && typeof target.SetSpacingBefore === "function") {
                target.SetSpacingBefore(pointsToTwips(spec.spacingBeforePt), false);
              }
              if (hasOwn(spec, "spacingAfterPt") && typeof target.SetSpacingAfter === "function") {
                target.SetSpacingAfter(pointsToTwips(spec.spacingAfterPt), false);
              }
              if (hasOwn(spec, "lineSpacing") && typeof target.SetSpacingLine === "function") {
                var lineRule = String(spec.lineRule || "auto");
                var lineValue = lineRule === "auto"
                  ? Math.round(asFinite(spec.lineSpacing, 1) * 240)
                  : pointsToTwips(spec.lineSpacing);
                target.SetSpacingLine(lineValue, lineRule);
              }
              if (hasOwn(spec, "level") && typeof target.SetOutlineLvl === "function") {
                target.SetOutlineLvl(clamp(Math.round(asFinite(spec.level, 0)), 0, 8));
              }
              setParagraphBullet(paragraph, spec);
              var runCount = typeof paragraph.GetElementsCount === "function" ? paragraph.GetElementsCount() : 0;
              for (var runIndex = 0; runIndex < runCount; runIndex += 1) {
                var run = paragraph.GetElement(runIndex);
                if (run && safeCall(run, "GetClassType") === "run") applyRunFormat(run, spec);
              }
            }

            function replaceShapeParagraphs(shape, paragraphSpecs) {
              if (!Array.isArray(paragraphSpecs) || !paragraphSpecs.length) {
                throw new Error("paragraphs 必须是非空数组");
              }
              var content = safeCall(shape, "GetContent") || safeCall(shape, "GetDocContent");
              if (!content || typeof content.Push !== "function") throw new Error("目标图形不支持文本内容");
              if (typeof content.RemoveAllElements === "function") content.RemoveAllElements();
              for (var index = 0; index < paragraphSpecs.length; index += 1) {
                var spec = paragraphSpecs[index] || {};
                var paragraph = Api.CreateParagraph();
                var run = paragraph.AddText(String(spec.text || ""));
                applyRunFormat(run, spec);
                applyParagraphFormat(paragraph, spec);
                content.Push(paragraph);
              }
            }

            function describeTransition(transition) {
              if (!transition) return null;
              return {
                effect: safeCall(transition, "GetEntryEffect"),
                speed: safeCall(transition, "GetSpeed"),
                durationMs: safeCall(transition, "GetDuration"),
                advanceOnClick: safeCall(transition, "GetAdvanceOnClick"),
                advanceOnTime: safeCall(transition, "GetAdvanceOnTime"),
                advanceTimeMs: safeCall(transition, "GetAdvanceTime"),
              };
            }

            function describeAnimationEffect(effect, index) {
              var shape = safeCall(effect, "GetShape");
              return {
                effectIndex: index,
                effectType: safeCall(effect, "GetEffectType"),
                trigger: safeCall(effect, "GetTriggerType"),
                durationMs: safeCall(effect, "GetDuration"),
                delayMs: safeCall(effect, "GetDelay"),
                repeatCount: safeCall(effect, "GetRepeatCount"),
                objectId: safeCall(shape, "GetInternalId"),
                objectName: safeCall(shape, "GetName"),
                objectKind: drawingKind(shape),
              };
            }

            function describeComment(comment, index) {
              var replies = [];
              var repliesCount = Math.max(0, Number(safeCall(comment, "GetRepliesCount")) || 0);
              for (var replyIndex = 0; replyIndex < repliesCount; replyIndex += 1) {
                var reply = safeCall(comment, "GetReply", replyIndex);
                if (reply) replies.push(describeComment(reply, replyIndex));
              }
              return {
                commentIndex: index,
                commentId: safeCall(comment, "GetId"),
                text: safeCall(comment, "GetText"),
                author: safeCall(comment, "GetAuthorName") || safeCall(comment, "GetAutorName"),
                userId: safeCall(comment, "GetUserId"),
                position: safeCall(comment, "GetPosition"),
                time: safeCall(comment, "GetTime"),
                timeUtc: safeCall(comment, "GetTimeUTC"),
                solved: safeCall(comment, "IsSolved"),
                replies: replies,
              };
            }

            function resolveComment(args) {
              var comments = safeCall(presentation, "GetAllComments") || [];
              for (var commentIndex = 0; commentIndex < comments.length; commentIndex += 1) {
                if (hasOwn(args, "commentIndex") && Number(args.commentIndex) !== commentIndex) continue;
                if (hasOwn(args, "commentId") && String(safeCall(comments[commentIndex], "GetId")) !== String(args.commentId)) continue;
                if (hasOwn(args, "commentIndex") || hasOwn(args, "commentId")) {
                  return { comment: comments[commentIndex], index: commentIndex };
                }
              }
              throw new Error("找不到目标批注；请提供有效的 commentId 或 commentIndex");
            }

            function createWordArtTextPr(args) {
              if (typeof Api.CreateTextPr !== "function") throw new Error("当前 ONLYOFFICE 版本不支持艺术字文本属性");
              var textPr = Api.CreateTextPr();
              if (hasOwn(args, "fontSize") && typeof textPr.SetFontSize === "function") {
                textPr.SetFontSize(Math.max(2, Math.round(asFinite(args.fontSize, 36) * 2)));
              }
              if (hasOwn(args, "fontFamily") && typeof textPr.SetFontFamily === "function") textPr.SetFontFamily(String(args.fontFamily));
              if (hasOwn(args, "bold") && typeof textPr.SetBold === "function") textPr.SetBold(Boolean(args.bold));
              if (hasOwn(args, "italic") && typeof textPr.SetItalic === "function") textPr.SetItalic(Boolean(args.italic));
              if (hasOwn(args, "underline") && typeof textPr.SetUnderline === "function") textPr.SetUnderline(Boolean(args.underline));
              if (hasOwn(args, "caps") && typeof textPr.SetCaps === "function") textPr.SetCaps(Boolean(args.caps));
              if ((hasOwn(args, "fontColor") || hasOwn(args, "color")) && typeof textPr.SetColor === "function") {
                textPr.SetColor(apiColor(hasOwn(args, "fontColor") ? args.fontColor : args.color));
              }
              return textPr;
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
              if (kind === "table") {
                var tableInfo = describeTable(drawing);
                info.rows = tableInfo.rows;
                info.columns = tableInfo.columns;
                info.data = tableInfo.data;
                info.columnWidthsMm = tableInfo.columnWidthsMm;
              }
              if (kind === "oleObject") {
                info.applicationId = safeCall(drawing, "GetApplicationId");
                var oleData = safeCall(drawing, "GetData");
                info.dataLength = typeof oleData === "string" ? oleData.length : null;
                if (includeRaw) info.data = oleData;
              }
              if (includeRaw) info.raw = serialized(drawing);
              return info;
            }

            function tableDimensions(table) {
              var rows = 0;
              var columns = 0;
              for (var rowIndex = 0; rowIndex < 1000; rowIndex += 1) {
                var row = safeCall(table, "GetRow", rowIndex);
                if (!row) break;
                rows += 1;
                var count = safeCall(row, "GetCellsCount");
                if (typeof count === "number") columns = Math.max(columns, count);
                else {
                  var detected = 0;
                  for (var columnIndex = 0; columnIndex < 1000; columnIndex += 1) {
                    if (!safeCall(row, "GetCell", columnIndex)) break;
                    detected += 1;
                  }
                  columns = Math.max(columns, detected);
                }
              }
              return { rows: rows, columns: columns };
            }

            function getTableCell(table, oneBasedRow, oneBasedColumn) {
              var rowIndex = Number(oneBasedRow);
              var columnIndex = Number(oneBasedColumn);
              if (!Number.isInteger(rowIndex) || rowIndex < 1 || !Number.isInteger(columnIndex) || columnIndex < 1) {
                throw new Error("表格行列序号必须是从 1 开始的整数");
              }
              var row = safeCall(table, "GetRow", rowIndex - 1);
              var cell = row ? safeCall(row, "GetCell", columnIndex - 1) : null;
              if (!cell) throw new Error("表格单元格超出范围");
              return cell;
            }

            function describeTable(table) {
              var dimensions = tableDimensions(table);
              var data = [];
              for (var rowIndex = 0; rowIndex < dimensions.rows; rowIndex += 1) {
                var row = table.GetRow(rowIndex);
                var values = [];
                var cells = safeCall(row, "GetCellsCount");
                var cellCount = typeof cells === "number" ? cells : dimensions.columns;
                for (var columnIndex = 0; columnIndex < cellCount; columnIndex += 1) {
                  var cell = safeCall(row, "GetCell", columnIndex);
                  values.push(cell ? String(safeCall(cell, "GetText") || "") : null);
                }
                data.push(values);
              }
              var columnWidths = [];
              for (var widthIndex = 0; widthIndex < dimensions.columns; widthIndex += 1) {
                columnWidths.push(emuToMm(safeCall(table, "GetColumnWidth", widthIndex)));
              }
              return {
                rows: dimensions.rows,
                columns: dimensions.columns,
                data: data,
                columnWidthsMm: columnWidths,
              };
            }

            function applyTableCellFormat(cell, args) {
              if (hasOwn(args, "text")) {
                if (typeof cell.SetText === "function") cell.SetText(String(args.text));
                else {
                  var cellContent = safeCall(cell, "GetContent");
                  if (!cellContent || typeof cellContent.Push !== "function") throw new Error("目标单元格不支持文本");
                  if (typeof cellContent.RemoveAllElements === "function") cellContent.RemoveAllElements();
                  var cellParagraph = Api.CreateParagraph();
                  cellParagraph.AddText(String(args.text));
                  cellContent.Push(cellParagraph);
                }
              }
              if (hasOwn(args, "fill") || hasOwn(args, "backgroundColor")) {
                if (typeof cell.SetShd !== "function") throw new Error("当前 ONLYOFFICE 版本不支持单元格填充");
                cell.SetShd(createFill(args.fill, args.backgroundColor));
              }
              if (hasOwn(args, "verticalAlign") && typeof cell.SetVerticalAlign === "function") {
                cell.SetVerticalAlign(String(args.verticalAlign));
              }
              var cellTextContent = safeCall(cell, "GetContent");
              var cellParagraphs = cellTextContent && typeof cellTextContent.GetAllParagraphs === "function"
                ? cellTextContent.GetAllParagraphs()
                : [];
              for (var paragraphIndex = 0; paragraphIndex < cellParagraphs.length; paragraphIndex += 1) {
                applyParagraphFormat(cellParagraphs[paragraphIndex], args);
              }
              if (isObject(args.border)) {
                var borderFill = createFill(args.border.fill, args.border.color);
                var borderWidth = Math.max(0, asFinite(args.border.widthMm, 0.25));
                var sides = Array.isArray(args.border.sides)
                  ? args.border.sides.map(String)
                  : ["top", "right", "bottom", "left"];
                var borderMethods = {
                  top: "SetCellBorderTop",
                  right: "SetCellBorderRight",
                  bottom: "SetCellBorderBottom",
                  left: "SetCellBorderLeft",
                };
                for (var sideIndex = 0; sideIndex < sides.length; sideIndex += 1) {
                  var borderMethod = borderMethods[sides[sideIndex]];
                  if (borderMethod && typeof cell[borderMethod] === "function") {
                    cell[borderMethod](borderWidth, borderFill);
                  }
                }
              }
            }

            function tableCellSpec(value) {
              if (value && typeof value === "object" && !Array.isArray(value)) return value;
              return { text: value === null || value === undefined ? "" : String(value) };
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

            function resolveDrawingTargets(slideNumber, targets) {
              if (!Array.isArray(targets) || !targets.length) throw new Error("targets 必须是非空对象选择器数组");
              var resolved = [];
              for (var index = 0; index < targets.length; index += 1) {
                var selector = targets[index] || {};
                var targetArgs = { slide: slideNumber };
                if (hasOwn(selector, "objectId")) targetArgs.objectId = selector.objectId;
                if (hasOwn(selector, "objectIndex")) targetArgs.objectIndex = selector.objectIndex;
                if (hasOwn(selector, "name")) targetArgs.name = selector.name;
                var item = resolveDrawing(targetArgs);
                if (resolved.some(function (existing) { return existing.drawing === item.drawing; })) {
                  throw new Error("targets 包含重复对象");
                }
                resolved.push(item);
              }
              return resolved;
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
                    var inspectedNotesPage = safeCall(inspectedSlide, "GetNotesPage");
                    slides.push({
                      slide: slideIndex + 1,
                      text: text.slice(0, remaining),
                      shapes: shapes.length,
                      charts: slideCharts(inspectedSlide).length,
                      objects: slideDrawings(inspectedSlide).length,
                      visible: safeCall(inspectedSlide, "GetVisible"),
                      layout: locateLayout(safeCall(inspectedSlide, "GetLayout"), false),
                      notes: inspectedNotesPage ? safeCall(inspectedNotesPage, "GetBodyShapeText") : null,
                      transition: describeTransition(safeCall(inspectedSlide, "GetSlideShowTransition")),
                    });
                    totalChars += Math.min(text.length, remaining);
                    if (totalChars >= maxChars) break;
                  }
                  results.push({
                    name: call.name,
                    slideCount: slideCount(),
                    currentSlide: presentation.GetCurSlideIndex() + 1,
                    widthMm: emuToMm(safeCall(presentation, "GetWidth")),
                    heightMm: emuToMm(safeCall(presentation, "GetHeight")),
                    masterCount: masterCount(),
                    loopUntilStopped: safeCall(presentation, "GetLoopUntilStopped"),
                    commentCount: (safeCall(presentation, "GetAllComments") || []).length,
                    slides: slides,
                    truncated: totalChars >= maxChars,
                  });
                  break;
                }

                case "slides_inspect_layouts": {
                  var inspectMasterStart = args.masterIndex ? Number(args.masterIndex) - 1 : 0;
                  var inspectMasterEnd = args.masterIndex ? inspectMasterStart + 1 : masterCount();
                  if (inspectMasterStart < 0 || inspectMasterEnd > masterCount()) {
                    throw new Error("幻灯片母版序号超出范围");
                  }
                  var inspectedMasters = [];
                  for (var inspectedMasterIndex = inspectMasterStart; inspectedMasterIndex < inspectMasterEnd; inspectedMasterIndex += 1) {
                    var inspectedMaster = presentation.GetMaster(inspectedMasterIndex);
                    var inspectedLayouts = [];
                    for (var inspectedLayoutIndex = 0; inspectedLayoutIndex < layoutCount(inspectedMaster); inspectedLayoutIndex += 1) {
                      inspectedLayouts.push(describeLayout(
                        inspectedMaster.GetLayout(inspectedLayoutIndex),
                        inspectedMasterIndex + 1,
                        inspectedLayoutIndex + 1,
                        Boolean(args.includeRaw),
                        Boolean(args.includeObjects)
                      ));
                    }
                    var inspectedMasterDrawings = safeCall(inspectedMaster, "GetAllDrawings") || [];
                    inspectedMasters.push({
                      masterIndex: inspectedMasterIndex + 1,
                      objectCount: inspectedMasterDrawings.length,
                      drawings: args.includeObjects ? inspectedMasterDrawings.map(function (drawing, index) {
                        return describeDrawing(drawing, index, Boolean(args.includeRaw));
                      }) : undefined,
                      layouts: inspectedLayouts,
                      raw: args.includeRaw ? serialized(inspectedMaster) : undefined,
                    });
                  }
                  results.push({
                    name: call.name,
                    masterCount: masterCount(),
                    masters: inspectedMasters,
                  });
                  break;
                }

                case "slides_inspect_themes": {
                  var themeMasters = [];
                  var firstThemeMaster = hasOwn(args, "masterIndex") ? Number(args.masterIndex) - 1 : 0;
                  var lastThemeMaster = hasOwn(args, "masterIndex") ? firstThemeMaster + 1 : masterCount();
                  if (firstThemeMaster < 0 || lastThemeMaster > masterCount()) {
                    throw new Error("幻灯片母版序号超出范围，当前共 " + masterCount() + " 个母版");
                  }
                  for (var themeMasterIndex = firstThemeMaster; themeMasterIndex < lastThemeMaster; themeMasterIndex += 1) {
                    var themeMaster = presentation.GetMaster(themeMasterIndex);
                    themeMasters.push({
                      masterIndex: themeMasterIndex + 1,
                      theme: describeTheme(safeCall(themeMaster, "GetTheme"), Boolean(args.includeRaw)),
                    });
                  }
                  results.push({
                    name: call.name,
                    masters: themeMasters,
                    currentSlideTheme: describeTheme(
                      safeCall(
                        hasOwn(args, "slide")
                          ? getSlide(args.slide)
                          : presentation.GetSlideByIndex(presentation.GetCurSlideIndex()),
                        "GetTheme"
                      ),
                      Boolean(args.includeRaw)
                    ),
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

                case "slides_inspect_animations": {
                  var firstAnimationSlide = hasOwn(args, "slide") ? Number(args.slide) - 1 : 0;
                  var lastAnimationSlide = hasOwn(args, "slide") ? firstAnimationSlide + 1 : slideCount();
                  if (firstAnimationSlide < 0 || lastAnimationSlide > slideCount()) throw new Error("幻灯片页码超出范围");
                  var animationSlides = [];
                  for (var animationSlideIndex = firstAnimationSlide; animationSlideIndex < lastAnimationSlide; animationSlideIndex += 1) {
                    var animationSlide = presentation.GetSlideByIndex(animationSlideIndex);
                    var timeline = safeCall(animationSlide, "GetTimeLine");
                    var effects = timeline ? (safeCall(timeline, "GetAllEffects") || []) : [];
                    var mainSequence = timeline ? safeCall(timeline, "GetMainSequence") : null;
                    var interactiveSequences = timeline ? (safeCall(timeline, "GetInteractiveSequences") || []) : [];
                    animationSlides.push({
                      slide: animationSlideIndex + 1,
                      mainSequenceCount: mainSequence ? safeCall(mainSequence, "GetCount") : 0,
                      interactiveSequenceCounts: interactiveSequences.map(function (sequence) {
                        return safeCall(sequence, "GetCount") || 0;
                      }),
                      effects: effects.map(function (effect, index) {
                        return describeAnimationEffect(effect, index);
                      }),
                    });
                  }
                  results.push({ name: call.name, slides: animationSlides });
                  break;
                }

                case "slides_inspect_comments": {
                  var inspectedComments = safeCall(presentation, "GetAllComments") || [];
                  results.push({
                    name: call.name,
                    comments: inspectedComments.map(function (comment, index) {
                      return describeComment(comment, index);
                    }),
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

                case "slides_move_slide": {
                  var movedSlide = getSlide(args.slide);
                  var moveDestination = Number(args.toIndex);
                  if (!Number.isInteger(moveDestination) || moveDestination < 1 || moveDestination > slideCount()) {
                    throw new Error("目标页码超出范围，当前共 " + slideCount() + " 页");
                  }
                  if (typeof movedSlide.MoveTo !== "function") throw new Error("当前 ONLYOFFICE 版本不支持移动幻灯片");
                  if (movedSlide.MoveTo(moveDestination - 1) === false) throw new Error("移动幻灯片失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    fromSlide: Number(args.slide),
                    toIndex: moveDestination,
                    slideCount: slideCount(),
                  });
                  break;
                }

                case "slides_set_visibility": {
                  var visibilitySlide = getSlide(args.slide);
                  if (typeof visibilitySlide.SetVisible !== "function") throw new Error("当前 ONLYOFFICE 版本不支持隐藏幻灯片");
                  if (visibilitySlide.SetVisible(Boolean(args.visible)) === false) throw new Error("设置幻灯片可见性失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    visible: safeCall(visibilitySlide, "GetVisible"),
                  });
                  break;
                }

                case "slides_set_size": {
                  if (typeof presentation.SetSizes !== "function") throw new Error("当前 ONLYOFFICE 版本不支持修改幻灯片尺寸");
                  var sizePreset = String(args.preset || "custom");
                  var sizeWidthMm = Number(args.widthMm);
                  var sizeHeightMm = Number(args.heightMm);
                  if (sizePreset === "wide") {
                    sizeWidthMm = 338.667;
                    sizeHeightMm = 190.5;
                  } else if (sizePreset === "standard") {
                    sizeWidthMm = 254;
                    sizeHeightMm = 190.5;
                  } else if (sizePreset !== "custom") {
                    throw new Error("不支持的幻灯片尺寸预设：" + sizePreset);
                  }
                  if (!isFinite(sizeWidthMm) || sizeWidthMm <= 0 || !isFinite(sizeHeightMm) || sizeHeightMm <= 0) {
                    throw new Error("自定义幻灯片尺寸需要有效的 widthMm 和 heightMm");
                  }
                  if (String(args.orientation || "landscape") === "portrait" && sizeWidthMm > sizeHeightMm) {
                    var portraitSwap = sizeWidthMm;
                    sizeWidthMm = sizeHeightMm;
                    sizeHeightMm = portraitSwap;
                  } else if (String(args.orientation || "landscape") === "landscape" && sizeHeightMm > sizeWidthMm) {
                    var landscapeSwap = sizeWidthMm;
                    sizeWidthMm = sizeHeightMm;
                    sizeHeightMm = landscapeSwap;
                  }
                  if (presentation.SetSizes(mmToEmu(sizeWidthMm), mmToEmu(sizeHeightMm)) === false) {
                    throw new Error("修改幻灯片尺寸失败");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    widthMm: emuToMm(safeCall(presentation, "GetWidth")),
                    heightMm: emuToMm(safeCall(presentation, "GetHeight")),
                  });
                  break;
                }

                case "slides_apply_layout": {
                  var layoutSlide = getSlide(args.slide);
                  var selectedMaster = getMaster(args.masterIndex || 1);
                  var selectedLayout = getLayout(selectedMaster, args.layoutIndex);
                  if (typeof layoutSlide.ApplyLayout !== "function") throw new Error("当前 ONLYOFFICE 版本不支持应用幻灯片版式");
                  if (layoutSlide.ApplyLayout(selectedLayout) === false) throw new Error("应用幻灯片版式失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    layout: locateLayout(selectedLayout, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_set_show_settings": {
                  if (!hasOwn(args, "loop")) throw new Error("slides_set_show_settings 至少需要 loop");
                  if (typeof presentation.SetLoopUntilStopped !== "function") {
                    throw new Error("当前 ONLYOFFICE 版本不支持设置循环放映");
                  }
                  if (presentation.SetLoopUntilStopped(Boolean(args.loop)) === false) throw new Error("设置循环放映失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    loop: safeCall(presentation, "GetLoopUntilStopped"),
                  });
                  break;
                }

                case "slides_apply_theme": {
                  var themeSource;
                  if (hasOwn(args, "sourceSlide")) themeSource = safeCall(getSlide(args.sourceSlide), "GetTheme");
                  else themeSource = safeCall(getMaster(args.sourceMasterIndex || 1), "GetTheme");
                  if (!themeSource) throw new Error("找不到可应用的主题");
                  if (hasOwn(args, "targetSlide")) {
                    var themeTargetSlide = getSlide(args.targetSlide);
                    if (typeof themeTargetSlide.ApplyTheme !== "function") throw new Error("当前 ONLYOFFICE 版本不支持向单页应用主题");
                    if (themeTargetSlide.ApplyTheme(themeSource) === false) throw new Error("应用幻灯片主题失败");
                  } else {
                    if (typeof presentation.ApplyTheme !== "function") throw new Error("当前 ONLYOFFICE 版本不支持应用演示文稿主题");
                    if (presentation.ApplyTheme(themeSource) === false) throw new Error("应用演示文稿主题失败");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    targetSlide: hasOwn(args, "targetSlide") ? Number(args.targetSlide) : "all",
                    theme: describeTheme(themeSource, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_set_theme": {
                  var editableTheme;
                  if (hasOwn(args, "slide")) editableTheme = safeCall(getSlide(args.slide), "GetTheme");
                  else editableTheme = safeCall(getMaster(args.masterIndex || 1), "GetTheme");
                  if (!editableTheme) throw new Error("找不到需要修改的主题");
                  if (Array.isArray(args.colors)) {
                    if (args.colors.length !== 12) throw new Error("主题颜色必须按 2 深色、2 浅色、6 强调色、超链接和已访问超链接提供 12 个颜色");
                    if (typeof Api.CreateThemeColorScheme !== "function") throw new Error("当前 ONLYOFFICE 版本不支持自定义主题颜色");
                    var customColorScheme = Api.CreateThemeColorScheme(
                      args.colors.map(apiColor),
                      String(args.colorSchemeName || "AI Bridge colors")
                    );
                    if (editableTheme.SetColorScheme(customColorScheme) === false) throw new Error("设置主题颜色失败");
                  }
                  if (isObject(args.fonts)) {
                    if (typeof Api.CreateThemeFontScheme !== "function") throw new Error("当前 ONLYOFFICE 版本不支持自定义主题字体");
                    var majorLatin = String(args.fonts.majorLatin || "");
                    var minorLatin = String(args.fonts.minorLatin || "");
                    if (!majorLatin || !minorLatin) throw new Error("自定义主题字体需要 majorLatin 和 minorLatin");
                    var customFontScheme = Api.CreateThemeFontScheme(
                      majorLatin,
                      String(args.fonts.majorEastAsian || majorLatin),
                      String(args.fonts.majorComplex || majorLatin),
                      minorLatin,
                      String(args.fonts.minorEastAsian || minorLatin),
                      String(args.fonts.minorComplex || minorLatin),
                      String(args.fonts.name || "AI Bridge fonts")
                    );
                    if (editableTheme.SetFontScheme(customFontScheme) === false) throw new Error("设置主题字体失败");
                  }
                  if (!Array.isArray(args.colors) && !isObject(args.fonts)) {
                    throw new Error("slides_set_theme 需要 colors 或 fonts");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    theme: describeTheme(editableTheme, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_create_layout": {
                  if (typeof Api.CreateLayout !== "function") throw new Error("当前 ONLYOFFICE 版本不支持创建自定义版式");
                  var layoutMasterIndex = Number(args.masterIndex || 1);
                  var layoutMaster = getMaster(layoutMasterIndex);
                  var createdLayout = Api.CreateLayout(layoutMaster);
                  if (!createdLayout) throw new Error("创建自定义版式失败");
                  if (hasOwn(args, "name") && typeof createdLayout.SetName === "function") createdLayout.SetName(String(args.name));
                  if (hasOwn(args, "fill")) createdLayout.SetBackground(createFill(args.fill));
                  else if (args.followMasterBackground && typeof createdLayout.FollowMasterBackground === "function") {
                    createdLayout.FollowMasterBackground();
                  }
                  var placeholderSpecs = Array.isArray(args.placeholders) ? args.placeholders : [];
                  for (var placeholderIndex = 0; placeholderIndex < placeholderSpecs.length; placeholderIndex += 1) {
                    var placeholderSpec = placeholderSpecs[placeholderIndex] || {};
                    var placeholderShape = Api.CreateShape(
                      String(placeholderSpec.shapeType || "rect"),
                      mmToEmu(asFinite(placeholderSpec.widthMm, 100)),
                      mmToEmu(asFinite(placeholderSpec.heightMm, 30)),
                      createFill(placeholderSpec.fill),
                      createStroke(placeholderSpec.line)
                    );
                    applyDrawingFrame(placeholderShape, {
                      xMm: asFinite(placeholderSpec.xMm, 15),
                      yMm: asFinite(placeholderSpec.yMm, 15),
                    });
                    if (hasOwn(placeholderSpec, "text")) replaceShapeText(placeholderShape, placeholderSpec);
                    if (typeof Api.CreatePlaceholder !== "function" || typeof placeholderShape.SetPlaceholder !== "function") {
                      throw new Error("当前 ONLYOFFICE 版本不支持版式占位符");
                    }
                    placeholderShape.SetPlaceholder(Api.CreatePlaceholder(String(placeholderSpec.type || "body")));
                    createdLayout.AddObject(placeholderShape);
                  }
                  if (Array.isArray(args.applyToSlides)) {
                    for (var applyLayoutIndex = 0; applyLayoutIndex < args.applyToSlides.length; applyLayoutIndex += 1) {
                      getSlide(args.applyToSlides[applyLayoutIndex]).ApplyLayout(createdLayout);
                    }
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    layout: locateLayout(createdLayout, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_add_template_shape": {
                  var templateTarget = getTemplateContainer(args);
                  var templateShape = Api.CreateShape(
                    String(args.shapeType || "rect"),
                    mmToEmu(asFinite(args.widthMm, 100)),
                    mmToEmu(asFinite(args.heightMm, 50)),
                    createFill(args.fill, args.fillColor),
                    createStroke(args.line, args.lineColor, args.lineWidthPt)
                  );
                  applyShapeOptions(templateShape, args);
                  applyDrawingFrame(templateShape, {
                    xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                    yMm: hasOwn(args, "yMm") ? args.yMm : 15,
                  });
                  if (hasOwn(args, "placeholderType")) {
                    if (typeof Api.CreatePlaceholder !== "function" || typeof templateShape.SetPlaceholder !== "function") {
                      throw new Error("当前 ONLYOFFICE 版本不支持占位符");
                    }
                    templateShape.SetPlaceholder(Api.CreatePlaceholder(String(args.placeholderType)));
                  }
                  if (templateTarget.container.AddObject(templateShape) === false) throw new Error("向母版或版式添加对象失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    scope: templateTarget.scope,
                    masterIndex: templateTarget.masterIndex,
                    layoutIndex: templateTarget.layoutIndex,
                    object: describeDrawing(
                      templateShape,
                      (safeCall(templateTarget.container, "GetAllDrawings") || []).indexOf(templateShape),
                      Boolean(args.includeRaw)
                    ),
                  });
                  break;
                }

                case "slides_manage_template_object": {
                  var managedTemplateObject = resolveTemplateDrawing(args);
                  var templateObjectAction = String(args.action || "update");
                  if (templateObjectAction === "delete") {
                    var removedTemplateObject = false;
                    if (typeof managedTemplateObject.container.RemoveObject === "function") {
                      removedTemplateObject = managedTemplateObject.container.RemoveObject(managedTemplateObject.drawing);
                    } else if (typeof managedTemplateObject.drawing.Delete === "function") {
                      removedTemplateObject = managedTemplateObject.drawing.Delete();
                    }
                    if (removedTemplateObject === false) throw new Error("删除母版或版式对象失败");
                  } else if (templateObjectAction === "update") {
                    if (drawingKind(managedTemplateObject.drawing) === "shape") {
                      applyShapeOptions(managedTemplateObject.drawing, args);
                    } else {
                      applyDrawingFrame(managedTemplateObject.drawing, args);
                      if (hasOwn(args, "line") || hasOwn(args, "lineColor") || hasOwn(args, "lineWidthPt")) {
                        if (typeof managedTemplateObject.drawing.SetOutLine !== "function") {
                          throw new Error("目标模板对象不支持边框");
                        }
                        managedTemplateObject.drawing.SetOutLine(createStroke(args.line, args.lineColor, args.lineWidthPt));
                      }
                    }
                    if (hasOwn(args, "newName") && typeof managedTemplateObject.drawing.SetName === "function") {
                      managedTemplateObject.drawing.SetName(String(args.newName));
                    }
                  } else {
                    throw new Error("不支持的模板对象动作：" + templateObjectAction);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: templateObjectAction,
                    scope: managedTemplateObject.scope,
                    masterIndex: managedTemplateObject.masterIndex,
                    layoutIndex: managedTemplateObject.layoutIndex,
                    object: templateObjectAction === "delete"
                      ? null
                      : describeDrawing(
                        managedTemplateObject.drawing,
                        managedTemplateObject.index,
                        Boolean(args.includeRaw)
                      ),
                  });
                  break;
                }

                case "slides_set_template_background": {
                  var backgroundTarget = getTemplateContainer(args);
                  var templateBackgroundMode = String(args.mode || "custom");
                  if (templateBackgroundMode === "clear") {
                    if (typeof backgroundTarget.container.ClearBackground !== "function") throw new Error("当前 ONLYOFFICE 版本不支持清除模板背景");
                    backgroundTarget.container.ClearBackground();
                  } else if (templateBackgroundMode === "master" && backgroundTarget.scope === "layout") {
                    if (typeof backgroundTarget.container.FollowMasterBackground !== "function") throw new Error("当前 ONLYOFFICE 版本不支持跟随母版背景");
                    backgroundTarget.container.FollowMasterBackground();
                  } else if (templateBackgroundMode === "custom") {
                    if (!hasOwn(args, "fill")) throw new Error("自定义模板背景需要 fill");
                    backgroundTarget.container.SetBackground(createFill(args.fill));
                  } else {
                    throw new Error("不支持的模板背景模式：" + templateBackgroundMode);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    scope: backgroundTarget.scope,
                    masterIndex: backgroundTarget.masterIndex,
                    layoutIndex: backgroundTarget.layoutIndex,
                    mode: templateBackgroundMode,
                  });
                  break;
                }

                case "slides_set_text_content": {
                  var textContentShape = resolveDrawing(args, "shape");
                  replaceShapeParagraphs(textContentShape.drawing, args.paragraphs);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    paragraphs: getParagraphsFromShape(textContentShape.drawing).length,
                    object: describeDrawing(textContentShape.drawing, textContentShape.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_format_paragraphs": {
                  var paragraphShape = resolveDrawing(args, "shape");
                  var targetParagraphs = getParagraphsFromShape(paragraphShape.drawing);
                  var requestedParagraphs = Array.isArray(args.paragraphIndexes)
                    ? args.paragraphIndexes.map(Number)
                    : null;
                  var formattedParagraphs = 0;
                  for (var targetParagraphIndex = 0; targetParagraphIndex < targetParagraphs.length; targetParagraphIndex += 1) {
                    if (requestedParagraphs && requestedParagraphs.indexOf(targetParagraphIndex + 1) === -1) continue;
                    applyParagraphFormat(targetParagraphs[targetParagraphIndex], args);
                    formattedParagraphs += 1;
                  }
                  if (!formattedParagraphs) throw new Error("找不到需要格式化的段落");
                  changed += formattedParagraphs;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    formattedParagraphs: formattedParagraphs,
                    object: describeDrawing(paragraphShape.drawing, paragraphShape.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_update_object": {
                  var updatedObject = resolveDrawing(args);
                  applyDrawingFrame(updatedObject.drawing, args);
                  if (hasOwn(args, "line") || hasOwn(args, "lineColor") || hasOwn(args, "lineWidthPt")) {
                    if (typeof updatedObject.drawing.SetOutLine !== "function") {
                      throw new Error("目标对象不支持边框");
                    }
                    updatedObject.drawing.SetOutLine(createStroke(args.line, args.lineColor, args.lineWidthPt));
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    object: describeDrawing(updatedObject.drawing, updatedObject.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_set_hyperlink": {
                  var hyperlinkObject = resolveDrawing(args);
                  if (typeof hyperlinkObject.drawing.SetHyperlink !== "function") {
                    throw new Error("目标对象不支持超链接");
                  }
                  var hyperlinkAction = String(args.action || "external");
                  var hyperlink = null;
                  var hyperlinkAddress = null;
                  if (hyperlinkAction !== "remove") {
                    if (typeof Api.CreateHyperlink !== "function") throw new Error("当前 ONLYOFFICE 版本不支持创建超链接");
                    if (hyperlinkAction === "external") {
                      hyperlinkAddress = String(args.url || "");
                      if (!/^(?:https?|mailto|ftp):/i.test(hyperlinkAddress)) {
                        throw new Error("外部超链接只允许 http、https、mailto 或 ftp 协议");
                      }
                    } else if (hyperlinkAction === "firstSlide") {
                      hyperlinkAddress = "ppaction://hlinkshowjump?jump=firstslide";
                    } else if (hyperlinkAction === "lastSlide") {
                      hyperlinkAddress = "ppaction://hlinkshowjump?jump=lastslide";
                    } else if (hyperlinkAction === "nextSlide") {
                      hyperlinkAddress = "ppaction://hlinkshowjump?jump=nextslide";
                    } else if (hyperlinkAction === "previousSlide") {
                      hyperlinkAddress = "ppaction://hlinkshowjump?jump=previousslide";
                    } else if (hyperlinkAction === "slide") {
                      var hyperlinkSlide = Number(args.targetSlide);
                      if (!Number.isInteger(hyperlinkSlide) || hyperlinkSlide < 1 || hyperlinkSlide > slideCount()) {
                        throw new Error("超链接目标幻灯片页码超出范围");
                      }
                      hyperlinkAddress = "ppaction://hlinksldjumpslide" + (hyperlinkSlide - 1);
                    } else {
                      throw new Error("不支持的超链接动作：" + hyperlinkAction);
                    }
                    hyperlink = Api.CreateHyperlink(hyperlinkAddress, String(args.tooltip || ""));
                  }
                  if (hyperlinkObject.drawing.SetHyperlink(hyperlink) === false) throw new Error("设置对象超链接失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    action: hyperlinkAction,
                    link: hyperlinkAddress,
                    object: describeDrawing(hyperlinkObject.drawing, hyperlinkObject.index, false),
                  });
                  break;
                }

                case "slides_set_notes": {
                  var notesSlide = getSlide(args.slide);
                  var notesText = String(args.text || "");
                  var notesSucceeded = false;
                  if (args.append) {
                    if (typeof notesSlide.AddNotesText !== "function") throw new Error("当前 ONLYOFFICE 版本不支持演讲者备注");
                    notesSucceeded = notesSlide.AddNotesText(notesText);
                  } else {
                    var notesPage = safeCall(notesSlide, "GetNotesPage");
                    var notesBody = notesPage ? safeCall(notesPage, "GetBodyShape") : null;
                    var notesContent = notesBody && (safeCall(notesBody, "GetContent") || safeCall(notesBody, "GetDocContent"));
                    if (notesContent && typeof notesContent.Push === "function") {
                      if (typeof notesContent.RemoveAllElements === "function") notesContent.RemoveAllElements();
                      var notesParagraph = Api.CreateParagraph();
                      notesParagraph.AddText(notesText);
                      notesContent.Push(notesParagraph);
                      notesSucceeded = true;
                    } else if (typeof notesSlide.AddNotesText === "function") {
                      notesSucceeded = notesSlide.AddNotesText(notesText);
                    }
                  }
                  if (notesSucceeded === false) throw new Error("设置演讲者备注失败");
                  changed += 1;
                  var updatedNotesPage = safeCall(notesSlide, "GetNotesPage");
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    notes: updatedNotesPage ? safeCall(updatedNotesPage, "GetBodyShapeText") : notesText,
                  });
                  break;
                }

                case "slides_add_comment": {
                  var commentSlide = getSlide(args.slide);
                  if (!args.text) throw new Error("批注内容不能为空");
                  if (typeof commentSlide.AddComment !== "function") throw new Error("当前 ONLYOFFICE 版本不支持幻灯片批注");
                  var commentAdded = commentSlide.AddComment(
                    mmToEmu(hasOwn(args, "xMm") ? args.xMm : 0),
                    mmToEmu(hasOwn(args, "yMm") ? args.yMm : 0),
                    String(args.text),
                    hasOwn(args, "author") ? String(args.author) : undefined,
                    hasOwn(args, "userId") ? String(args.userId) : undefined
                  );
                  if (commentAdded === false) throw new Error("添加幻灯片批注失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    commentCount: (safeCall(presentation, "GetAllComments") || []).length,
                  });
                  break;
                }

                case "slides_set_transition": {
                  var transitionSlide = getSlide(args.slide);
                  if (typeof transitionSlide.SetSlideShowTransition !== "function") {
                    throw new Error("当前 ONLYOFFICE 版本不支持幻灯片切换");
                  }
                  var configuredTransition = null;
                  if (!args.clear) {
                    if (typeof Api.CreateSlideShowTransition !== "function") {
                      throw new Error("当前 ONLYOFFICE 版本不支持创建幻灯片切换");
                    }
                    configuredTransition = Api.CreateSlideShowTransition();
                    if (hasOwn(args, "effect")) configuredTransition.SetEntryEffect(String(args.effect));
                    if (hasOwn(args, "speed")) configuredTransition.SetSpeed(String(args.speed));
                    if (hasOwn(args, "durationMs") && typeof configuredTransition.SetDuration === "function") {
                      configuredTransition.SetDuration(Math.max(0, Math.round(asFinite(args.durationMs, 0))));
                    }
                    if (hasOwn(args, "advanceOnClick")) configuredTransition.SetAdvanceOnClick(Boolean(args.advanceOnClick));
                    if (hasOwn(args, "advanceOnTime")) configuredTransition.SetAdvanceOnTime(Boolean(args.advanceOnTime));
                    if (hasOwn(args, "advanceTimeMs")) {
                      configuredTransition.SetAdvanceTime(Math.max(0, Math.round(asFinite(args.advanceTimeMs, 0))));
                    }
                  }
                  if (transitionSlide.SetSlideShowTransition(configuredTransition) === false) {
                    throw new Error("设置幻灯片切换失败");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    transition: describeTransition(safeCall(transitionSlide, "GetSlideShowTransition")),
                  });
                  break;
                }

                case "slides_manage_animation": {
                  var managedAnimationSlide = getSlide(args.slide);
                  var managedTimeline = safeCall(managedAnimationSlide, "GetTimeLine");
                  if (!managedTimeline) throw new Error("当前 ONLYOFFICE 版本不支持对象动画");
                  var animationAction = String(args.action || "add");
                  var managedEffect = null;
                  if (animationAction === "add") {
                    var animationTarget = resolveDrawing(args, null).drawing;
                    var sequenceKind = String(args.sequence || "main");
                    var targetSequence;
                    if (sequenceKind === "interactive") {
                      if (!isObject(args.triggerObject)) throw new Error("交互动画需要 triggerObject");
                      var triggerSelector = { slide: args.slide };
                      if (hasOwn(args.triggerObject, "objectId")) triggerSelector.objectId = args.triggerObject.objectId;
                      if (hasOwn(args.triggerObject, "objectIndex")) triggerSelector.objectIndex = args.triggerObject.objectIndex;
                      if (hasOwn(args.triggerObject, "name")) triggerSelector.name = args.triggerObject.name;
                      var triggerDrawing = resolveDrawing(triggerSelector).drawing;
                      targetSequence = safeCall(managedTimeline, "AddInteractiveSequence", triggerDrawing);
                    } else if (sequenceKind === "main") {
                      targetSequence = safeCall(managedTimeline, "GetMainSequence");
                    } else {
                      throw new Error("animation.sequence 必须是 main 或 interactive");
                    }
                    if (!targetSequence || typeof targetSequence.AddEffect !== "function") {
                      throw new Error("无法创建动画序列");
                    }
                    managedEffect = targetSequence.AddEffect(
                      animationTarget,
                      String(args.effectType || "entranceFade"),
                      String(args.trigger || "onclick")
                    );
                    if (!managedEffect) throw new Error("添加动画失败：请检查 effectType");
                  } else if (animationAction === "clear") {
                    var clearSequenceKind = String(args.sequence || "main");
                    var clearSequence;
                    if (clearSequenceKind === "interactive") {
                      if (!isObject(args.triggerObject)) throw new Error("清除交互动画需要 triggerObject");
                      var clearTriggerSelector = { slide: args.slide };
                      if (hasOwn(args.triggerObject, "objectId")) clearTriggerSelector.objectId = args.triggerObject.objectId;
                      if (hasOwn(args.triggerObject, "objectIndex")) clearTriggerSelector.objectIndex = args.triggerObject.objectIndex;
                      if (hasOwn(args.triggerObject, "name")) clearTriggerSelector.name = args.triggerObject.name;
                      clearSequence = safeCall(
                        managedTimeline,
                        "AddInteractiveSequence",
                        resolveDrawing(clearTriggerSelector).drawing
                      );
                    } else {
                      clearSequence = safeCall(managedTimeline, "GetMainSequence");
                    }
                    if (!clearSequence || safeCall(clearSequence, "RemoveAllEffects") === false) throw new Error("清除动画序列失败");
                  } else {
                    var allManagedEffects = safeCall(managedTimeline, "GetAllEffects") || [];
                    var managedEffectIndex = Number(args.effectIndex);
                    if (!Number.isInteger(managedEffectIndex) || managedEffectIndex < 0 || managedEffectIndex >= allManagedEffects.length) {
                      throw new Error("effectIndex 超出范围");
                    }
                    managedEffect = allManagedEffects[managedEffectIndex];
                    if (animationAction === "delete") {
                      if (safeCall(managedEffect, "Delete") === false) throw new Error("删除动画失败");
                      managedEffect = null;
                    } else if (animationAction !== "update") {
                      throw new Error("不支持的动画动作：" + animationAction);
                    }
                  }
                  if (managedEffect) {
                    if (hasOwn(args, "trigger") && typeof managedEffect.SetTriggerType === "function") {
                      if (managedEffect.SetTriggerType(String(args.trigger)) === false) throw new Error("设置动画触发方式失败");
                    }
                    if (hasOwn(args, "durationMs") && typeof managedEffect.SetDuration === "function") {
                      if (managedEffect.SetDuration(Math.max(0, Math.round(asFinite(args.durationMs, 0)))) === false) {
                        throw new Error("设置动画时长失败");
                      }
                    }
                    if (hasOwn(args, "delayMs") && typeof managedEffect.SetDelay === "function") {
                      if (managedEffect.SetDelay(Math.max(0, Math.round(asFinite(args.delayMs, 0)))) === false) {
                        throw new Error("设置动画延迟失败");
                      }
                    }
                    if (hasOwn(args, "repeatCount") && typeof managedEffect.SetRepeatCount === "function") {
                      if (managedEffect.SetRepeatCount(Math.max(0, asFinite(args.repeatCount, 1))) === false) {
                        throw new Error("设置动画重复次数失败");
                      }
                    }
                    if (hasOwn(args, "toIndex") && typeof managedEffect.MoveTo === "function") {
                      var animationDestination = Number(args.toIndex);
                      if (!Number.isInteger(animationDestination) || animationDestination < 0) throw new Error("toIndex 必须是从 0 开始的整数");
                      if (managedEffect.MoveTo(animationDestination) === false) throw new Error("移动动画顺序失败");
                    }
                  }
                  changed += 1;
                  var finalEffects = safeCall(managedTimeline, "GetAllEffects") || [];
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    action: animationAction,
                    effects: finalEffects.map(function (effect, index) {
                      return describeAnimationEffect(effect, index);
                    }),
                  });
                  break;
                }

                case "slides_manage_comment": {
                  var commentAction = String(args.action || "update");
                  var resolvedComment = resolveComment(args);
                  var targetComment = resolvedComment.comment;
                  if (commentAction === "update") {
                    if (hasOwn(args, "text")) safeCall(targetComment, "SetText", String(args.text));
                    if (hasOwn(args, "author")) {
                      if (typeof targetComment.SetAuthorName === "function") targetComment.SetAuthorName(String(args.author));
                      else safeCall(targetComment, "SetAutorName", String(args.author));
                    }
                    if (hasOwn(args, "userId")) safeCall(targetComment, "SetUserId", String(args.userId));
                    if (hasOwn(args, "solved")) safeCall(targetComment, "SetSolved", Boolean(args.solved));
                    if (hasOwn(args, "xMm") || hasOwn(args, "yMm")) {
                      safeCall(targetComment, "SetPosition", mmToEmu(asFinite(args.xMm, 0)), mmToEmu(asFinite(args.yMm, 0)));
                    }
                  } else if (commentAction === "addReply") {
                    if (!hasOwn(args, "text")) throw new Error("添加批注回复需要 text");
                    var reply = safeCall(targetComment, "AddReply", String(args.text), String(args.author || ""), String(args.userId || ""));
                    if (reply === false || reply === null) throw new Error("添加批注回复失败");
                  } else if (commentAction === "removeReplies") {
                    if (safeCall(targetComment, "RemoveReplies", Math.max(0, Math.round(asFinite(args.start, 0))), Math.max(1, Math.round(asFinite(args.count, 1)))) === false) {
                      throw new Error("删除批注回复失败");
                    }
                  } else if (commentAction === "delete") {
                    if (safeCall(targetComment, "Delete") === false) throw new Error("删除批注失败");
                  } else {
                    throw new Error("不支持的批注动作：" + commentAction);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    action: commentAction,
                    comment: commentAction === "delete" ? null : describeComment(targetComment, resolvedComment.index),
                  });
                  break;
                }

                case "slides_add_table": {
                  if (typeof Api.CreateTable !== "function") throw new Error("当前 ONLYOFFICE 版本不支持创建表格");
                  var tableRows = Math.round(asFinite(args.rows, 0));
                  var tableColumns = Math.round(asFinite(args.columns, 0));
                  if (tableRows < 1 || tableRows > 200 || tableColumns < 1 || tableColumns > 50) {
                    throw new Error("表格需要 1-200 行、1-50 列");
                  }
                  var tableSlide = getSlide(args.slide);
                  var addedTable = Api.CreateTable(tableRows, tableColumns);
                  if (!addedTable) throw new Error("创建 PPT 表格失败");
                  var addedTableFrame = {
                    xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                    yMm: hasOwn(args, "yMm") ? args.yMm : 35,
                    widthMm: hasOwn(args, "widthMm") ? args.widthMm : 220,
                    heightMm: hasOwn(args, "heightMm") ? args.heightMm : Math.max(20, tableRows * 12),
                  };
                  if (hasOwn(args, "name")) addedTableFrame.name = args.name;
                  applyDrawingFrame(addedTable, addedTableFrame);
                  if (Array.isArray(args.data)) {
                    for (var addTableRow = 0; addTableRow < Math.min(tableRows, args.data.length); addTableRow += 1) {
                      var addTableValues = Array.isArray(args.data[addTableRow]) ? args.data[addTableRow] : [args.data[addTableRow]];
                      for (var addTableColumn = 0; addTableColumn < Math.min(tableColumns, addTableValues.length); addTableColumn += 1) {
                        applyTableCellFormat(
                          getTableCell(addedTable, addTableRow + 1, addTableColumn + 1),
                          tableCellSpec(addTableValues[addTableColumn])
                        );
                      }
                    }
                  }
                  if (isObject(args.header)) {
                    for (var headerColumn = 1; headerColumn <= tableColumns; headerColumn += 1) {
                      applyTableCellFormat(getTableCell(addedTable, 1, headerColumn), args.header);
                    }
                  }
                  tableSlide.AddObject(addedTable);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    table: describeDrawing(addedTable, slideDrawings(tableSlide).indexOf(addedTable), Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_set_table_cell": {
                  var cellTable = resolveDrawing(args, "table");
                  var selectedCell = getTableCell(cellTable.drawing, args.row, args.column);
                  applyTableCellFormat(selectedCell, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    row: Number(args.row),
                    column: Number(args.column),
                    text: String(safeCall(selectedCell, "GetText") || ""),
                    table: describeDrawing(cellTable.drawing, cellTable.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_edit_table": {
                  var editedTable = resolveDrawing(args, "table");
                  var tableAction = String(args.action || "");
                  var editCell = null;
                  if (hasOwn(args, "row")) editCell = getTableCell(editedTable.drawing, args.row, hasOwn(args, "column") ? args.column : 1);
                  if (tableAction === "addRow") {
                    if (typeof editedTable.drawing.AddRow !== "function") throw new Error("当前 ONLYOFFICE 版本不支持添加表格行");
                    editedTable.drawing.AddRow(editCell || undefined, String(args.position || "after") === "before");
                  } else if (tableAction === "addColumn") {
                    if (typeof editedTable.drawing.AddColumn !== "function") throw new Error("当前 ONLYOFFICE 版本不支持添加表格列");
                    editedTable.drawing.AddColumn(editCell || undefined, String(args.position || "after") === "before");
                  } else if (tableAction === "removeRow") {
                    if (!editCell) throw new Error("删除表格行需要 row");
                    if (editedTable.drawing.RemoveRow(editCell) === false) throw new Error("删除表格行失败");
                  } else if (tableAction === "removeColumn") {
                    if (!editCell) throw new Error("删除表格列需要 row 和 column");
                    if (editedTable.drawing.RemoveColumn(editCell) === false) throw new Error("删除表格列失败");
                  } else if (tableAction === "mergeCells") {
                    var mergeCells = [];
                    var mergeRowStart = Math.round(asFinite(args.rowStart, args.row));
                    var mergeRowEnd = Math.round(asFinite(args.rowEnd, mergeRowStart));
                    var mergeColumnStart = Math.round(asFinite(args.columnStart, args.column));
                    var mergeColumnEnd = Math.round(asFinite(args.columnEnd, mergeColumnStart));
                    if (mergeRowEnd < mergeRowStart || mergeColumnEnd < mergeColumnStart) {
                      throw new Error("合并单元格范围无效");
                    }
                    for (var mergeRow = mergeRowStart; mergeRow <= mergeRowEnd; mergeRow += 1) {
                      for (var mergeColumn = mergeColumnStart; mergeColumn <= mergeColumnEnd; mergeColumn += 1) {
                        mergeCells.push(getTableCell(editedTable.drawing, mergeRow, mergeColumn));
                      }
                    }
                    if (mergeCells.length < 2) throw new Error("至少选择两个单元格进行合并");
                    if (!editedTable.drawing.MergeCells(mergeCells)) throw new Error("合并单元格失败");
                  } else if (tableAction === "splitCell") {
                    if (!editCell || !hasOwn(args, "column")) throw new Error("拆分单元格需要 row 和 column");
                    if (typeof editCell.Split !== "function") throw new Error("当前 ONLYOFFICE 版本不支持拆分单元格");
                    var splitRows = Math.max(1, Math.round(asFinite(args.rows, 1)));
                    var splitColumns = Math.max(1, Math.round(asFinite(args.columns, 1)));
                    if (splitRows > 20 || splitColumns > 20) throw new Error("单元格最多拆分为 20×20");
                    if (editCell.Split(splitRows, splitColumns) === false) throw new Error("拆分单元格失败");
                  } else {
                    throw new Error("不支持的表格编辑动作：" + tableAction);
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    action: tableAction,
                    table: describeDrawing(editedTable.drawing, editedTable.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_format_table": {
                  var formattedTable = resolveDrawing(args, "table");
                  applyDrawingFrame(formattedTable.drawing, args);
                  if (typeof formattedTable.drawing.SetTableLook === "function" && isObject(args.tableLook)) {
                    formattedTable.drawing.SetTableLook(
                      Boolean(args.tableLook.firstColumn),
                      Boolean(args.tableLook.firstRow),
                      Boolean(args.tableLook.lastColumn),
                      Boolean(args.tableLook.lastRow),
                      Boolean(args.tableLook.horizontalBanding),
                      Boolean(args.tableLook.verticalBanding)
                    );
                  }
                  if (Array.isArray(args.columnWidthsMm) && typeof formattedTable.drawing.SetColumnWidth === "function") {
                    for (var tableWidthIndex = 0; tableWidthIndex < args.columnWidthsMm.length; tableWidthIndex += 1) {
                      formattedTable.drawing.SetColumnWidth(tableWidthIndex, mmToEmu(args.columnWidthsMm[tableWidthIndex]));
                    }
                  }
                  var formattedDimensions = tableDimensions(formattedTable.drawing);
                  if (Array.isArray(args.rowHeightsMm)) {
                    for (var tableHeightIndex = 0; tableHeightIndex < Math.min(args.rowHeightsMm.length, formattedDimensions.rows); tableHeightIndex += 1) {
                      var formattedRow = formattedTable.drawing.GetRow(tableHeightIndex);
                      if (formattedRow && typeof formattedRow.SetHeight === "function") {
                        formattedRow.SetHeight(mmToEmu(args.rowHeightsMm[tableHeightIndex]));
                      }
                    }
                  }
                  var formatRowStart = Math.max(1, Math.round(asFinite(args.rowStart, 1)));
                  var formatRowEnd = Math.min(formattedDimensions.rows, Math.round(asFinite(args.rowEnd, formattedDimensions.rows)));
                  var formatColumnStart = Math.max(1, Math.round(asFinite(args.columnStart, 1)));
                  var formatColumnEnd = Math.min(formattedDimensions.columns, Math.round(asFinite(args.columnEnd, formattedDimensions.columns)));
                  for (var formatTableRow = formatRowStart; formatTableRow <= formatRowEnd; formatTableRow += 1) {
                    for (var formatTableColumn = formatColumnStart; formatTableColumn <= formatColumnEnd; formatTableColumn += 1) {
                      applyTableCellFormat(getTableCell(formattedTable.drawing, formatTableRow, formatTableColumn), args);
                    }
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    table: describeDrawing(formattedTable.drawing, formattedTable.index, Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_align_objects": {
                  var alignedTargets = resolveDrawingTargets(args.slide, args.targets);
                  var alignment = hasOwn(args, "align") ? String(args.align) : null;
                  var distribution = hasOwn(args, "distribute") ? String(args.distribute) : null;
                  if (!alignment && !distribution) throw new Error("slides_align_objects 需要 align 或 distribute");
                  var alignSlide = getSlide(args.slide);
                  var alignFrames = alignedTargets.map(function (item) {
                    var frame = getDrawingFrame(item.drawing);
                    return {
                      item: item,
                      x: asFinite(frame.x, 0),
                      y: asFinite(frame.y, 0),
                      width: Math.max(1, asFinite(frame.width, 1)),
                      height: Math.max(1, asFinite(frame.height, 1)),
                    };
                  });
                  var relativeToSlide = String(args.relativeTo || "selection") === "slide";
                  var boundsLeft = relativeToSlide ? 0 : Math.min.apply(null, alignFrames.map(function (frame) { return frame.x; }));
                  var boundsTop = relativeToSlide ? 0 : Math.min.apply(null, alignFrames.map(function (frame) { return frame.y; }));
                  var boundsRight = relativeToSlide
                    ? asFinite(safeCall(alignSlide, "GetWidth"), asFinite(safeCall(presentation, "GetWidth"), 0))
                    : Math.max.apply(null, alignFrames.map(function (frame) { return frame.x + frame.width; }));
                  var boundsBottom = relativeToSlide
                    ? asFinite(safeCall(alignSlide, "GetHeight"), asFinite(safeCall(presentation, "GetHeight"), 0))
                    : Math.max.apply(null, alignFrames.map(function (frame) { return frame.y + frame.height; }));
                  for (var alignIndex = 0; alignIndex < alignFrames.length; alignIndex += 1) {
                    var alignedFrame = alignFrames[alignIndex];
                    if (alignment === "left") alignedFrame.x = boundsLeft;
                    else if (alignment === "center") alignedFrame.x = boundsLeft + (boundsRight - boundsLeft - alignedFrame.width) / 2;
                    else if (alignment === "right") alignedFrame.x = boundsRight - alignedFrame.width;
                    else if (alignment === "top") alignedFrame.y = boundsTop;
                    else if (alignment === "middle") alignedFrame.y = boundsTop + (boundsBottom - boundsTop - alignedFrame.height) / 2;
                    else if (alignment === "bottom") alignedFrame.y = boundsBottom - alignedFrame.height;
                    else if (alignment) throw new Error("不支持的对象对齐方式：" + alignment);
                  }
                  if (distribution) {
                    if (alignFrames.length < 3) throw new Error("等间距分布至少需要三个对象");
                    var horizontal = distribution === "horizontal";
                    if (!horizontal && distribution !== "vertical") throw new Error("不支持的对象分布方式：" + distribution);
                    alignFrames.sort(function (left, right) {
                      return horizontal ? left.x - right.x : left.y - right.y;
                    });
                    var totalSize = alignFrames.reduce(function (sum, frame) {
                      return sum + (horizontal ? frame.width : frame.height);
                    }, 0);
                    var distributionStart = horizontal ? boundsLeft : boundsTop;
                    var distributionEnd = horizontal ? boundsRight : boundsBottom;
                    var distributionGap = (distributionEnd - distributionStart - totalSize) / (alignFrames.length - 1);
                    var cursor = distributionStart;
                    for (var distributeIndex = 0; distributeIndex < alignFrames.length; distributeIndex += 1) {
                      if (horizontal) alignFrames[distributeIndex].x = cursor;
                      else alignFrames[distributeIndex].y = cursor;
                      cursor += (horizontal ? alignFrames[distributeIndex].width : alignFrames[distributeIndex].height) + distributionGap;
                    }
                  }
                  for (var setAlignedIndex = 0; setAlignedIndex < alignFrames.length; setAlignedIndex += 1) {
                    alignFrames[setAlignedIndex].item.drawing.SetPosition(
                      Math.round(alignFrames[setAlignedIndex].x),
                      Math.round(alignFrames[setAlignedIndex].y)
                    );
                  }
                  changed += alignedTargets.length;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    alignedObjects: alignedTargets.length,
                    align: alignment,
                    distribute: distribution,
                  });
                  break;
                }

                case "slides_group_objects": {
                  var groupAction = String(args.action || "group");
                  var groupingSlide = getSlide(args.slide);
                  if (groupAction === "group") {
                    var groupTargets = resolveDrawingTargets(args.slide, args.targets);
                    if (groupTargets.length < 2) throw new Error("组合至少需要两个对象");
                    if (typeof groupingSlide.GroupDrawings !== "function") throw new Error("当前 ONLYOFFICE 版本不支持组合对象");
                    var createdGroup = groupingSlide.GroupDrawings(groupTargets.map(function (item) { return item.drawing; }));
                    if (!createdGroup) throw new Error("组合对象失败");
                    if (hasOwn(args, "name") && typeof createdGroup.SetName === "function") createdGroup.SetName(String(args.name));
                    changed += 1;
                    results.push({
                      name: call.name,
                      slide: Number(args.slide),
                      action: groupAction,
                      group: describeDrawing(createdGroup, slideDrawings(groupingSlide).indexOf(createdGroup), Boolean(args.includeRaw)),
                    });
                  } else if (groupAction === "ungroup") {
                    var ungrouped = resolveDrawing(args, "group");
                    if (typeof ungrouped.drawing.Ungroup !== "function") throw new Error("当前 ONLYOFFICE 版本不支持取消组合");
                    if (ungrouped.drawing.Ungroup() === false) throw new Error("取消组合失败");
                    changed += 1;
                    results.push({
                      name: call.name,
                      slide: Number(args.slide),
                      action: groupAction,
                      objectCount: slideDrawings(groupingSlide).length,
                    });
                  } else {
                    throw new Error("不支持的组合动作：" + groupAction);
                  }
                  break;
                }

                case "slides_reorder_object": {
                  var reordered = resolveDrawing(args);
                  var reorderAction = String(args.action || "");
                  var reorderDrawings = slideDrawings(reordered.slide).slice();
                  var reorderIndex = reorderDrawings.indexOf(reordered.drawing);
                  if (reorderIndex < 0) throw new Error("找不到需要调整层级的对象");
                  reorderDrawings.splice(reorderIndex, 1);
                  var reorderedIndex;
                  if (reorderAction === "front") reorderedIndex = reorderDrawings.length;
                  else if (reorderAction === "back") reorderedIndex = 0;
                  else if (reorderAction === "forward") reorderedIndex = Math.min(reorderDrawings.length, reorderIndex + 1);
                  else if (reorderAction === "backward") reorderedIndex = Math.max(0, reorderIndex - 1);
                  else throw new Error("不支持的对象层级动作：" + reorderAction);
                  reorderDrawings.splice(reorderedIndex, 0, reordered.drawing);
                  for (var detachIndex = 0; detachIndex < reorderDrawings.length; detachIndex += 1) {
                    if (reordered.slide.RemoveObject(reorderDrawings[detachIndex]) === false) {
                      throw new Error("调整对象层级时无法移除对象");
                    }
                  }
                  for (var attachIndex = 0; attachIndex < reorderDrawings.length; attachIndex += 1) {
                    if (reordered.slide.AddObject(reorderDrawings[attachIndex]) === false) {
                      throw new Error("调整对象层级时无法重新添加对象");
                    }
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    action: reorderAction,
                    objectIndex: slideDrawings(reordered.slide).indexOf(reordered.drawing),
                  });
                  break;
                }

                case "slides_add_connector": {
                  var connectorSlide = getSlide(args.slide);
                  var startX = asFinite(args.startXmm, 0);
                  var startY = asFinite(args.startYmm, 0);
                  var endX = asFinite(args.endXmm, startX + 30);
                  var endY = asFinite(args.endYmm, startY);
                  var connectorLeft = Math.min(startX, endX);
                  var connectorTop = Math.min(startY, endY);
                  var connectorWidth = Math.max(0.1, Math.abs(endX - startX));
                  var connectorHeight = Math.max(0.1, Math.abs(endY - startY));
                  var connectorType = String(args.connectorType || "straightConnector1");
                  var connector = Api.CreateShape(
                    connectorType,
                    mmToEmu(connectorWidth),
                    mmToEmu(connectorHeight),
                    Api.CreateNoFill(),
                    createStroke(args.line, args.lineColor, args.lineWidthPt)
                  );
                  connector.SetPosition(mmToEmu(connectorLeft), mmToEmu(connectorTop));
                  if ((endX < startX) !== (endY < startY) && typeof connector.SetFlipV === "function") connector.SetFlipV(true);
                  if (hasOwn(args, "name") && typeof connector.SetName === "function") connector.SetName(String(args.name));
                  connectorSlide.AddObject(connector);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    object: describeDrawing(connector, slideDrawings(connectorSlide).indexOf(connector), Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_add_freeform": {
                  if (typeof Api.CreateCustomGeometry !== "function") throw new Error("当前 ONLYOFFICE 版本不支持自由形状");
                  if (!Array.isArray(args.paths) || !args.paths.length) throw new Error("自由形状至少需要一条路径");
                  var freeformSlide = getSlide(args.slide);
                  var freeformWidth = Math.max(0.1, asFinite(args.widthMm, 100));
                  var freeformHeight = Math.max(0.1, asFinite(args.heightMm, 60));
                  var customGeometry = Api.CreateCustomGeometry();
                  for (var pathIndex = 0; pathIndex < args.paths.length; pathIndex += 1) {
                    var pathSpec = args.paths[pathIndex] || {};
                    var customPath = customGeometry.AddPath();
                    if (!customPath) throw new Error("创建自由形状路径失败");
                    customPath.SetWidth(mmToEmu(freeformWidth));
                    customPath.SetHeight(mmToEmu(freeformHeight));
                    customPath.SetStroke(pathSpec.stroke !== false);
                    customPath.SetFill(String(pathSpec.fill || "norm"));
                    var commands = Array.isArray(pathSpec.commands) ? pathSpec.commands : [];
                    for (var commandIndex = 0; commandIndex < commands.length; commandIndex += 1) {
                      var pathCommand = commands[commandIndex] || {};
                      var commandType = String(pathCommand.type || "");
                      if (commandType === "moveTo") customPath.MoveTo(mmToEmu(pathCommand.xMm), mmToEmu(pathCommand.yMm));
                      else if (commandType === "lineTo") customPath.LineTo(mmToEmu(pathCommand.xMm), mmToEmu(pathCommand.yMm));
                      else if (commandType === "quadBezTo") {
                        customPath.QuadBezTo(
                          mmToEmu(pathCommand.controlXmm),
                          mmToEmu(pathCommand.controlYmm),
                          mmToEmu(pathCommand.xMm),
                          mmToEmu(pathCommand.yMm)
                        );
                      } else if (commandType === "cubicBezTo") {
                        customPath.CubicBezTo(
                          mmToEmu(pathCommand.control1Xmm),
                          mmToEmu(pathCommand.control1Ymm),
                          mmToEmu(pathCommand.control2Xmm),
                          mmToEmu(pathCommand.control2Ymm),
                          mmToEmu(pathCommand.xMm),
                          mmToEmu(pathCommand.yMm)
                        );
                      } else if (commandType === "arcTo") {
                        customPath.ArcTo(
                          mmToEmu(pathCommand.widthRadiusMm),
                          mmToEmu(pathCommand.heightRadiusMm),
                          Math.round(asFinite(pathCommand.startAngleDeg, 0) * 60000),
                          Math.round(asFinite(pathCommand.sweepAngleDeg, 0) * 60000)
                        );
                      } else if (commandType === "close") customPath.Close();
                      else throw new Error("不支持的自由形状路径命令：" + commandType);
                    }
                  }
                  var freeform = Api.CreateShape(
                    "rect",
                    mmToEmu(freeformWidth),
                    mmToEmu(freeformHeight),
                    createFill(args.fill, args.fillColor),
                    createStroke(args.line, args.lineColor, args.lineWidthPt)
                  );
                  if (!freeform || typeof freeform.SetGeometry !== "function") throw new Error("创建自由形状失败");
                  freeform.SetGeometry(customGeometry);
                  applyDrawingFrame(freeform, {
                    xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                    yMm: hasOwn(args, "yMm") ? args.yMm : 15,
                    widthMm: freeformWidth,
                    heightMm: freeformHeight,
                  });
                  if (hasOwn(args, "name") && typeof freeform.SetName === "function") freeform.SetName(String(args.name));
                  freeformSlide.AddObject(freeform);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    object: describeDrawing(freeform, slideDrawings(freeformSlide).indexOf(freeform), Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_add_word_art": {
                  if (typeof Api.CreateWordArt !== "function") throw new Error("当前 ONLYOFFICE 版本不支持艺术字");
                  var wordArtSlide = getSlide(args.slide);
                  var wordArt = Api.CreateWordArt(
                    createWordArtTextPr(args),
                    String(args.text || ""),
                    String(args.transform || "textNoShape"),
                    createFill(args.fill, args.fillColor),
                    createStroke(args.line, args.lineColor, args.lineWidthPt),
                    asFinite(args.rotationDeg, 0),
                    mmToEmu(asFinite(args.widthMm, 100)),
                    mmToEmu(asFinite(args.heightMm, 30))
                  );
                  if (!wordArt) throw new Error("创建艺术字失败");
                  var wordArtFrame = {
                    xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                    yMm: hasOwn(args, "yMm") ? args.yMm : 15,
                    widthMm: hasOwn(args, "widthMm") ? args.widthMm : 100,
                    heightMm: hasOwn(args, "heightMm") ? args.heightMm : 30,
                    rotationDeg: hasOwn(args, "rotationDeg") ? args.rotationDeg : 0,
                  };
                  if (hasOwn(args, "name")) wordArtFrame.name = args.name;
                  applyDrawingFrame(wordArt, wordArtFrame);
                  if (wordArtSlide.AddObject(wordArt) === false) throw new Error("向幻灯片添加艺术字失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    object: describeDrawing(wordArt, slideDrawings(wordArtSlide).indexOf(wordArt), Boolean(args.includeRaw)),
                  });
                  break;
                }

                case "slides_add_math": {
                  if (typeof presentation.AddMathEquation !== "function") throw new Error("当前 ONLYOFFICE 版本不支持数学公式");
                  var mathSlide = getSlide(args.slide);
                  var drawingsBeforeMath = slideDrawings(mathSlide).slice();
                  if (typeof mathSlide.Select === "function") mathSlide.Select();
                  if (presentation.AddMathEquation(String(args.text || ""), String(args.format || "latex")) === false) {
                    throw new Error("插入数学公式失败");
                  }
                  var drawingsAfterMath = slideDrawings(mathSlide);
                  var mathDrawing = null;
                  for (var mathDrawingIndex = 0; mathDrawingIndex < drawingsAfterMath.length; mathDrawingIndex += 1) {
                    if (drawingsBeforeMath.indexOf(drawingsAfterMath[mathDrawingIndex]) === -1) {
                      mathDrawing = drawingsAfterMath[mathDrawingIndex];
                      break;
                    }
                  }
                  if (mathDrawing) applyDrawingFrame(mathDrawing, args);
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    format: String(args.format || "latex"),
                    object: mathDrawing
                      ? describeDrawing(mathDrawing, drawingsAfterMath.indexOf(mathDrawing), Boolean(args.includeRaw))
                      : null,
                  });
                  break;
                }

                case "slides_add_ole_object": {
                  if (typeof Api.CreateOleObject !== "function") throw new Error("当前 ONLYOFFICE 版本不支持 OLE 对象");
                  if (!args._image || typeof args._image.url !== "string") throw new Error("OLE 对象缺少安全导入的预览图片");
                  var oleSlide = getSlide(args.slide);
                  var oleWidthMm = asFinite(args.widthMm, 130);
                  var oleHeightMm = asFinite(args.heightMm, 90);
                  var oleObject = Api.CreateOleObject(
                    args._image.url,
                    mmToEmu(oleWidthMm),
                    mmToEmu(oleHeightMm),
                    String(args.data || ""),
                    String(args.appId || "")
                  );
                  if (!oleObject) throw new Error("创建 OLE 对象失败");
                  var oleFrame = {
                    xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                    yMm: hasOwn(args, "yMm") ? args.yMm : 15,
                    widthMm: oleWidthMm,
                    heightMm: oleHeightMm,
                  };
                  if (hasOwn(args, "rotationDeg")) oleFrame.rotationDeg = args.rotationDeg;
                  if (hasOwn(args, "name")) oleFrame.name = args.name;
                  applyDrawingFrame(oleObject, oleFrame);
                  if (oleSlide.AddObject(oleObject) === false) throw new Error("向幻灯片添加 OLE 对象失败");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    object: describeDrawing(oleObject, slideDrawings(oleSlide).indexOf(oleObject), Boolean(args.includeRaw)),
                  });
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

                case "slides_add_image": {
                  if (!args._image || !args._image.url) {
                    commandError("INVALID_IMAGE_SOURCE", "图片资源尚未安全导入");
                  }
                  if (typeof Api.CreateImage !== "function") {
                    commandError("IMAGE_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持 Api.CreateImage");
                  }
                  var imageSlideNumber = Number(args.slide);
                  var imageSlide = getSlide(imageSlideNumber);
                  var imageSize = resolveImageSize(args, 160, 90);
                  var imageDrawing = Api.CreateImage(
                    String(args._image.url),
                    mmToEmu(imageSize.widthMm),
                    mmToEmu(imageSize.heightMm)
                  );
                  if (!imageDrawing) commandError("IMAGE_FETCH_FAILED", "ONLYOFFICE 无法创建图片对象");
                  var slideWidthMm = emuToMm(safeCall(imageSlide, "GetWidth"));
                  var slideHeightMm = emuToMm(safeCall(imageSlide, "GetHeight"));
                  if (!(slideWidthMm > 0)) slideWidthMm = 254;
                  if (!(slideHeightMm > 0)) slideHeightMm = 142.875;
                  var imageFrame = {
                    xMm: hasOwn(args, "xMm") ? args.xMm : Math.max(0, (slideWidthMm - imageSize.widthMm) / 2),
                    yMm: hasOwn(args, "yMm") ? args.yMm : Math.max(0, (slideHeightMm - imageSize.heightMm) / 2),
                    widthMm: imageSize.widthMm,
                    heightMm: imageSize.heightMm,
                  };
                  if (hasOwn(args, "rotationDeg")) imageFrame.rotationDeg = args.rotationDeg;
                  if (hasOwn(args, "flipH")) imageFrame.flipH = args.flipH;
                  if (hasOwn(args, "flipV")) imageFrame.flipV = args.flipV;
                  if (hasOwn(args, "name")) imageFrame.name = args.name;
                  applyDrawingFrame(imageDrawing, imageFrame);
                  if (imageSlide.AddObject(imageDrawing) === false) throw new Error("ONLYOFFICE 拒绝向幻灯片添加图片");
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: imageSlideNumber,
                    assetId: String(args._image.assetId || ""),
                    object: describeDrawing(
                      imageDrawing,
                      slideDrawings(imageSlide).indexOf(imageDrawing),
                      false
                    ),
                  });
                  break;
                }

                case "slides_add_image_shape": {
                  if (!args._image || !args._image.url) {
                    commandError("INVALID_IMAGE_SOURCE", "图片资源尚未安全导入");
                  }
                  if (typeof Api.CreateBlipFill !== "function" || typeof Api.CreateShape !== "function") {
                    commandError("IMAGE_API_UNSUPPORTED", "当前 ONLYOFFICE 版本不支持图片形状填充");
                  }
                  var imageShapeSlide = getSlide(args.slide);
                  var imageShapeSize = resolveImageSize(args, 80, 80);
                  var imageShapeFill = Api.CreateBlipFill(
                    String(args._image.url),
                    String(args.fillMode || "stretch")
                  );
                  if (!imageShapeFill) commandError("IMAGE_FETCH_FAILED", "ONLYOFFICE 无法创建图片填充");
                  var imageShape = Api.CreateShape(
                    String(args.shapeType || "ellipse"),
                    mmToEmu(imageShapeSize.widthMm),
                    mmToEmu(imageShapeSize.heightMm),
                    imageShapeFill,
                    createStroke(args.line, args.lineColor, args.lineWidthPt)
                  );
                  if (!imageShape) commandError("IMAGE_FETCH_FAILED", "ONLYOFFICE 无法创建图片形状");
                  var imageShapeFrame = {
                    xMm: hasOwn(args, "xMm") ? args.xMm : 15,
                    yMm: hasOwn(args, "yMm") ? args.yMm : 15,
                    widthMm: imageShapeSize.widthMm,
                    heightMm: imageShapeSize.heightMm,
                  };
                  if (hasOwn(args, "rotationDeg")) imageShapeFrame.rotationDeg = args.rotationDeg;
                  if (hasOwn(args, "flipH")) imageShapeFrame.flipH = args.flipH;
                  if (hasOwn(args, "flipV")) imageShapeFrame.flipV = args.flipV;
                  if (hasOwn(args, "name")) imageShapeFrame.name = args.name;
                  applyDrawingFrame(imageShape, imageShapeFrame);
                  if (imageShapeSlide.AddObject(imageShape) === false) {
                    commandError("IMAGE_FETCH_FAILED", "向幻灯片添加图片形状失败");
                  }
                  changed += 1;
                  results.push({
                    name: call.name,
                    slide: Number(args.slide),
                    source: {
                      assetId: args._image.assetId || null,
                      widthPx: Number(args._image.widthPx),
                      heightPx: Number(args._image.heightPx),
                    },
                    object: describeDrawing(
                      imageShape,
                      slideDrawings(imageShapeSlide).indexOf(imageShape),
                      Boolean(args.includeRaw)
                    ),
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
            return JSON.stringify({
              ok: false,
              code: error && error.code ? error.code : undefined,
              error: error && error.message ? error.message : String(error),
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

  function executePluginTool(toolCalls) {
    return new Promise(function (resolve, reject) {
      if (toolCalls.length !== 1) {
        reject(new Error("主题库、宏和放映控制工具必须单独调用"));
        return;
      }
      if (!window.Asc || !Asc.plugin || typeof Asc.plugin.executeMethod !== "function") {
        reject(new Error("当前 ONLYOFFICE 版本不支持所需的插件方法"));
        return;
      }
      var call = toolCalls[0] || {};
      var args = call.arguments || call.args || {};
      var method;
      var params = null;
      var needsSave = false;
      var resultKind = null;
      if (call.name === "slides_inspect_builtin_themes") {
        method = "GetEditorThemes";
        resultKind = "themes";
      } else if (call.name === "slides_apply_builtin_theme") {
        method = "ApplyTheme";
        params = [args.theme];
        needsSave = true;
        resultKind = "theme";
      } else if (call.name === "slides_inspect_macros") {
        method = args.kind === "vba" ? "GetVBAMacros" : "GetMacros";
        resultKind = method === "GetVBAMacros" ? "vba" : "onlyoffice";
      } else if (call.name === "slides_set_macros") {
        if (!args.content || typeof args.content !== "object" || Array.isArray(args.content)) {
          reject(new Error("slides_set_macros.content 必须是宏配置对象"));
          return;
        }
        method = "SetMacros";
        params = [JSON.stringify(args.content)];
        needsSave = true;
        resultKind = "onlyoffice";
      } else if (call.name === "slides_control_slideshow") {
        var slideshowMethods = {
          start: "StartSlideShow",
          end: "EndSlideShow",
          pause: "PauseSlideShow",
          resume: "ResumeSlideShow",
          next: "GoToNextSlideInSlideShow",
          previous: "GoToPreviousSlideInSlideShow",
          goto: "GoToSlideInSlideShow",
        };
        method = slideshowMethods[String(args.action || "")];
        if (!method) {
          reject(new Error("不支持的放映控制动作：" + String(args.action || "")));
          return;
        }
        if (args.action === "goto") {
          var slideIndex = Number(args.slide);
          if (!Number.isInteger(slideIndex) || slideIndex < 1) {
            reject(new Error("goto 需要从 1 开始的 slide"));
            return;
          }
          params = [slideIndex - 1];
        }
        resultKind = "slideshow";
      } else {
        reject(new Error("未知插件工具：" + call.name));
        return;
      }

      Asc.plugin.executeMethod(method, params, function (data) {
        try {
          var content = data;
          if ((method === "GetMacros" || method === "GetEditorThemes") && typeof data === "string") {
            try {
              content = JSON.parse(data);
            } catch (error) {
              content = { raw: data };
            }
          }
          resolve({
            ok: true,
            editorType: "slide",
            changed: needsSave ? 1 : 0,
            needsSave: needsSave,
            results: [{
              name: call.name,
              kind: resultKind,
              action: call.name === "slides_control_slideshow" ? args.action : undefined,
              content: content,
            }],
          });
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  function isPluginTool(call) {
    return Boolean(call) && (
      call.name === "slides_inspect_builtin_themes"
      || call.name === "slides_apply_builtin_theme"
      || call.name === "slides_inspect_macros"
      || call.name === "slides_set_macros"
      || call.name === "slides_control_slideshow"
    );
  }

  function execute(toolCalls) {
    requireStaticValidation(toolCalls);
    if (toolCalls.some(isPluginTool)) return executePluginTool(toolCalls);
    return executeOfficeCommands(toolCalls);
  }

  window.AICopilotBridges.slide = {
    execute: execute,
    inspect: function () {
      return execute([{ name: "slides_inspect", arguments: { maxChars: 12000 } }]);
    },
  };
})();
