/**
 * DOCX Parser — разбирает .docx файл и извлекает данные для проверки по СТП
 * Использует JSZip для разархивирования и DOMParser для парсинга XML
 */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ============================================================
// XML Helper — удобная работа с OOXML namespace
// ============================================================
const XML = {
  parse(str) {
    const p = new DOMParser();
    const doc = p.parseFromString(str, 'application/xml');
    // Проверка на ошибки парсинга
    const err = doc.querySelector('parsererror');
    if (err) throw new Error('XML parse error: ' + err.textContent.slice(0, 100));
    return doc;
  },

  // Получить элементы по имени тега с namespace w:
  els(parent, tag) {
    if (!parent) return [];
    const byNS = parent.getElementsByTagNameNS(W_NS, tag);
    if (byNS.length > 0) return Array.from(byNS);
    // Fallback для файлов без proper namespace
    return Array.from(parent.getElementsByTagName('w:' + tag));
  },

  el(parent, tag) { return XML.els(parent, tag)[0] || null; },

  // Получить атрибут w:val или просто по имени
  attr(el, name) {
    if (!el) return null;
    return el.getAttributeNS(W_NS, name) ||
           el.getAttribute('w:' + name) ||
           el.getAttribute(name) ||
           null;
  },

  // Получить весь текст из элемента (все w:t)
  text(parent) {
    if (!parent) return '';
    const tEls = XML.els(parent, 't');
    return tEls.map(t => t.textContent || '').join('');
  }
};

// ============================================================
// DocxParser — основной класс
// ============================================================
class DocxParser {

  async parse(file) {
    if (!window.JSZip) throw new Error('JSZip не загружен');

    const zip = await JSZip.loadAsync(file);

    // Читаем нужные XML файлы
    const [docXmlStr, stylesXmlStr, numXmlStr] = await Promise.all([
      zip.file('word/document.xml')?.async('string'),
      zip.file('word/styles.xml')?.async('string'),
      zip.file('word/numbering.xml')?.async('string')
    ]);

    if (!docXmlStr) throw new Error('Файл не является корректным .docx документом (word/document.xml не найден)');

    let docXml, stylesXml;
    try {
      docXml = XML.parse(docXmlStr);
      stylesXml = stylesXmlStr ? XML.parse(stylesXmlStr) : null;
    } catch (e) {
      throw new Error('Ошибка парсинга XML: ' + e.message);
    }

    // Извлекаем данные
    const pageSettings = this._extractPageSettings(docXml);
    const styles = stylesXml ? this._extractStyles(stylesXml) : {};
    const defaultStyle = stylesXml ? this._extractDefaultStyle(stylesXml) : null;

    const paragraphs = this._extractParagraphs(docXml, styles, defaultStyle);
    const tables = this._extractTables(docXml, styles, defaultStyle);
    const footerData = await this._extractFooters(zip);
    const headerData = await this._extractHeaders(zip);

    // Получаем HTML-представление через mammoth (если доступен)
    let htmlContent = null;
    if (window.mammoth) {
      try {
        const ab = await file.arrayBuffer();
        const mammothOptions = {
          arrayBuffer: ab,
          // Конвертируем изображения в data-URI для отображения в браузере
          convertImage: mammoth.images.imgElement(function(image) {
            return image.read('base64').then(function(base64) {
              return { src: 'data:' + image.contentType + ';base64,' + base64 };
            });
          }),
          styleMap: [
            "p[style-name='Normal'] => p:fresh",
            "p[style-name='Heading 1'] => h1:fresh",
            "p[style-name='Heading 2'] => h2:fresh",
            "p[style-name='Heading 3'] => h3:fresh",
            "p[style-name='Heading 4'] => h4:fresh",
            "p[style-name='Заголовок 1'] => h1:fresh",
            "p[style-name='Заголовок 2'] => h2:fresh",
            "p[style-name='Заголовок 3'] => h3:fresh",
          ]
        };
        const result = await mammoth.convertToHtml(mammothOptions);
        htmlContent = result.value;
        if (result.messages && result.messages.length > 0) {
          console.debug('mammoth messages:', result.messages.slice(0, 5));
        }
      } catch (e) {
        console.warn('mammoth error:', e);
      }
    }

    // Получаем полный текст для поиска паттернов
    const fullText = paragraphs.map(p => p.text).join('\n');

    // Оцениваем номера страниц для каждого параграфа
    this._estimatePageNumbers(paragraphs);

    return {
      pageSettings,
      styles,
      defaultStyle,
      paragraphs,
      tables,
      footerData,
      headerData,
      htmlContent,
      fullText,
      fileName: file.name
    };
  }

  // -----------------------------------------------------------
  // Оценка номера страницы для каждого параграфа
  // Учитывает: явные разрывы страниц, рисунки, реальный межстрочный интервал
  _estimatePageNumbers(paragraphs) {
    // A4 с полями 30/15/20/20 мм → высота печатной области: 297-20-20=257мм
    // 1pt = 0.353мм → 257мм / 0.353 = ~728pt
    const PAGE_HEIGHT_PT = 728;
    const CHARS_PER_LINE = 68;    // ~68 символов в строке для Times New Roman 14pt при ширине 165мм

    const twipToPt = (twips) => (twips || 0) / 20;

    let heightPt = 0;
    let pageNum  = 1;

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      para.paraIndex     = i;
      para.estimatedPage = pageNum;

      // 1. Явный разрыв страницы ДО параграфа (pageBreakBefore или w:br type=page)
      if (para.hasPageBreak && i > 0) {
        pageNum++;
        heightPt = 0;
        para.estimatedPage = pageNum;
      }

      // 2. Высота параграфа в pt
      let paraPt = 0;

      if (para.hasDrawing) {
        // Рисунок: используем реальные размеры из XML или fallback
        paraPt = (para.drawingHeightPt > 0 ? para.drawingHeightPt : 150) + 20;
      } else if (para.isEmpty) {
        paraPt = 14;
      } else {
        const sp = para.pPr?.spacing;

        // Межстрочный интервал
        let lineHeightPt = 14; // single spacing 14pt
        if (sp && sp.line > 0) {
          lineHeightPt = sp.lineRule === 'exact'
            ? twipToPt(sp.line)
            : (sp.line / 240) * 14;
        }
        // Гарантируем минимум (не меньше размера шрифта)
        lineHeightPt = Math.max(lineHeightPt, 14);

        const textLines = Math.max(1, Math.ceil(para.text.length / CHARS_PER_LINE));
        const spaceBefore = sp ? twipToPt(sp.before) : 0;
        const spaceAfter  = sp ? twipToPt(sp.after)  : 8; // Word default ~8pt aftertext

        paraPt = textLines * lineHeightPt + spaceBefore + spaceAfter;

        // Для заголовков добавляем дополнительный отступ если не задан явно
        if (para.isHeading) {
          paraPt += (spaceBefore < 6 ? 10 : 0) + (spaceAfter < 4 ? 6 : 0);
        }
      }

      heightPt += paraPt;

      // 3. Переход на следующую страницу
      if (heightPt >= PAGE_HEIGHT_PT) {
        const overflow = heightPt - PAGE_HEIGHT_PT;
        pageNum++;
        // Если параграф сам больше страницы — просто обнуляем
        heightPt = paraPt > PAGE_HEIGHT_PT ? 0 : overflow;
      }
    }
  }

  // -----------------------------------------------------------
  _extractPageSettings(docXml) {
    const body = XML.el(docXml, 'body');
    const sectPrs = XML.els(docXml, 'sectPr');
    if (!sectPrs.length) return null;

    // Берём последний sectPr (документ-уровень)
    const sectPr = sectPrs[sectPrs.length - 1];
    const pgMar = XML.el(sectPr, 'pgMar');
    const pgSz = XML.el(sectPr, 'pgSz');
    const pgNumType = XML.el(sectPr, 'pgNumType');

    const settings = {
      margins: null,
      pageSize: null,
      pageNumberStart: null,
    };

    if (pgMar) {
      settings.margins = {
        left:   parseInt(XML.attr(pgMar, 'left')   || '0'),
        right:  parseInt(XML.attr(pgMar, 'right')  || '0'),
        top:    parseInt(XML.attr(pgMar, 'top')    || '0'),
        bottom: parseInt(XML.attr(pgMar, 'bottom') || '0'),
        header: parseInt(XML.attr(pgMar, 'header') || '0'),
        footer: parseInt(XML.attr(pgMar, 'footer') || '0'),
      };
    }

    if (pgSz) {
      settings.pageSize = {
        w: parseInt(XML.attr(pgSz, 'w') || '0'),
        h: parseInt(XML.attr(pgSz, 'h') || '0'),
        orient: XML.attr(pgSz, 'orient') || 'portrait',
      };
    }

    if (pgNumType) {
      settings.pageNumberStart = parseInt(XML.attr(pgNumType, 'start') || '1');
    }

    return settings;
  }

  // -----------------------------------------------------------
  _extractStyles(stylesXml) {
    const styles = {};
    const styleEls = XML.els(stylesXml, 'style');

    for (const styleEl of styleEls) {
      const id = XML.attr(styleEl, 'styleId');
      if (!id) continue;

      const nameEl = XML.el(styleEl, 'name');
      const basedOnEl = XML.el(styleEl, 'basedOn');
      const pPrEl = XML.el(styleEl, 'pPr');
      const rPrEl = XML.el(styleEl, 'rPr');

      styles[id] = {
        id,
        type: XML.attr(styleEl, 'type'),
        name: (XML.attr(nameEl, 'val') || '').toLowerCase(),
        basedOn: XML.attr(basedOnEl, 'val'),
        pPr: pPrEl ? this._parsePPr(pPrEl) : null,
        rPr: rPrEl ? this._parseRPr(rPrEl) : null,
      };
    }
    return styles;
  }

  // -----------------------------------------------------------
  _extractDefaultStyle(stylesXml) {
    const docDefaults = XML.el(stylesXml, 'docDefaults');
    if (!docDefaults) return null;

    const rPrDefaultEl = XML.el(docDefaults, 'rPrDefault');
    const pPrDefaultEl = XML.el(docDefaults, 'pPrDefault');

    return {
      rPr: rPrDefaultEl ? this._parseRPr(XML.el(rPrDefaultEl, 'rPr')) : null,
      pPr: pPrDefaultEl ? this._parsePPr(XML.el(pPrDefaultEl, 'pPr')) : null,
    };
  }

  // -----------------------------------------------------------
  _parsePPr(pPr) {
    if (!pPr) return null;

    const jcEl = XML.el(pPr, 'jc');
    const spacingEl = XML.el(pPr, 'spacing');
    const indEl = XML.el(pPr, 'ind');
    const pStyleEl = XML.el(pPr, 'pStyle');
    const outlineEl = XML.el(pPr, 'outlineLvl');
    const numPrEl = XML.el(pPr, 'numPr');

    return {
      styleId: XML.attr(pStyleEl, 'val'),
      alignment: XML.attr(jcEl, 'val'),   // 'both', 'center', 'left', 'right'
      spacing: spacingEl ? {
        before:    parseInt(XML.attr(spacingEl, 'before')   || '0'),
        after:     parseInt(XML.attr(spacingEl, 'after')    || '0'),
        line:      parseInt(XML.attr(spacingEl, 'line')     || '0'),
        lineRule:  XML.attr(spacingEl, 'lineRule') || 'auto',
      } : null,
      indent: indEl ? {
        left:      parseInt(XML.attr(indEl, 'left')      || '0'),
        right:     parseInt(XML.attr(indEl, 'right')     || '0'),
        firstLine: parseInt(XML.attr(indEl, 'firstLine') || '0'),
        hanging:   parseInt(XML.attr(indEl, 'hanging')   || '0'),
      } : null,
      outlineLevel: outlineEl ? parseInt(XML.attr(outlineEl, 'val') || '9') : null,
      isList: !!numPrEl,
      pageBreakBefore: XML.el(pPr, 'pageBreakBefore') ? XML.attr(XML.el(pPr, 'pageBreakBefore'), 'val') !== 'false' : false,
    };
  }

  // -----------------------------------------------------------
  // _parseRPr — возвращает SPARSE объект: только явно заданные свойства.
  // Отсутствие ключа означает «не задано — наследуется от родительского стиля».
  _parseRPr(rPr) {
    if (!rPr) return {};

    const result = {};

    // null если элемент отсутствует (= наследовать), true/false если явно задан
    const boolOrNull = (el) => {
      if (!el) return null;
      const v = XML.attr(el, 'val');
      return v !== 'false' && v !== '0';
    };

    const rFontsEl = XML.el(rPr, 'rFonts');
    if (rFontsEl) {
      const font = XML.attr(rFontsEl, 'ascii') || XML.attr(rFontsEl, 'hAnsi') || null;
      if (font) result.font = font;
      const theme = XML.attr(rFontsEl, 'asciiTheme');
      if (theme) result.fontTheme = theme;
    }

    // sz — размер шрифта (в полупунктах), szCs — для кириллицы; берём любой
    const szEl  = XML.el(rPr, 'sz') || XML.el(rPr, 'szCs');
    if (szEl) result.size = parseInt(XML.attr(szEl, 'val') || '0') / 2;

    const bold = boolOrNull(XML.el(rPr, 'b'));
    if (bold !== null) result.bold = bold;

    const italic = boolOrNull(XML.el(rPr, 'i'));
    if (italic !== null) result.italic = italic;

    const caps = boolOrNull(XML.el(rPr, 'caps'));
    if (caps !== null) result.caps = caps;

    const smCaps = boolOrNull(XML.el(rPr, 'smallCaps'));
    if (smCaps !== null) result.smallCaps = smCaps;

    const colorEl = XML.el(rPr, 'color');
    if (colorEl) result.color = XML.attr(colorEl, 'val');

    const uEl = XML.el(rPr, 'u');
    if (uEl) result.underline = XML.attr(uEl, 'val') !== 'none';

    return result;
  }

  // -----------------------------------------------------------
  // Слияние двух sparse-rPr объектов. override побеждает для всех определённых значений.
  _mergeRPr(base, override) {
    const result = Object.assign({}, base);
    if (!override) return result;
    for (const [k, v] of Object.entries(override)) {
      if (v !== undefined && v !== null) result[k] = v;
    }
    return result;
  }

  // -----------------------------------------------------------
  // Слияние двух pPr объектов. Для spacing/indent выполняется глубокое слияние.
  _mergePPr(base, override) {
    const result = Object.assign({}, base);
    if (!override) return result;
    for (const [k, v] of Object.entries(override)) {
      if (v === undefined || v === null) continue;
      if ((k === 'spacing' || k === 'indent') &&
          typeof v === 'object' && result[k] && typeof result[k] === 'object') {
        result[k] = Object.assign({}, result[k]);
        for (const [sk, sv] of Object.entries(v)) {
          if (sv !== undefined && sv !== null) result[k][sk] = sv;
        }
      } else {
        result[k] = v;
      }
    }
    return result;
  }

  // -----------------------------------------------------------
  // Разрешает полную цепочку наследования стилей (basedOn) для styleId.
  // Порядок: docDefaults → корневой стиль → ... → прямой стиль.
  // Возвращает {rPr, pPr} — итоговые эффективные свойства стиля.
  _resolveStyleChain(styleId, styles, defaultStyle) {
    const chain = [];
    let current = styleId;
    const visited = new Set();
    while (current && !visited.has(current)) {
      visited.add(current);
      const s = styles[current];
      if (!s) break;
      chain.unshift({ rPr: s.rPr || {}, pPr: s.pPr || {} }); // prepend → корень первый
      current = s.basedOn;
    }

    let rPr = Object.assign({}, defaultStyle?.rPr || {});
    let pPr = Object.assign({}, defaultStyle?.pPr || {});

    for (const { rPr: sr, pPr: sp } of chain) {
      rPr = this._mergeRPr(rPr, sr);
      pPr = this._mergePPr(pPr, sp);
    }

    return { rPr, pPr };
  }

  // -----------------------------------------------------------
  _extractParagraphs(docXml, styles, defaultStyle) {
    const body = XML.el(docXml, 'body');
    if (!body) return [];

    const paragraphs = [];
    this._traverseParagraphs(body, paragraphs, styles, defaultStyle, false);
    return paragraphs;
  }

  _traverseParagraphs(node, result, styles, defaultStyle, tableInfo = null) {
    let rowIndex = -1;
    let cellIndex = -1;

    for (const child of node.childNodes) {
      const localName = child.localName || child.nodeName.replace('w:', '');

      if (localName === 'p') {
        const para = this._parseParagraph(child, styles, defaultStyle, tableInfo);
        result.push(para);
      } else if (localName === 'tbl') {
        // Инициализируем инфо о таблице
        const thisTblIdx = (tableInfo ? tableInfo.tableIndex + 1 : 0); // Упрощенно, вложенные таблицы редки
        this._traverseParagraphs(child, result, styles, defaultStyle, { tableIndex: thisTblIdx, rowIndex: -1, cellIndex: -1 });
      } else if (localName === 'tr') {
        rowIndex++;
        cellIndex = -1;
        this._traverseParagraphs(child, result, styles, defaultStyle, { ...tableInfo, rowIndex });
      } else if (localName === 'tc') {
        cellIndex++;
        this._traverseParagraphs(child, result, styles, defaultStyle, { ...tableInfo, rowIndex, cellIndex });
      } else if (localName === 'sdt' || localName === 'body') {
        this._traverseParagraphs(child, result, styles, defaultStyle, tableInfo);
      }
    }
  }

  _parseParagraph(pNode, styles, defaultStyle, tableInfo) {
    const pPrEl = XML.el(pNode, 'pPr');
    const pPr = pPrEl ? this._parsePPr(pPrEl) : null;

    // Прямое форматирование параграфа
    const pRPrEl = pPrEl ? XML.el(pPrEl, 'rPr') : null;

    // Собираем все runs
    const runs = [];
    const rEls = XML.els(pNode, 'r');
    for (const rEl of rEls) {
      const rPrEl = XML.el(rEl, 'rPr');
      const text = XML.text(rEl);
      if (text !== '') {
        runs.push({
          text,
          rPr: rPrEl ? this._parseRPr(rPrEl) : null,
        });
      }
    }

    // Поля инструкции (например, PAGE для номеров страниц)
    const instrTexts = XML.els(pNode, 'instrText').map(t => t.textContent);
    const isPageNumberField = instrTexts.some(t => t.trim().startsWith('PAGE') || t.includes('NUMPAGES'));

    // Явные разрывы страниц внутри параграфа (w:br w:type="page")
    const brEls = XML.els(pNode, 'br');
    const hasExplicitPageBreak = brEls.some(br => XML.attr(br, 'type') === 'page');

    // Изображения/рисунки — ищем w:drawing (содержит размеры в EMU)
    const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    const drawingEls = Array.from(pNode.getElementsByTagNameNS
      ? pNode.getElementsByTagNameNS(W_NS, 'drawing')
      : pNode.getElementsByTagName('w:drawing'));

    let drawingHeightPt = 0;
    for (const dr of drawingEls) {
      const extEls = dr.getElementsByTagNameNS(WP_NS, 'extent');
      const extEl  = extEls.length ? extEls[0] : null;
      const cy     = extEl ? parseInt(extEl.getAttribute('cy') || '0') : 0;
      drawingHeightPt += cy > 0 ? cy / 12700 : 120; // 12700 EMU = 1pt; fallback ~120pt
    }
    const hasDrawing = drawingEls.length > 0;

    const fullText = runs.map(r => r.text).join('');

    // Определяем стиль параграфа
    const styleId = pPr?.styleId;
    const style = styleId ? (styles[styleId] || null) : null;

    // Определяем уровень заголовка
    let headingLevel = 0;
    let isHeading = false;

    // Из стиля
    if (style) {
      const sName = style.name || '';
      const headingMatch = sName.match(/^(heading|заголовок)\s*(\d+)$/i);
      if (headingMatch) {
        isHeading = true;
        headingLevel = parseInt(headingMatch[2]);
      }
      if (style.pPr?.outlineLevel !== null && style.pPr?.outlineLevel !== undefined && style.pPr.outlineLevel < 9) {
        isHeading = true;
        headingLevel = style.pPr.outlineLevel + 1;
      }
    }

    // Из outlineLevel параграфа
    if (pPr?.outlineLevel !== null && pPr?.outlineLevel !== undefined && pPr.outlineLevel < 9) {
      isHeading = true;
      headingLevel = pPr.outlineLevel + 1;
    }

    // Разрешаем эффективный rPr (из run или pPr rPr)
    const firstRunRPr = runs.length > 0 ? runs[0].rPr : null;

    // -------------------------------------------------------
    // Полное разрешение через цепочку наследования стилей:
    //   docDefaults → basedOn-цепочка → прямой стиль → pPr/rPr → run-rPr
    // Это единственный надёжный способ узнать реальный шрифт/размер/жирность.
    const { rPr: styleRPr, pPr: stylePPr } = this._resolveStyleChain(styleId, styles, defaultStyle);

    // effectiveRPr для первого run: стиль + paragraph-rPr + run-rPr
    const effectiveRPr = this._mergeRPr(
      this._mergeRPr(styleRPr, pRPrEl ? this._parseRPr(pRPrEl) : {}),
      firstRunRPr || {}
    );

    // effectivePPr: стиль + paragraph-pPr
    const effectivePPr = this._mergePPr(stylePPr, pPr || {});

    return {
      text: fullText,
      textTrimmed: fullText.trim(),
      runs,
      pPr,
      pRPr: pRPrEl ? this._parseRPr(pRPrEl) : null,
      firstRunRPr,
      effectiveRPr,   // полностью разрешённые свойства символов ← используй это
      effectivePPr,   // полностью разрешённые свойства абзаца  ← используй это
      styleId,
      style,
      isHeading,
      headingLevel,
      inTable: !!tableInfo,
      tableIndex: tableInfo ? tableInfo.tableIndex : -1,
      rowIndex: tableInfo ? tableInfo.rowIndex : -1,
      cellIndex: tableInfo ? tableInfo.cellIndex : -1,
      isEmpty: fullText.trim() === '',
      isPageNumberField,
      hasPageBreak: hasExplicitPageBreak || (pPr?.pageBreakBefore === true),
      hasDrawing,
      drawingHeightPt,
    };
  }

  // -----------------------------------------------------------
  _extractTables(docXml, styles, defaultStyle) {
    const tables = [];
    const tblEls = XML.els(docXml, 'tbl');

    for (let i = 0; i < tblEls.length; i++) {
      const tbl = tblEls[i];
      const rows = XML.els(tbl, 'tr');
      const cells = XML.els(tbl, 'tc');

      // Ищем заголовок таблицы — параграф перед таблицей
      tables.push({
        index: i,
        rowCount: rows.length,
        cellCount: cells.length,
      });
    }
    return tables;
  }

  // -----------------------------------------------------------
  async _extractFooters(zip) {
    const result = {
      hasFooter: false,
      hasPageNumber: false,
      pageNumAlignment: null,
      pageNumRight: false,
    };

    // Ищем файлы футера
    for (let i = 1; i <= 5; i++) {
      const content = await zip.file(`word/footer${i}.xml`)?.async('string');
      if (!content) continue;

      result.hasFooter = true;

      // Проверяем наличие номера страницы
      if (content.includes('>PAGE<') || content.includes('PAGE') || content.includes('page')) {
        result.hasPageNumber = true;

        // Проверяем выравнивание
        const xmlDoc = XML.parse(content);
        const jcEls = XML.els(xmlDoc, 'jc');
        for (const jc of jcEls) {
          const val = XML.attr(jc, 'val');
          if (val) {
            result.pageNumAlignment = val;
            result.pageNumRight = (val === 'right');
          }
        }
      }
      break;
    }

    return result;
  }

  // -----------------------------------------------------------
  async _extractHeaders(zip) {
    const result = { hasHeader: false, hasPageNumber: false };

    for (let i = 1; i <= 5; i++) {
      const content = await zip.file(`word/header${i}.xml`)?.async('string');
      if (!content) continue;
      result.hasHeader = true;
      if (content.includes('PAGE')) result.hasPageNumber = true;
      break;
    }

    return result;
  }

  // -----------------------------------------------------------
  // Вспомогательный метод: twips → мм
  static twipsToMm(twips) {
    return Math.round((twips / 1440) * 25.4 * 10) / 10;
  }

  // -----------------------------------------------------------
  // Вспомогательный метод: разрешить эффективный стиль с учётом наследования
  static resolveStyle(styleId, styles, defaultStyle) {
    if (!styleId || !styles) return defaultStyle;

    const chain = [];
    let current = styles[styleId];

    // Строим цепочку наследования (до 10 уровней)
    let depth = 0;
    while (current && depth < 10) {
      chain.unshift(current);
      current = current.basedOn ? styles[current.basedOn] : null;
      depth++;
    }

    if (defaultStyle) chain.unshift({ pPr: defaultStyle.pPr, rPr: defaultStyle.rPr });

    // Объединяем стили (последний имеет приоритет)
    const merged = { pPr: null, rPr: null };
    for (const s of chain) {
      if (s.pPr) merged.pPr = { ...(merged.pPr || {}), ...s.pPr };
      if (s.rPr) merged.rPr = { ...(merged.rPr || {}), ...s.rPr };
    }

    return merged;
  }
}
