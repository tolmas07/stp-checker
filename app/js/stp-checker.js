/**
 * STP Checker — проверяет данные документа по правилам СТП 01-2024 БГУИР
 * Возвращает массив объектов Finding
 */

class StpChecker {

  constructor() {
    this.findings = [];
  }

  // ============================================================
  // Главный метод проверки
  // ============================================================
  check(docData, options = {}) {
    this.findings = [];
    this.docData = docData;

    const opts = {
      checkFormat: true,
      checkStructure: true,
      checkHeadings: true,
      checkFigures: true,
      checkFormulas: true,
      checkRefs: true,
      checkText: true,
      ...options
    };

    if (opts.checkFormat)    this._checkPageSettings();
    if (opts.checkFormat)    this._checkDefaultFormatting();
    if (opts.checkStructure) this._checkDocumentStructure();
    if (opts.checkStructure) this._checkDocumentOrder();
    if (opts.checkStructure) this._checkAbstract();
    if (opts.checkHeadings)  this._checkHeadings();
    if (opts.checkHeadings)  this._checkHeadingFormatting();
    if (opts.checkHeadings)  this._checkTableOfContents();
    if (opts.checkFigures)   this._checkFigureCaptions();
    if (opts.checkFigures)   this._checkTableCaptions();
    if (opts.checkFigures)   this._checkFigureTableSequencing();
    if (opts.checkFormulas)  this._checkFormulas();
    if (opts.checkRefs)      this._checkReferences();
    if (opts.checkText)      this._checkTextRules();
    if (opts.checkFigures)   this._checkAppendices();
    this._checkPageNumbering();

    return this.findings;
  }

  // ============================================================
  // ФАБРИКА НАХОДОК
  // ============================================================
  // pageNum — примерная страница (undefined = неизвестно)
  _add(severity, id, section, description, location, recommendation, pageNum) {
    const locWithPage = pageNum
      ? `${location ? location + ' · ' : ''}~стр. ${pageNum}`
      : location;
    this.findings.push({ severity, id, section, description, location: locWithPage, pageNum, recommendation, source: 'programmatic' });
  }

  _pass(id, section, description) {
    this.findings.push({ severity: 'pass', id, section, description, location: null, recommendation: null, source: 'programmatic' });
  }

  // Хелпер: получить номер страницы параграфа
  _page(para) { return para?.estimatedPage; }

  // Конвертация twips → мм
  _mm(twips) { return DocxParser.twipsToMm(twips); }
  _mmTolerance(actual, expected, tol = 2) { return Math.abs(actual - expected) <= tol; }

  // ============================================================
  // 1. ПОЛЯ И РАЗМЕР СТРАНИЦЫ (п. 2.1.1)
  // ============================================================
  _checkPageSettings() {
    const ps = this.docData.pageSettings;

    if (!ps) {
      this._add('warning', 'no-page-settings', 'п. 2.1.1',
        'Не удалось определить параметры страницы',
        'Параметры страницы',
        'Проверьте параметры страницы вручную в документе');
      return;
    }

    // Размер страницы — А4 (210×297 мм)
    if (ps.pageSize) {
      const wMm = this._mm(ps.pageSize.w);
      const hMm = this._mm(ps.pageSize.h);
      const isA4 = (Math.abs(wMm - 210) <= 5 && Math.abs(hMm - 297) <= 5) ||
                   (Math.abs(wMm - 297) <= 5 && Math.abs(hMm - 210) <= 5);
      if (!isA4) {
        this._add('critical', 'page-size', 'п. 2.1.1',
          `Размер страницы: ${wMm}×${hMm} мм. Требуется А4 (210×297 мм)`,
          'Параметры страницы',
          'Макет → Размер → А4 (210×297 мм)');
      } else {
        this._pass('page-size-ok', 'п. 2.1.1', `Размер страницы А4 (${wMm}×${hMm} мм) ✓`);
      }
      if (ps.pageSize.orient === 'landscape') {
        this._add('critical', 'page-orient', 'п. 2.1.1',
          'Документ в альбомной ориентации. Требуется книжная.',
          'Параметры страницы',
          'Установите книжную ориентацию: Макет → Ориентация → Книжная');
      }
    }

    // Поля страницы
    if (ps.margins) {
      const { left, right, top, bottom } = ps.margins;
      const lMm = this._mm(left);
      const rMm = this._mm(right);
      const tMm = this._mm(top);
      const bMm = this._mm(bottom);

      const tol = 2;
      let marginsOk = true;

      if (!this._mmTolerance(lMm, 30, tol)) {
        marginsOk = false;
        this._add('critical', 'margin-left', 'п. 2.1.1',
          `Левое поле: ${lMm} мм (требуется 30 мм)`,
          'Поля страницы',
          'Установите левое поле = 30 мм: Макет → Поля → Настраиваемые поля');
      }
      if (!this._mmTolerance(rMm, 15, tol)) {
        marginsOk = false;
        this._add('critical', 'margin-right', 'п. 2.1.1',
          `Правое поле: ${rMm} мм (требуется 15 мм)`,
          'Поля страницы',
          'Установите правое поле = 15 мм');
      }
      if (!this._mmTolerance(tMm, 20, tol)) {
        marginsOk = false;
        this._add('critical', 'margin-top', 'п. 2.1.1',
          `Верхнее поле: ${tMm} мм (требуется 20 мм)`,
          'Поля страницы',
          'Установите верхнее поле = 20 мм');
      }
      if (!this._mmTolerance(bMm, 20, tol)) {
        marginsOk = false;
        this._add('critical', 'margin-bottom', 'п. 2.1.1',
          `Нижнее поле: ${bMm} мм (требуется 20 мм)`,
          'Поля страницы',
          'Установите нижнее поле = 20 мм');
      }
      if (marginsOk) {
        this._pass('margins-ok', 'п. 2.1.1',
          `Поля страницы верные (л=${lMm}, п=${rMm}, в=${tMm}, н=${bMm} мм) ✓`);
      }
    }
  }

  // ============================================================
  // 2. ШРИФТ, ИНТЕРВАЛ, ОТСТУП (п. 2.1.1)
  // ============================================================
  _checkDefaultFormatting() {
    const { styles, defaultStyle } = this.docData;

    // Определяем «нормальный» стиль
    const normalIds = ['Normal', 'Обычный', 'normal', 'Default Paragraph Font'];
    let normalStyle = null;
    for (const id of normalIds) {
      if (styles[id]) { normalStyle = styles[id]; break; }
    }
    // Если не нашли по ID, ищем по имени
    if (!normalStyle) {
      normalStyle = Object.values(styles).find(s => s.name === 'normal' || s.name === 'обычный');
    }

    const rPr = normalStyle?.rPr || defaultStyle?.rPr;
    const pPr = normalStyle?.pPr || defaultStyle?.pPr;

    // --- Шрифт ---
    if (rPr?.font) {
      if (!rPr.font.toLowerCase().includes('times new roman')) {
        this._add('critical', 'font-name', 'п. 2.1.1',
          `Основной шрифт: «${rPr.font}». Требуется Times New Roman`,
          'Стиль «Обычный»',
          'Выделите весь текст (Ctrl+A) → установите Times New Roman');
      } else {
        this._pass('font-ok', 'п. 2.1.1', `Шрифт основного текста: Times New Roman ✓`);
      }
    } else if (rPr?.fontTheme) {
      // Тема шрифта — обычно не Times New Roman
      this._add('warning', 'font-theme', 'п. 2.1.1',
        'Шрифт задан через тему документа. Убедитесь, что это Times New Roman',
        'Стиль «Обычный»',
        'Явно укажите шрифт Times New Roman вместо темы');
    } else {
      this._add('info', 'font-not-detected', 'п. 2.1.1',
        'Не удалось определить шрифт основного текста из стилей',
        'Стиль «Обычный»',
        'Убедитесь, что основной текст набран шрифтом Times New Roman 14пт');
    }

    // --- Размер шрифта ---
    if (rPr?.size) {
      if (Math.abs(rPr.size - 14) > 0.5) {
        this._add('critical', 'font-size', 'п. 2.1.1',
          `Размер шрифта: ${rPr.size} пт. Требуется 14 пт`,
          'Стиль «Обычный»',
          'Установите размер шрифта 14 пт для всего основного текста');
      } else {
        this._pass('font-size-ok', 'п. 2.1.1', `Размер шрифта: 14 пт ✓`);
      }
    }

    // --- Межстрочный интервал ---
    // СТП: 1,0 (18 пунктов) — одинарный интервал
    // Частая ошибка: 1,5 = w:line=360, lineRule=auto
    if (pPr?.spacing) {
      const { line, lineRule } = pPr.spacing;
      if (line > 0) {
        const isOneHalf = (lineRule === 'auto' || !lineRule) && line >= 340 && line <= 380;
        const isDouble  = (lineRule === 'auto' || !lineRule) && line >= 450;
        if (isDouble) {
          this._add('critical', 'spacing-double', 'п. 2.1.1',
            `Межстрочный интервал «двойной» (line=${line}). Требуется одинарный (1,0 = 18 пт)`,
            'Стиль «Обычный»',
            'Установите межстрочный интервал: Одинарный или «Точно» 18 пт');
        } else if (isOneHalf) {
          this._add('critical', 'spacing-1-5', 'п. 2.1.1',
            `Межстрочный интервал 1,5 строки (line=${line}). Требуется одинарный (1,0 = 18 пт)`,
            'Стиль «Обычный»',
            'Формат → Абзац → Межстрочный → Одинарный (или «Точно» 18 пт)');
        } else {
          this._pass('spacing-ok', 'п. 2.1.1', `Межстрочный интервал соответствует норме ✓`);
        }
      }
    }

    // --- Абзацный отступ (1,25 см = ~709 twips) ---
    if (pPr?.indent) {
      const fl = pPr.indent.firstLine;
      if (fl > 0) {
        const flCm = Math.round((fl / 1440) * 2.54 * 100) / 100;
        if (Math.abs(flCm - 1.25) > 0.15) {
          this._add('warning', 'para-indent', 'п. 2.1.1',
            `Абзацный отступ: ${flCm} см. Требуется 1,25 см`,
            'Стиль «Обычный»',
            'Установите абзацный отступ первой строки 1,25 см');
        } else {
          this._pass('indent-ok', 'п. 2.1.1', `Абзацный отступ: ${flCm} см ✓`);
        }
      } else if (fl === 0 && pPr.indent.left === 0) {
        // Нет отступа совсем — это может быть норм для заголовков, но не для тела
        // не будем ругаться здесь, проверим в отдельном блоке
      }
    }

    // --- Выравнивание ---
    if (pPr?.alignment) {
      if (pPr.alignment !== 'both' && pPr.alignment !== 'distribute') {
        this._add('warning', 'alignment', 'п. 2.1.1',
          `Выравнивание основного текста: «${pPr.alignment}». Требуется по ширине (both)`,
          'Стиль «Обычный»',
          'Установите выравнивание «По ширине» (Ctrl+J)');
      } else {
        this._pass('alignment-ok', 'п. 2.1.1', `Выравнивание по ширине ✓`);
      }
    }

    // Дополнительная проверка — смотрим на реальные параграфы с телом текста
    this._checkBodyParagraphFormatting();
  }

  _checkBodyParagraphFormatting() {
    // Выборка: реальные абзацы тела документа (не заголовки, не таблицы, достаточно длинные)
    const bodyParas = this.docData.paragraphs
      .filter(p => !p.isEmpty && !p.isHeading && !p.inTable && p.textTrimmed.length > 20)
      .slice(0, 50);

    if (bodyParas.length === 0) return;

    const issues = { font: [], size: [], spacing: [], indent: [], align: [] };

    for (const para of bodyParas) {
      const eff  = para.effectiveRPr || {};
      const effP = para.effectivePPr || {};

      // Шрифт — через effectiveRPr (включает наследование от стиля)
      if (eff.font && !eff.font.toLowerCase().includes('times')) {
        issues.font.push({ text: para.textTrimmed.slice(0, 50), page: this._page(para), font: eff.font });
      }

      // Размер — через effectiveRPr
      if (eff.size && Math.abs(eff.size - 14) > 0.5) {
        issues.size.push({ text: para.textTrimmed.slice(0, 50), page: this._page(para), size: eff.size });
      }

      // Межстрочный интервал — через effectivePPr
      const sp = effP.spacing;
      if (sp?.line > 0) {
        const rule = sp.lineRule;
        const isOneHalf = (!rule || rule === 'auto') && sp.line >= 340 && sp.line <= 400;
        const isDouble  = (!rule || rule === 'auto') && sp.line >= 450;
        if (isOneHalf || isDouble) {
          issues.spacing.push({ text: para.textTrimmed.slice(0, 50), page: this._page(para), line: sp.line });
        }
      }

      // Абзацный отступ — через effectivePPr
      const fl = effP.indent?.firstLine;
      if (fl !== undefined && fl !== null) {
        const flCm = Math.round((fl / 1440) * 2.54 * 100) / 100;
        if (fl > 0 && Math.abs(flCm - 1.25) > 0.2) {
          issues.indent.push({ text: para.textTrimmed.slice(0, 50), page: this._page(para), flCm });
        }
      }

      // Выравнивание — через effectivePPr
      const align = effP.alignment;
      if (align && align !== 'both' && align !== 'distribute') {
        issues.align.push({ text: para.textTrimmed.slice(0, 50), page: this._page(para), align });
      }
    }

    const total = bodyParas.length;
    const threshold = 0.25; // 25% порог

    if (issues.font.length > total * threshold) {
      const example = issues.font[0];
      this._add('critical', 'body-font-inconsistent', 'п. 2.1.1',
        `В ${issues.font.length} из ${total} абзацев шрифт не Times New Roman (пример: «${issues.font[0].font}» — «${example.text}»)`,
        'Основной текст',
        'Выделите весь текст (Ctrl+A) → установите шрифт Times New Roman',
        issues.font[0].page);
    }
    if (issues.size.length > total * threshold) {
      const example = issues.size[0];
      this._add('critical', 'body-size-inconsistent', 'п. 2.1.1',
        `В ${issues.size.length} из ${total} абзацев размер шрифта ≠ 14 пт (пример: ${example.size} пт — «${example.text}»)`,
        'Основной текст',
        'Выделите весь текст (Ctrl+A) → установите размер 14 пт',
        issues.size[0].page);
    }
    if (issues.spacing.length > total * threshold) {
      this._add('critical', 'body-spacing-inconsistent', 'п. 2.1.1',
        `Обнаружен увеличенный межстрочный интервал (≥1,5) в ${issues.spacing.length} абзацах`,
        'Основной текст',
        'Выделите весь текст → Формат → Абзац → Межстрочный → Одинарный',
        issues.spacing[0].page);
    }
    if (issues.indent.length > total * threshold) {
      const example = issues.indent[0];
      this._add('warning', 'body-indent-inconsistent', 'п. 2.1.1',
        `В ${issues.indent.length} из ${total} абзацев отступ ≠ 1,25 см (пример: ${example.flCm} см — «${example.text}»)`,
        'Основной текст',
        'Установите абзацный отступ 1,25 см: Формат → Абзац → Отступ первой строки',
        issues.indent[0].page);
    }
    if (issues.align.length > total * threshold) {
      this._add('warning', 'body-align-inconsistent', 'п. 2.1.1',
        `В ${issues.align.length} из ${total} абзацев выравнивание не «По ширине»`,
        'Основной текст',
        'Выделите весь основной текст → Ctrl+J (выравнивание по ширине)',
        issues.align[0].page);
    }
  }

  // ============================================================
  // 3. СТРУКТУРА ДОКУМЕНТА (п. 1.2.5)
  // ============================================================
  _checkDocumentStructure() {
    const fullText = this.docData.fullText.toUpperCase();
    const paragraphs = this.docData.paragraphs;

    // Обязательные разделы
    const required = [
      {
        id: 'has-abstract',     key: 'РЕФЕРАТ',
        label: 'Реферат',       section: 'п. 1.2.5, п. 1.2.8',
        rec: 'Добавьте раздел РЕФЕРАТ после титульного листа'
      },
      {
        id: 'has-toc',          key: 'СОДЕРЖАНИЕ',
        label: 'Содержание',    section: 'п. 1.2.5, п. 2.2.7',
        rec: 'Добавьте раздел СОДЕРЖАНИЕ (после задания)'
      },
      {
        id: 'has-intro',        key: 'ВВЕДЕНИЕ',
        label: 'Введение',      section: 'п. 1.2.5, п. 1.2.11',
        rec: 'Добавьте раздел ВВЕДЕНИЕ'
      },
      {
        id: 'has-conclusion',   key: 'ЗАКЛЮЧЕНИЕ',
        label: 'Заключение',    section: 'п. 1.2.5, п. 1.2.15',
        rec: 'Добавьте раздел ЗАКЛЮЧЕНИЕ'
      },
      {
        id: 'has-refs',         key: 'СПИСОК ИСПОЛЬЗОВАННЫХ',
        label: 'Список источников', section: 'п. 1.2.5, п. 2.8.1',
        rec: 'Добавьте «СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ»'
      },
    ];

    for (const req of required) {
      if (fullText.includes(req.key)) {
        this._pass(req.id + '-ok', req.section, `Раздел «${req.label}» найден ✓`);
      } else {
        this._add('critical', req.id, req.section,
          `Обязательный раздел «${req.label}» не найден`,
          'Структура документа',
          req.rec);
      }
    }

    // Ведомость документов
    if (!fullText.includes('ВЕДОМОСТЬ')) {
      this._add('warning', 'has-doc-list', 'п. 1.2.5, п. 1.2.19',
        'Не найдена «ВЕДОМОСТЬ ДОКУМЕНТОВ» (последний обязательный лист ПЗ)',
        'Структура документа',
        'Добавьте ведомость документов по форме из Приложения Д СТП');
    }

    // Экономический раздел
    const hasEcon = fullText.includes('ТЕХНИКО-ЭКОНОМИ') ||
                    fullText.includes('ЭКОНОМИЧЕСК') ||
                    fullText.includes('ЭКОНОМИЧЕСКОЕ ОБОСНОВАНИЕ');
    if (!hasEcon) {
      this._add('warning', 'has-economic', 'п. 1.2.5',
        'Не найден экономический раздел (ТЭО)',
        'Структура документа',
        'Добавьте технико-экономическое обоснование (не более 18% объёма ПЗ)');
    }

    // Раздел охраны труда/экологии
    const hasSafety = fullText.includes('ОХРАНА ТРУДА') ||
                      fullText.includes('ЭКОЛОГИЧЕСКАЯ БЕЗОПАСНОСТЬ') ||
                      fullText.includes('ЭНЕРГО') ||
                      fullText.includes('РЕСУРСОСБЕРЕЖЕН');
    if (!hasSafety) {
      this._add('info', 'has-safety', 'п. 1.2.5',
        'Не найден раздел охраны труда / экологической безопасности / энергосбережения',
        'Структура документа',
        'Добавьте соответствующий раздел (не более 5–7% объёма ПЗ)');
    }

    // Проверка перечня условных обозначений (если не тривиальный документ)
    // (необязательно, но стоит проверить если есть сокращения)
    const headingTexts = paragraphs.filter(p => p.isHeading).map(p => p.textTrimmed.toUpperCase());
    const hasAbbrSection = headingTexts.some(h => h.includes('ПЕРЕЧЕНЬ') && (h.includes('ОБОЗНАЧЕНИ') || h.includes('СИМВОЛ')));
    // Не критично, просто информируем
    if (!hasAbbrSection) {
      this._add('info', 'no-abbr-list', 'п. 1.2.5',
        'Не найден «Перечень условных обозначений, символов и терминов»',
        'Структура документа',
        'Добавьте перечень, если в документе используются нестандартные обозначения');
    }

    // Проверяем отчёт об оригинальности в заключении (п. 1.2.11)
    if (fullText.includes('ВВЕДЕНИЕ')) {
      const introIdx = this.docData.paragraphs.findIndex(
        p => p.textTrimmed.toUpperCase() === 'ВВЕДЕНИЕ'
      );
      if (introIdx !== -1) {
        const introText = this.docData.paragraphs
          .slice(introIdx, introIdx + 20)
          .map(p => p.text.toUpperCase())
          .join(' ');
        const introPara = this.docData.paragraphs[introIdx];
        const introPage = this._page(introPara);
        if (!introText.includes('ОРИГИНАЛЬНОСТ') && !introText.includes('ЗАИМСТВОВАН')) {
          this._add('warning', 'no-originality-stmt', 'п. 1.2.11',
            'В Введении не найдена фраза о проценте оригинальности (антиплагиат)',
            'Введение',
            'Добавьте в конец введения: «Данный дипломный проект выполнен мной лично, проверен на заимствования, процент оригинальности составляет ХХ%»',
            introPage);
        }
      }
    }
  }

  // ============================================================
  // 4. ЗАГОЛОВКИ (п. 2.2.2–2.2.6)
  // ============================================================
  _checkHeadings() {
    const headings = this.docData.paragraphs.filter(p => p.isHeading);

    if (headings.length === 0) {
      this._add('warning', 'no-headings', 'п. 2.2.2',
        'Структурные заголовки (с использованием стилей «Заголовок N») не обнаружены',
        'Заголовки',
        'Используйте стили Заголовок 1, Заголовок 2 и т.д. для разделов');
      return;
    }

    // Нумерация разделов верхнего уровня
    const sectionHeadings = headings.filter(h => h.headingLevel === 1);
    let expectedNum = 1;
    let numberingIssues = 0;

    for (const h of sectionHeadings) {
      const text = h.textTrimmed;
      if (!text) continue; // пропускаем пустые параграфы со 'style Heading'
      const numMatch = text.match(/^(\d+)\s+(.+)/);

      if (!numMatch) {
        // Проверяем, не является ли это ненумерованным разделом (Введение, Заключение и т.д.)
        const unnumbered = ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'РЕФЕРАТ', 'СОДЕРЖАНИЕ',
                            'СПИСОК', 'ПРИЛОЖЕНИЕ', 'ВЕДОМОСТЬ', 'ПЕРЕЧЕНЬ'];
        const isUnnumbered = unnumbered.some(u => text.toUpperCase().startsWith(u));
        if (!isUnnumbered) {
          this._add('warning', 'heading-no-number', 'п. 2.2.2',
            `Заголовок раздела без номера: «${text.slice(0, 60)}»`,
            `Заголовок: «${text.slice(0, 40)}»`,
            'Разделы нумеруются арабскими цифрами без точки: 1, 2, 3...',
            this._page(h));
          numberingIssues++;
        }
      } else {
        // Проверяем точку после номера
        const numPart = numMatch[1];
        if (text.startsWith(numPart + '.')) {
          this._add('warning', 'heading-dot-after-num', 'п. 2.2.2',
            `Номер раздела с точкой: «${numPart}.» Точка после номера не ставится`,
            `Заголовок: «${text.slice(0, 40)}»`,
            'Уберите точку после номера раздела: «1 Введение», не «1. Введение»',
            this._page(h));
          numberingIssues++;
        }
      }

      // Точка в конце заголовка
      if (/[.!?]$/.test(text)) {
        this._add('warning', 'heading-ends-dot', 'п. 2.2.5',
          `Заголовок заканчивается знаком препинания: «${text.slice(-20)}»`,
          `Заголовок: «${text.slice(0, 40)}»`,
          'Уберите точку (или другой знак) в конце заголовка (п. 2.2.5)',
          this._page(h));
      }

      // Проверяем: заголовок 1-го уровня должен быть заглавными буквами
      const textWithoutNum = numMatch ? numMatch[2] : text;
      const hasLower = /[а-яёa-z]/.test(textWithoutNum);
      if (hasLower && !this._isUnnumberedSection(text)) {
        this._add('warning', 'heading1-not-caps', 'п. 2.2.5',
          `Заголовок раздела содержит строчные буквы: «${text.slice(0, 60)}»`,
          `Заголовок: «${text.slice(0, 40)}»`,
          'Заголовки разделов (1-й уровень) пишутся ПРОПИСНЫМИ буквами (п. 2.2.5)',
          this._page(h));
      }

      // Перенос слов в заголовке
      if (text.includes('\u00AD') || text.includes('-\n') || text.match(/\w-\n/)) {
        this._add('warning', 'heading-hyphen', 'п. 2.1.1',
          `В заголовке обнаружен перенос слова: «${text.slice(0, 60)}»`,
          `Заголовок: «${text.slice(0, 40)}»`,
          'Переносы в заголовках не допускаются (п. 2.1.1)',
          this._page(h));
      }
    }

    // Проверяем заголовки 2-го уровня (подразделы)
    const subsectionHeadings = headings.filter(h => h.headingLevel === 2);
    for (const h of subsectionHeadings) {
      const text = h.textTrimmed;
      if (!text) continue; // пропускаем пустые параграфы

      // Формат: Х.Х Название (строчными, с прописной первой)
      const numMatch = text.match(/^(\d+\.\d+)\s+(.+)/);
      if (numMatch) {
        const titlePart = numMatch[2];
        // Первая буква должна быть прописная, остальные — строчные (первое слово)
        if (/^[А-ЯЁA-Z]{2,}/.test(titlePart.split(' ').slice(1).join(' '))) {
          this._add('info', 'heading2-all-caps', 'п. 2.2.5',
            `Заголовок подраздела написан ПРОПИСНЫМИ буквами: «${text.slice(0, 60)}»`,
            `Заголовок: «${text.slice(0, 40)}»`,
            'Заголовки подразделов пишутся строчными буквами, начиная с прописной (п. 2.2.5)',
            this._page(h));
        }
      }

      // Точка в конце
      if (/[.!?]$/.test(text)) {
        this._add('warning', 'heading2-ends-dot', 'п. 2.2.5',
          `Заголовок подраздела заканчивается знаком препинания`,
          `Заголовок: «${text.slice(0, 40)}»`,
          'Уберите точку в конце заголовка подраздела',
          this._page(h));
      }
    }

    // Пустая строка между заголовком и текстом (п. 2.2.6)
    // Проверяем, есть ли пустой параграф после заголовка
    const paras = this.docData.paragraphs;
    let missingBlankCount = 0;
    for (let i = 0; i < paras.length - 1; i++) {
      if (paras[i].isHeading) {
        const next = paras[i + 1];
        // Следующий должен быть пустым или заголовком, или сам иметь spacing
        if (!next.isEmpty && !next.isHeading) {
          const spacingBefore = next.pPr?.spacing?.before || 0;
          if (spacingBefore < 100) { // < ~7pt
            missingBlankCount++;
          }
        }
      }
    }
    if (missingBlankCount > 2) {
      this._add('warning', 'heading-no-blank-line', 'п. 2.2.6',
        `Обнаружено ${missingBlankCount} заголовков без пробельной строки перед текстом`,
        'Заголовки',
        'После каждого заголовка оставляйте одну пробельную строку перед текстом (п. 2.2.6)');
    } else if (missingBlankCount === 0 && sectionHeadings.length > 0) {
      this._pass('heading-blank-ok', 'п. 2.2.6', 'Пробельные строки после заголовков ✓');
    }

    if (numberingIssues === 0 && sectionHeadings.length > 0) {
      this._pass('headings-numbering-ok', 'п. 2.2.2', 'Нумерация заголовков разделов ✓');
    }
  }

  _isUnnumberedSection(text) {
    const t = text.toUpperCase();
    return ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'РЕФЕРАТ', 'СОДЕРЖАНИЕ',
            'СПИСОК', 'ПРИЛОЖЕНИЕ', 'ВЕДОМОСТЬ', 'ПЕРЕЧЕНЬ'].some(k => t.startsWith(k));
  }

  // ============================================================
  // 4б. ФОРМАТИРОВАНИЕ ЗАГОЛОВКОВ (п. 2.1.1, п. 2.2.5)
  // ============================================================
  _checkHeadingFormatting() {
    const headings = this.docData.paragraphs.filter(p => p.isHeading && p.textTrimmed);
    if (headings.length === 0) return;

    const UNNUMBERED = ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'РЕФЕРАТ', 'СОДЕРЖАНИЕ',
                        'СПИСОК', 'ПРИЛОЖЕНИЕ', 'ВЕДОМОСТЬ', 'ПЕРЕЧЕНЬ'];

    let notBoldCount = 0;
    let notBoldExample = null;
    let underlineCount = 0;
    let underlineExample = null;
    let centeredIssues = [];   // ненумерованные не по центру
    let leftIssues = [];       // нумерованные не по левому краю

    for (const h of headings) {
      const text = h.textTrimmed;
      const eff  = h.effectiveRPr || {};
      const effP = h.effectivePPr || {};

      // ── Жирность ─────────────────────────────────────────
      const isBold = eff.bold ||
                     h.firstRunRPr?.bold ||
                     h.pRPr?.bold ||
                     h.style?.rPr?.bold;
      if (!isBold) {
        notBoldCount++;
        if (!notBoldExample) notBoldExample = h;
      }

      // ── Подчёркивание (не допускается) ────────────────────
      const isUnderlined = eff.underline ||
                           h.firstRunRPr?.underline ||
                           h.runs.some(r => r.rPr?.underline);
      if (isUnderlined) {
        underlineCount++;
        if (!underlineExample) underlineExample = h;
      }

      // ── Выравнивание ──────────────────────────────────────
      const align = effP.alignment || h.pPr?.alignment;
      const isUnnumbered = UNNUMBERED.some(u => text.toUpperCase().startsWith(u));
      if (isUnnumbered) {
        // ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и т.д. должны быть по центру
        if (align && align !== 'center') {
          centeredIssues.push({ text: text.slice(0, 40), align, page: this._page(h) });
        }
      } else {
        // Нумерованные разделы — по левому краю (left или null/undefined)
        if (align && align !== 'left' && align !== 'start') {
          // «по ширине» у заголовка тоже ошибка, но менее критично
          if (align === 'center' || align === 'right') {
            leftIssues.push({ text: text.slice(0, 40), align, page: this._page(h) });
          }
        }
      }
    }

    // Отчёт по жирности
    if (notBoldCount > 0) {
      const ex = notBoldExample;
      this._add('warning', 'headings-not-bold', 'п. 2.1.1',
        `${notBoldCount} заголовк(ов) не выделены полужирным шрифтом. Пример: «${ex.textTrimmed.slice(0, 50)}»`,
        `Заголовок: «${ex.textTrimmed.slice(0, 40)}»`,
        'Все заголовки разделов и подразделов должны быть полужирными (п. 2.1.1)',
        this._page(notBoldExample));
    } else {
      this._pass('headings-bold-ok', 'п. 2.1.1', 'Все заголовки выделены полужирным ✓');
    }

    // Отчёт по подчёркиванию
    if (underlineCount > 0) {
      this._add('warning', 'headings-underline', 'п. 2.2.5',
        `${underlineCount} заголовк(ов) содержат подчёркивание — запрещено. Пример: «${underlineExample.textTrimmed.slice(0, 50)}»`,
        `Заголовок: «${underlineExample.textTrimmed.slice(0, 40)}»`,
        'Заголовки не подчёркивают (п. 2.2.5). Снимите подчёркивание: выделите → Ctrl+U',
        this._page(underlineExample));
    }

    // Отчёт по выравниванию ненумерованных
    if (centeredIssues.length > 0) {
      const ex = centeredIssues[0];
      this._add('warning', 'unnumbered-heading-not-centered', 'п. 2.1.1',
        `Ненумерованный раздел «${ex.text}» выровнен не по центру (текущее: «${ex.align}»)`,
        `Заголовок: «${ex.text}»`,
        'ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, РЕФЕРАТ, СОДЕРЖАНИЕ и СПИСОК выравниваются по центру строки без абзацного отступа',
        ex.page);
    }

    // Отчёт по выравниванию нумерованных
    if (leftIssues.length > 0) {
      const ex = leftIssues[0];
      this._add('info', 'numbered-heading-not-left', 'п. 2.1.1',
        `Нумерованный заголовок «${ex.text}» выровнен не по левому краю (текущее: «${ex.align}»)`,
        `Заголовок: «${ex.text}»`,
        'Заголовки нумерованных разделов и подразделов выравниваются по левому краю',
        ex.page);
    }
  }

  // ============================================================
  // 3б. ПОРЯДОК РАЗДЕЛОВ (п. 1.2.5)
  // ============================================================
  _checkDocumentOrder() {
    const paras = this.docData.paragraphs;
    const findIdx = (keyword) => paras.findIndex(p => p.textTrimmed.toUpperCase().includes(keyword));

    const idxAbstract   = findIdx('РЕФЕРАТ');
    const idxToc        = findIdx('СОДЕРЖАНИЕ');
    const idxIntro      = findIdx('ВВЕДЕНИЕ');
    const idxConclusion = findIdx('ЗАКЛЮЧЕНИЕ');
    const idxRefs       = findIdx('СПИСОК ИСПОЛЬЗОВАННЫХ');

    // Порядок: РЕФЕРАТ → СОДЕРЖАНИЕ → ВВЕДЕНИЕ → ... → ЗАКЛЮЧЕНИЕ → СПИСОК
    const order = [
      { name: 'РЕФЕРАТ',           idx: idxAbstract,   id: 'order-abstract' },
      { name: 'СОДЕРЖАНИЕ',        idx: idxToc,        id: 'order-toc' },
      { name: 'ВВЕДЕНИЕ',          idx: idxIntro,      id: 'order-intro' },
      { name: 'ЗАКЛЮЧЕНИЕ',        idx: idxConclusion, id: 'order-conclusion' },
      { name: 'СПИСОК ИСТОЧНИКОВ', idx: idxRefs,       id: 'order-refs' },
    ];

    let prevIdx = -1;
    let prevName = '';
    for (const item of order) {
      if (item.idx === -1) continue; // раздел не найден — проверяется в _checkDocumentStructure
      if (item.idx < prevIdx) {
        this._add('critical', item.id, 'п. 1.2.5',
          `Раздел «${item.name}» расположен ДО «${prevName}» — нарушен порядок разделов`,
          `Структура документа`,
          `По СТП порядок: РЕФЕРАТ → СОДЕРЖАНИЕ → ВВЕДЕНИЕ → основная часть → ЗАКЛЮЧЕНИЕ → СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ → ВЕДОМОСТЬ`);
      }
      prevIdx = item.idx;
      prevName = item.name;
    }
  }

  // ============================================================
  // 3в. РЕФЕРАТ (п. 1.2.8)
  // ============================================================
  _checkAbstract() {
    const paras = this.docData.paragraphs;
    const refIdx = paras.findIndex(p => p.textTrimmed.toUpperCase() === 'РЕФЕРАТ');
    if (refIdx === -1) return;

    const refPara = paras[refIdx];
    const pg = this._page(refPara);

    // 1. РЕФЕРАТ должен быть написан прописными, жирным, по центру
    const isBold = refPara.effectiveRPr?.bold || refPara.firstRunRPr?.bold ||
                   refPara.pRPr?.bold || refPara.style?.rPr?.bold || refPara.isHeading;
    if (!isBold) {
      this._add('warning', 'abstract-not-bold', 'п. 1.2.8',
        'Слово «РЕФЕРАТ» должно быть выделено полужирным шрифтом',
        'Реферат', 'Выделите «РЕФЕРАТ» → Ctrl+B', pg);
    }

    const refAlign = refPara.effectivePPr?.alignment || refPara.pPr?.alignment || refPara.style?.pPr?.alignment;
    if (refAlign && refAlign !== 'center') {
      this._add('warning', 'abstract-not-centered', 'п. 1.2.8',
        'Слово «РЕФЕРАТ» должно быть выровнено по центру',
        'Реферат', 'Выберите выравнивание по центру для «РЕФЕРАТ»', pg);
    }

    // 2. Содержание реферата — берём ~30 абзацев после заголовка
    const refBody = paras.slice(refIdx + 1, refIdx + 35).map(p => p.text.toUpperCase()).join(' ');
    const KEYWORDS = [
      { key: 'КЛЮЧЕВЫЕ СЛОВА',   id: 'abstract-no-keywords',   msg: 'Не найден раздел «Ключевые слова» в реферате', rec: 'Добавьте строку «КЛЮЧЕВЫЕ СЛОВА:» в конец реферата (п. 1.2.8)' },
      { key: 'ОБЪЕКТ',           id: 'abstract-no-object',     msg: 'Не указан объект исследования в реферате',      rec: 'Укажите объект исследования/разработки в реферате' },
      { key: 'ЦЕЛЬ',             id: 'abstract-no-goal',       msg: 'Не указана цель работы в реферате',             rec: 'Укажите цель работы/проекта в реферате' },
    ];
    for (const kw of KEYWORDS) {
      if (!refBody.includes(kw.key)) {
        this._add('info', kw.id, 'п. 1.2.8', kw.msg, 'Реферат', kw.rec, pg);
      }
    }

    // 3. Количество слов в реферате (~200–500 слов по ГОСТ 7.9)
    const wordCount = refBody.split(/\s+/).filter(w => w.length > 2).length;
    if (wordCount < 50) {
      this._add('warning', 'abstract-too-short', 'п. 1.2.8, ГОСТ 7.9',
        `Реферат слишком короткий (~${wordCount} слов). По ГОСТ 7.9 рекомендуется 200–500 слов`,
        'Реферат', 'Расширьте реферат: добавьте объект, методы, результаты, ключевые слова', pg);
    }
  }

  // ============================================================
  // 6б. ПОСЛЕДОВАТЕЛЬНАЯ НУМЕРАЦИЯ РИСУНКОВ И ТАБЛИЦ (п. 2.5.5, п. 2.6.2)
  // ============================================================
  _checkFigureTableSequencing() {
    const paras = this.docData.paragraphs;

    // Извлекаем номера рисунков
    const figNums = [];
    for (const para of paras) {
      const m = para.textTrimmed.match(/^Рисунок\s+(\d+)/i);
      if (m) figNums.push({ num: parseInt(m[1]), text: para.textTrimmed.slice(0, 50), page: this._page(para) });
    }
    // Извлекаем номера таблиц
    const tblNums = [];
    for (const para of paras) {
      const m = para.textTrimmed.match(/^Таблица\s+(\d+)/i);
      if (m) tblNums.push({ num: parseInt(m[1]), text: para.textTrimmed.slice(0, 50), page: this._page(para) });
    }

    // Проверяем сквозную последовательность
    const checkSeq = (items, label) => {
      for (let i = 1; i < items.length; i++) {
        const expected = items[i - 1].num + 1;
        if (items[i].num !== expected && items[i].num !== items[i - 1].num) {
          this._add('warning', `${label.toLowerCase()}-seq-gap`, 'п. 2.5.5',
            `Нарушена нумерация ${label}: после ${label.toLowerCase()} ${items[i-1].num} идёт ${items[i].num} (ожидался ${expected}). Пример: «${items[i].text}»`,
            `${label}`,
            `Нумерация ${label.toLowerCase()} должна быть сквозной (1, 2, 3, ...) или по разделам (1.1, 1.2, ...)`,
            items[i].page);
          return; // одно сообщение, не засыпать ошибками
        }
      }
    };
    if (figNums.length > 1) checkSeq(figNums, 'Рисунков');
    if (tblNums.length > 1) checkSeq(tblNums, 'Таблиц');
  }

  // ============================================================
  // 5. СОДЕРЖАНИЕ (п. 2.2.7)
  // ============================================================
  _checkTableOfContents() {
    const paras = this.docData.paragraphs;
    const tocIdx = paras.findIndex(p => p.textTrimmed.toUpperCase() === 'СОДЕРЖАНИЕ');

    if (tocIdx === -1) return; // Уже проверили в структуре

    const tocPara = paras[tocIdx];
    const text = tocPara.textTrimmed;

    // СОДЕРЖАНИЕ должно быть жирным
    // Жирность может быть задана: в run, в pPr/rPr, в стиле параграфа, или параграф — заголовок
    const isBold = tocPara.firstRunRPr?.bold ||
                   tocPara.pRPr?.bold ||
                   tocPara.style?.rPr?.bold ||
                   tocPara.isHeading; // стили «Заголовок N» всегда жирные
    if (!isBold) {
      this._add('warning', 'toc-not-bold', 'п. 2.2.7',
        'Слово «СОДЕРЖАНИЕ» должно быть выделено полужирным шрифтом',
        'Содержание',
        'Установите полужирный (Ctrl+B) для слова «СОДЕРЖАНИЕ»');
    }

    // Выравнивание — по центру
    const alignment = tocPara.pPr?.alignment || tocPara.style?.pPr?.alignment;
    if (alignment && alignment !== 'center') {
      this._add('warning', 'toc-not-centered', 'п. 2.2.7',
        'Слово «СОДЕРЖАНИЕ» должно быть выровнено по центру строки',
        'Содержание',
        'Установите выравнивание по центру для «СОДЕРЖАНИЕ»');
    }
  }

  // ============================================================
  // 6. РИСУНКИ (п. 2.5)
  // ============================================================
  _checkFigureCaptions() {
    const paras = this.docData.paragraphs;

    // Паттерны подписей
    const validPattern   = /^Рисунок\s+[\dА-Яа-яA-Za-z]+(\.\d+)?\s*[–—]\s*.+/i;
    const abbrevPattern  = /^Рис\.\s*\d/i;
    const badDashPattern = /^Рисунок\s+[\dА-Яа-яA-Za-z]+(\.\d+)?\s*[-–]\s*/i; // обычный дефис вместо тире

    const figures = [];
    let badCaptionCount = 0;
    let abbrevCount = 0;
    let badDashCount = 0;
    let trailingDotCount = 0;

    for (const para of paras) {
      const text = para.textTrimmed;
      if (!text) continue;
      const pg = this._page(para);

      // Правильная подпись
      if (validPattern.test(text)) {
        figures.push({ text, valid: true });
        // Проверяем точку в конце
        if (/[.]$/.test(text)) {
          trailingDotCount++;
          this._add('warning', 'figure-trailing-dot', 'п. 2.5.5',
            `Точка в конце подписи рисунка: «${text.slice(0, 60)}»`,
            `«${text.slice(0, 50)}»`,
            'Убрать точку в конце подрисуночной подписи (п. 2.5.5)', pg);
        }
        // Проверяем обычный дефис вместо тире (–)
        if (/^Рисунок\s+[^\s]+\s+-\s+/.test(text)) {
          badDashCount++;
          this._add('warning', 'figure-bad-dash', 'п. 2.5.5',
            `Обычный дефис вместо тире: «${text.slice(0, 60)}»`,
            `«${text.slice(0, 50)}»`,
            'Используйте длинное тире «–» (Alt+0150), а не дефис «-»: «Рисунок 1 – Название»', pg);
        }
        continue;
      }

      // Сокращение «Рис.»
      if (abbrevPattern.test(text)) {
        abbrevCount++;
        figures.push({ text, valid: false });
        this._add('warning', 'figure-abbrev', 'п. 2.5.5',
          `Сокращение «Рис.» не допускается: «${text.slice(0, 60)}»`,
          `«${text.slice(0, 50)}»`,
          'Пишите «Рисунок» полностью без сокращений (п. 2.5.5)', pg);
        continue;
      }

      // Проверяем на паттерн «Рисунок» без тире
      if (/^Рисунок\s+\d+\s+[А-Яа-я]/i.test(text)) {
        badCaptionCount++;
        this._add('warning', 'figure-no-dash', 'п. 2.5.5',
          `Нет тире в подписи рисунка: «${text.slice(0, 60)}»`,
          `«${text.slice(0, 50)}»`,
          'Формат: «Рисунок 1 – Название рисунка» (тире «–» обязательно)', pg);
      }
    }

    if (figures.length === 0) {
      this._add('info', 'no-figures', 'п. 2.5.5',
        'Подписи рисунков (начинающиеся с «Рисунок N – ...») не обнаружены',
        'Рисунки',
        'Если в работе есть иллюстрации, каждая должна иметь подпись «Рисунок N – Название»');
    } else if (abbrevCount === 0 && badCaptionCount === 0 && badDashCount === 0 && trailingDotCount === 0) {
      this._pass('figures-ok', 'п. 2.5.5', `Найдено ${figures.length} рисунков, формат подписей верный ✓`);
    }

    // Проверяем ссылки на рисунки в тексте (п. 2.5.6)
    if (figures.length > 0) {
      const fullText = this.docData.fullText;
      const hasRefs = /рисунк[еу]?\s+\d+|рисунок\s+\d+/i.test(fullText);
      if (!hasRefs) {
        this._add('warning', 'figure-no-refs', 'п. 2.5.6',
          'Не обнаружены ссылки на рисунки в тексте',
          'Ссылки на рисунки',
          'В тексте должны быть ссылки на все рисунки: «в соответствии с рисунком 1» или «на рисунке 2 изображено...»');
      }
    }
  }

  // ============================================================
  // 7. ТАБЛИЦЫ (п. 2.6)
  // ============================================================
  _checkTableCaptions() {
    const paras = this.docData.paragraphs;

    const validPattern  = /^Таблица\s+[\dА-Яа-яA-Za-z]+(\.\d+)?\s*[–—]\s*.+/i;
    const abbrevPattern = /^Табл?\.\s*\d/i;

    const tables = [];
    let abbrevCount = 0;
    let badDashCount = 0;
    let trailingDotCount = 0;
    let noDashCount = 0;

    for (const para of paras) {
      const text = para.textTrimmed;
      if (!text) continue;
      const pg = this._page(para);

      if (validPattern.test(text)) {
        tables.push({ text, valid: true });
        if (/[.]$/.test(text)) {
          trailingDotCount++;
          this._add('warning', 'table-trailing-dot', 'п. 2.6.2',
            `Точка в конце заголовка таблицы: «${text.slice(0, 60)}»`,
            `«${text.slice(0, 50)}»`,
            'Убрать точку в конце заголовка таблицы', pg);
        }
        if (/^Таблица\s+[^\s]+\s+-\s+/.test(text)) {
          badDashCount++;
          this._add('warning', 'table-bad-dash', 'п. 2.6.2',
            `Дефис вместо тире в заголовке таблицы: «${text.slice(0, 60)}»`,
            `«${text.slice(0, 50)}»`,
            'Используйте «–» (тире), а не «-» (дефис): «Таблица 1 – Название»', pg);
        }
        continue;
      }

      if (abbrevPattern.test(text)) {
        abbrevCount++;
        this._add('warning', 'table-abbrev', 'п. 2.6.2',
          `Сокращение «Табл.» не допускается: «${text.slice(0, 60)}»`,
          `«${text.slice(0, 50)}»`,
          'Слово «Таблица» пишется полностью (п. 2.6.2)', pg);
        continue;
      }

      if (/^Таблица\s+\d+\s+[А-Яа-я]/i.test(text)) {
        noDashCount++;
        this._add('warning', 'table-no-dash', 'п. 2.6.2',
          `Нет тире после номера таблицы: «${text.slice(0, 60)}»`,
          `«${text.slice(0, 50)}»`,
          'Формат: «Таблица 1 – Название таблицы»', pg);
      }
    }

    if (tables.length === 0 && this.docData.tables.length > 0) {
      this._add('warning', 'tables-no-captions', 'п. 2.6.2',
        `Найдено ${this.docData.tables.length} таблиц без заголовков`,
        'Таблицы',
        'Каждая таблица должна иметь заголовок формата «Таблица N – Название»');
    } else if (tables.length > 0 && abbrevCount === 0 && badDashCount === 0 && trailingDotCount === 0 && noDashCount === 0) {
      this._pass('tables-ok', 'п. 2.6.2', `Найдено ${tables.length} таблиц, формат заголовков верный ✓`);
    }

    // Проверяем ссылки на таблицы (п. 2.6.2)
    if (tables.length > 0) {
      const hasRefs = /таблиц[еуы]?\s+\d+|таблицу\s+\d+/i.test(this.docData.fullText);
      if (!hasRefs) {
        this._add('warning', 'table-no-refs', 'п. 2.6.2',
          'Не обнаружены ссылки на таблицы в тексте',
          'Ссылки на таблицы',
          'В тексте должны быть ссылки на все таблицы (п. 2.6.2)');
      }
    }
  }

  // ============================================================
  // 8. ФОРМУЛЫ (п. 2.4)
  // ============================================================
  _checkFormulas() {
    const paras = this.docData.paragraphs;
    const fullText = this.docData.fullText;

    // Нумерация формул — ищем паттерн (N.N) или (N) в конце параграфа
    const formulaNumPattern = /\(\d+\.\d+\)$|\(\d+\)$/;
    const formulas = paras.filter(p => formulaNumPattern.test(p.textTrimmed));

    if (formulas.length === 0) {
      this._add('info', 'no-formulas', 'п. 2.4.6',
        'Нумерованные формулы не обнаружены',
        'Формулы',
        'Все формулы на отдельных строках должны быть пронумерованы в формате (N.N) (п. 2.4.6)');
      return;
    }

    this._pass('formulas-found', 'п. 2.4.6', `Найдено ${formulas.length} формул с нумерацией ✓`);

    // Проверяем, что после формулы есть «где»
    let noWhereCount = 0;
    let whereWithColonCount = 0;
    for (let i = 0; i < paras.length; i++) {
      if (formulaNumPattern.test(paras[i].textTrimmed)) {
        // Ищем «где» в следующих 3 параграфах
        const nextTexts = paras.slice(i + 1, i + 4).map(p => p.textTrimmed);
        const whereIdx = nextTexts.findIndex(t => /^(где|здесь|here)\b/i.test(t));
        const hasWhere = whereIdx !== -1;

        if (!hasWhere && i < paras.length - 1) {
          const next = paras[i + 1];
          if (next && !next.isEmpty && !next.isHeading) {
            noWhereCount++;
          }
        }

        // «где:» — двоеточие после «где» не ставится (п. 2.4.7)
        if (hasWhere && /^где\s*:/i.test(nextTexts[whereIdx])) {
          whereWithColonCount++;
          const pg = this._page(paras[i + 1 + whereIdx]);
          this._add('info', 'formula-where-colon', 'п. 2.4.7',
            'После слова «где» двоеточие не ставится',
            'Формулы',
            'Пишите: «где  J – момент инерции...», без двоеточия после «где» (п. 2.4.7)',
            pg);
        }
      }
    }

    if (noWhereCount > 2) {
      this._add('warning', 'formula-no-where', 'п. 2.4.7',
        `В ${noWhereCount} формулах не найдена расшифровка символов («где ...»)`,
        'Формулы',
        'После каждой формулы необходимо расшифровать символы: «где», с новой строки без абзацного отступа (п. 2.4.7)');
    }

    // Ссылки на формулы в тексте
    const hasFormulaRefs = /формул[еуи]\s*\([^)]+\)|уравнени[еи]\s*\(|выражени[еи]\s*\(/.test(fullText);
    if (!hasFormulaRefs && formulas.length > 0) {
      this._add('info', 'formula-no-refs', 'п. 2.4.7',
        'Не обнаружены ссылки на формулы в тексте',
        'Ссылки на формулы',
        'Ссылайтесь на формулы: «подставляя в формулу (2.1)...» или «из выражения (2.7)...»');
    }
  }

  // ============================================================
  // 9. СПИСОК ИСТОЧНИКОВ (п. 2.8)
  // ============================================================
  _checkReferences() {
    const fullText = this.docData.fullText;
    const paras = this.docData.paragraphs;

    // Ищем начало списка источников
    const listStart = paras.findIndex(p =>
      p.textTrimmed.toUpperCase().includes('СПИСОК ИСПОЛЬЗОВАННЫХ') ||
      p.textTrimmed.toUpperCase().includes('СПИСОК ИСТОЧНИКОВ')
    );

    // Проверяем ссылки в тексте формата [N]
    const inTextRefs = fullText.match(/\[\d+\]/g) || [];
    if (inTextRefs.length === 0) {
      this._add('warning', 'refs-no-inline', 'п. 2.8.2',
        'Ссылки на источники в тексте [N] не обнаружены',
        'Список источников',
        'Ссылайтесь на источники в тексте: «...архитектура применяется [6]»');
    } else {
      this._pass('refs-inline-ok', 'п. 2.8.2', `Найдено ${inTextRefs.length} ссылок [N] в тексте ✓`);
    }

    if (listStart === -1) return;

    // Анализируем записи в списке источников
    const refParas = paras.slice(listStart + 1, listStart + 60).filter(p => !p.isEmpty);

    // Проверяем Wikipedia
    const hasWikipedia = refParas.some(p =>
      p.text.toLowerCase().includes('wikipedia') || p.text.toLowerCase().includes('en.wiki') || p.text.toLowerCase().includes('ru.wiki')
    );
    if (hasWikipedia) {
      this._add('critical', 'refs-wikipedia', 'п. 1.2.16',
        'В списке источников обнаружен Wikipedia — запрещённый источник',
        'Список использованных источников',
        'Удалите ссылки на Wikipedia. Этот источник запрещён (п. 1.2.16)');
    } else if (refParas.length > 0) {
      this._pass('refs-no-wiki', 'п. 1.2.16', 'Wikipedia в списке источников не обнаружена ✓');
    }

    // Проверяем формат записей по ГОСТ 7.1-2003
    // Признак: запись начинается с [N] или просто N
    let wrongFormatCount = 0;
    let goodFormatCount = 0;

    for (const para of refParas.slice(0, 20)) {
      const text = para.textTrimmed;
      if (!text) continue;

      // Паттерн ГОСТ: [N] Автор, Инициалы. Название / ... – Место : Издательство, Год. – N с.
      // Или просто начинается с порядкового номера
      const isNumbered = /^\[\d+\]/.test(text) || /^\d+\s/.test(text);
      if (isNumbered) {
        goodFormatCount++;
        // Проверяем наличие тире-разделителя по ГОСТ
        if (!text.includes(' – ') && !text.includes(' — ') && !text.includes(' : ')) {
          wrongFormatCount++;
        }
      }
    }

    if (wrongFormatCount > goodFormatCount * 0.5 && goodFormatCount > 2) {
      this._add('warning', 'refs-format', 'п. 2.8.5',
        `${wrongFormatCount} записей в списке источников могут не соответствовать ГОСТ 7.1-2003`,
        'Список использованных источников',
        'Оформляйте источники по ГОСТ 7.1-2003: Фамилия, И.О. Название / Соавторы. – Место : Изд-во, Год. – N с.');
    }

    if (refParas.length > 0 && !hasWikipedia) {
      this._pass('refs-found', 'п. 2.8', `Найдено ~${refParas.length} источников в списке ✓`);
    }

    // Проверяем последовательность нумерации
    if (goodFormatCount > 0 && inTextRefs.length > 0) {
      const maxInText = Math.max(...inTextRefs.map(r => parseInt(r.replace(/\[|\]/g, ''))));
      if (maxInText > refParas.length + 5) {
        this._add('warning', 'refs-count-mismatch', 'п. 2.8.3',
          `Максимальная ссылка в тексте [${maxInText}], но источников в списке ~${refParas.length}`,
          'Список источников',
          'Убедитесь, что каждый источник в тексте есть в списке (п. 2.8.3)');
      }
    }
  }

  // ============================================================
  // 10. ПРАВИЛА ТЕКСТА (п. 2.3)
  // ============================================================
  _checkTextRules() {
    const allParas = this.docData.paragraphs.filter(p => !p.isEmpty && !p.isHeading && !p.inTable);
    const paras = allParas;
    const fullText = this.docData.fullText;

    // ── 1. Дефис вместо тире в предложениях (п. 2.3.3) ──────────────────────
    // Пробел-дефис-пробел → должно быть тире «–» (U+2013)
    const badDashParas = [];
    for (const para of paras) {
      const t = para.textTrimmed;
      // Ищем « - » но исключаем числовые диапазоны (2 - 5) и «(-» начало скобки
      if (/ - /.test(t) && !/^\s*\d/.test(t)) {
        badDashParas.push(para);
      }
    }
    if (badDashParas.length > 0) {
      const ex = badDashParas[0];
      const snip = ex.textTrimmed.slice(0, 80);
      this._add('warning', 'bad-hyphen-as-dash', 'п. 2.3.3',
        `Дефис «-» вместо тире «–» в ${badDashParas.length} абзаце(ах). Пример: «${snip}»`,
        'Текст',
        'Замените « - » на « – » (тире — Alt+0150 или Ctrl+Minuspad). Дефис используется только для переносов и составных слов.',
        this._page(ex));
    }

    // ── 2. Двойной пробел ──────────────────────────────────────────────────
    const doubleSpaceParas = paras.filter(p => /  /.test(p.textTrimmed));
    if (doubleSpaceParas.length > 2) {
      this._add('info', 'double-space', 'п. 2.1.1',
        `Двойные пробелы обнаружены в ${doubleSpaceParas.length} абзацах`,
        'Текст',
        'Используйте Ctrl+H → Найти: «  » (2 пробела), Заменить: « » (1 пробел)',
        this._page(doubleSpaceParas[0]));
    }

    // ── 3. Пробел перед знаком препинания (.,;:?!) ─────────────────────────
    const spaceBeforePunct = paras.filter(p => / [,;:.!?]/.test(p.textTrimmed));
    if (spaceBeforePunct.length > 0) {
      const ex = spaceBeforePunct[0];
      this._add('warning', 'space-before-punct', 'п. 2.1.1',
        `Пробел перед знаком препинания в ${spaceBeforePunct.length} абзаце(ах). Пример: «${ex.textTrimmed.slice(0, 60)}»`,
        'Текст',
        'Запятая/точка ставится СРАЗУ после слова, без пробела перед ней.',
        this._page(ex));
    }

    // ── 4. Кавычки: « » (ёлочки) vs " " или '' ────────────────────────────
    // В русском тексте используются «ёлочки» (п. 2.3.9)
    const wrongQuoteParas = paras.filter(p => /"[^"]*"/.test(p.textTrimmed) || /'[^']*'/.test(p.textTrimmed));
    if (wrongQuoteParas.length > 0) {
      this._add('warning', 'wrong-quotes', 'п. 2.3.9',
        `Прямые кавычки " " или ' ' в ${wrongQuoteParas.length} абзаце(ах)`,
        'Текст',
        'Используйте «ёлочки» — Alt+0171 («) и Alt+0187 (») для кавычек в русском тексте.',
        this._page(wrongQuoteParas[0]));
    }

    // ── 5. №, % без числа (п. 2.3.11) ──────────────────────────────────────
    const badPercentSign = fullText.match(/(?<!\d\s*)%(?!\s*\d)/g) || [];
    if (badPercentSign.length > 0) {
      this._add('info', 'bad-percent', 'п. 2.3.11',
        'Знак «%» без числового значения — следует писать словом «процент»',
        'Текст',
        'В тексте (не в формулах/таблицах) знак % без числа пишется словом: «процент» (п. 2.3.11)');
    }

    // ── 6. Предлог «в» перед числом с размерностью (п. 2.3.12) ─────────────
    const badInPrepositions = fullText.match(/\bв\s+\d+[\s,]\s*(Вт|кВт|Гц|МГц|кб|МБ|ГБ|мм|см|км|нс|мкс|мс|кОм|нФ|мкФ)/g) || [];
    if (badInPrepositions.length > 0) {
      this._add('info', 'bad-in-preposition', 'п. 2.3.12',
        `Найдено ${badInPrepositions.length} случаев предлога «в» перед числом с размерностью`,
        'Текст',
        'Не пишите: «мощностью в 600 Вт» → пишите: «мощностью 600 Вт» (п. 2.3.12)');
    }

    // ── 7. Маркер «•» вместо тире в перечислениях (п. 2.3.5) ─────────────
    let badListChars = 0;
    const badBulletParas = [];
    for (const para of paras) {
      const text = para.textTrimmed;
      if (text.startsWith('•') || text.startsWith('·') || text.startsWith('*') || text.startsWith('▪') || text.startsWith('○')) {
        badListChars++;
        if (badListChars <= 3) badBulletParas.push(para);
      }
    }
    if (badListChars > 0) {
      this._add('warning', 'bad-list-marker', 'п. 2.3.5',
        `Найдено ${badListChars} пунктов перечисления с маркером «•»/«▪» вместо тире «–»`,
        'Перечисления',
        'Для перечислений используйте тире «–» (Alt+0150): «– пункт первый;» (п. 2.3.5)',
        this._page(badBulletParas[0]));
    }

    // ── 8. Точка с запятой в конце пунктов перечисления (п. 2.3.5) ─────────
    const goodListItems = paras.filter(p =>
      p.textTrimmed.startsWith('–') || p.textTrimmed.startsWith('—')
    );
    if (goodListItems.length > 0) {
      const noSemicolon = goodListItems.filter(p =>
        !p.textTrimmed.endsWith(';') && !p.textTrimmed.endsWith('.')
      );
      if (noSemicolon.length > goodListItems.length * 0.3) {
        this._add('info', 'list-no-semicolon', 'п. 2.3.5',
          `В ${noSemicolon.length} пунктах перечисления нет точки с запятой в конце`,
          'Перечисления',
          'Элементы перечисления заканчиваются точкой с запятой «;», последний — точкой «.» (п. 2.3.5)',
          this._page(noSemicolon[0]));
      }
    }

    // ── 9. Иностранные слова в основном тексте (п. 2.3.1) ───────────────────
    const engWordPattern = /\b[a-zA-Z]{4,}\b/g;
    const engWords = (fullText.match(engWordPattern) || []).filter(w => !/^[A-Z]+$/.test(w));
    if (engWords.length > 50) {
      this._add('info', 'many-foreign-words', 'п. 2.3.1',
        `Обнаружено значительное количество иностранных слов (~${engWords.length}). Проверьте наличие русских аналогов`,
        'Текст',
        'Запрещается применять иностранные термины при наличии равнозначных слов в русском языке (п. 2.3.1)');
    }

    // ── 11. Знак «№» без числа (п. 2.3.11) ─────────────────────────────────
    const badNumSign = fullText.match(/№(?!\s*\d)/g) || [];
    if (badNumSign.length > 0) {
      this._add('info', 'bad-num-sign', 'п. 2.3.11',
        `Знак «№» без числового значения (${badNumSign.length} случай). Следует писать словом «номер»`,
        'Текст',
        'В тексте (без числа) знак № пишется словом: «номер» (п. 2.3.11)');
    }

    // ── 12. Знак «−» (минус) перед числом в тексте (п. 2.3.11) ─────────────
    // «–25 °С» → должно быть «минус 25 °С»
    const badMinusSign = paras.filter(p => /[–—-]\s*\d+\s*(°|градус)/i.test(p.textTrimmed));
    if (badMinusSign.length > 0) {
      this._add('info', 'bad-minus-sign', 'п. 2.3.11',
        `Знак «−» (минус) перед числом с градусами — следует писать словом «минус». Пример: «${badMinusSign[0].textTrimmed.slice(0, 60)}»`,
        'Текст',
        'Пишите: «минус 25 °С», не «−25 °С» (п. 2.3.11)',
        this._page(badMinusSign[0]));
    }

    // ── 13. Цифры 1–9 без единиц измерений должны писаться словами (п. 2.3.12) ─
    // Ищем отдельно стоящие цифры 1-9 в русском тексте (не рядом с единицами и не в номерах)
    const digitWordParas = [];
    const digitPattern = /(?<![0-9А-Яа-яA-Za-z/.])\b([1-9])\b(?!\s*(мм|см|м|км|кг|г|т|пт|кВт|Вт|МГц|ГГц|кГц|Гц|В\b|А\b|мА|нс|мкс|мс|кОм|МОм|нФ|мкФ|пФ|%|°|шт))/g;
    for (const para of paras) {
      const t = para.textTrimmed;
      // Пропускаем подписи рисунков/таблиц, перечисления, ссылки
      if (/^(Рисунок|Таблица|где|Рис\.)/i.test(t)) continue;
      if (/^\[?\d+\]?$/.test(t.trim())) continue;
      const matches = [...t.matchAll(digitPattern)];
      if (matches.length > 0) {
        digitWordParas.push({ para, count: matches.length, examples: matches.slice(0, 3).map(m => m[0]) });
      }
    }
    if (digitWordParas.length > 3) {
      const ex = digitWordParas[0];
      this._add('info', 'digit-should-be-word', 'п. 2.3.12',
        `Числа 1–9 без единиц измерений следует писать словами. Найдено в ${digitWordParas.length} абзацах. Пример: «${ex.para.textTrimmed.slice(0, 80)}»`,
        'Текст',
        'Пишите: «три эксперимента», «пять объектов» и т.д. вместо цифр 1–9 без единиц (п. 2.3.12)',
        this._page(ex.para));
    }

    // ── 14. Вводная фраза перечисления не должна обрываться на предлоге (п. 2.3.10) ──
    const badListIntro = [];
    for (let i = 0; i < paras.length - 1; i++) {
      const t = paras[i].textTrimmed;
      // Вводная фраза заканчивается предлогом или союзом перед перечислением
      if (/\b(из|на|то|как|что|для|от|по|в|с|к|о|по)\s*:?\s*$/.test(t.toLowerCase()) &&
          paras[i + 1]?.textTrimmed.startsWith('–')) {
        badListIntro.push(paras[i]);
      }
    }
    if (badListIntro.length > 0) {
      const ex = badListIntro[0];
      this._add('warning', 'bad-list-intro', 'п. 2.3.10',
        `Вводная фраза перечисления обрывается на предлоге: «${ex.textTrimmed.slice(-40)}»`,
        `Текст: «${ex.textTrimmed.slice(0, 50)}»`,
        'Нельзя обрывать вводную фразу перед перечислением на предлогах «из», «на», «то», «как» и т.д. (п. 2.3.10)',
        this._page(ex));
    }

    // ── 10. Проверка шрифта и размера на уровне run-ов ──────────────────────
    // Ищем отдельные runs с явно другим шрифтом/размером (не через наследование)
    const runIssues = [];
    for (const para of paras.slice(0, 60)) {
      for (const run of para.runs) {
        if (!run.rPr) continue;
        const rpr = run.rPr;
        if (rpr.font && !rpr.font.toLowerCase().includes('times') && run.text.trim().length > 3) {
          runIssues.push({ type: 'font', val: rpr.font, text: run.text.trim().slice(0, 30), page: this._page(para) });
        }
        if (rpr.size && Math.abs(rpr.size - 14) > 0.5 && run.text.trim().length > 3) {
          runIssues.push({ type: 'size', val: rpr.size, text: run.text.trim().slice(0, 30), page: this._page(para) });
        }
      }
    }
    const badFontRuns = runIssues.filter(r => r.type === 'font');
    const badSizeRuns = runIssues.filter(r => r.type === 'size');
    if (badFontRuns.length > 3) {
      const ex = badFontRuns[0];
      this._add('warning', 'inline-font-mismatch', 'п. 2.1.1',
        `В тексте обнаружены участки с шрифтом «${ex.val}» (не Times New Roman) — всего ${badFontRuns.length} фрагментов. Пример: «${ex.text}»`,
        'Основной текст',
        'Выделите проблемные фрагменты и установите Times New Roman 14 пт',
        ex.page);
    }
    if (badSizeRuns.length > 3) {
      const ex = badSizeRuns[0];
      this._add('warning', 'inline-size-mismatch', 'п. 2.1.1',
        `В тексте обнаружены участки с размером ${ex.val} пт (требуется 14 пт) — всего ${badSizeRuns.length} фрагментов. Пример: «${ex.text}»`,
        'Основной текст',
        'Выделите проблемные фрагменты и установите размер шрифта 14 пт',
        ex.page);
    }
  }

  // ============================================================
  // 11. ПРИЛОЖЕНИЯ (п. 2.7)
  // ============================================================
  _checkAppendices() {
    const paras = this.docData.paragraphs;
    const appendixParas = paras.filter(p => /^ПРИЛОЖЕНИЕ\s+[А-ЯЁA-Z]/i.test(p.textTrimmed));

    if (appendixParas.length === 0) {
      this._add('info', 'no-appendices', 'п. 2.7',
        'Приложения не обнаружены',
        'Приложения',
        'При наличии приложений оформляйте их по п. 2.7 СТП');
      return;
    }

    // Проверяем буквы приложений (нельзя: Ё, З, Й, О, Ч, Ъ, Ы, Ь)
    const forbiddenLetters = ['Ё', 'З', 'Й', 'О', 'Ч', 'Ъ', 'Ы', 'Ь'];
    let badLetterCount = 0;

    for (const para of appendixParas) {
      const match = para.textTrimmed.match(/^ПРИЛОЖЕНИЕ\s+([А-ЯЁA-Z])/i);
      if (match) {
        const letter = match[1].toUpperCase();
        if (forbiddenLetters.includes(letter)) {
          badLetterCount++;
          this._add('warning', `appendix-bad-letter-${letter}`, 'п. 2.7.2',
            `Недопустимая буква для приложения: «${letter}»`,
            `Приложение «${letter}»`,
            `Нельзя использовать буквы: ${forbiddenLetters.join(', ')} (п. 2.7.2)`,
            this._page(para));
        }
      }
    }

    // Проверяем наличие «обязательное»/«рекомендуемое»/«справочное» после ПРИЛОЖЕНИЕ
    let noTypeCount = 0;
    for (let i = 0; i < paras.length; i++) {
      if (/^ПРИЛОЖЕНИЕ\s+[А-ЯЁA-Z]/i.test(paras[i].textTrimmed)) {
        // Следующий параграф должен содержать тип приложения
        const nextTexts = paras.slice(i + 1, i + 3).map(p => p.textTrimmed.toLowerCase());
        const hasType = nextTexts.some(t =>
          t.includes('обязательное') || t.includes('рекомендуемое') || t.includes('справочное')
        );
        if (!hasType) {
          noTypeCount++;
        }
      }
    }

    if (noTypeCount > 0) {
      this._add('warning', 'appendix-no-type', 'п. 2.7.3',
        `${noTypeCount} приложений без указания типа (обязательное/рекомендуемое/справочное)`,
        'Приложения',
        'После ПРИЛОЖЕНИЕ X в скобках укажите тип: (обязательное), (рекомендуемое) или (справочное)');
    }

    // Проверяем ссылки на приложения в тексте
    const hasAppRefs = /приложени[еи]\s+[А-ЯЁ]/i.test(this.docData.fullText);
    if (!hasAppRefs && appendixParas.length > 0) {
      this._add('warning', 'appendix-no-refs', 'п. 2.7.2',
        'Не найдены ссылки на приложения в тексте',
        'Приложения',
        'В тексте должны быть ссылки на все приложения (п. 2.7.2): «(приложение А)»');
    }

    if (badLetterCount === 0 && noTypeCount === 0) {
      this._pass('appendices-ok', 'п. 2.7', `Найдено ${appendixParas.length} приложений, оформление верное ✓`);
    }
  }

  // ============================================================
  // 12. НУМЕРАЦИЯ СТРАНИЦ (п. 2.2.8)
  // ============================================================
  _checkPageNumbering() {
    const footer = this.docData.footerData;
    const header = this.docData.headerData;

    if (!footer.hasFooter && !header.hasHeader) {
      this._add('warning', 'no-page-number-field', 'п. 2.2.8',
        'Колонтитулы не обнаружены. Убедитесь, что нумерация страниц настроена',
        'Нумерация страниц',
        'Вставьте номер страницы в правом нижнем углу: Вставка → Номер страницы → Внизу страницы → Справа');
      return;
    }

    if (footer.hasPageNumber) {
      if (footer.pageNumAlignment && footer.pageNumAlignment !== 'right') {
        this._add('critical', 'page-num-not-right', 'п. 2.2.8',
          `Номер страницы не в правом нижнем углу (выравнивание: ${footer.pageNumAlignment})`,
          'Нумерация страниц',
          'Расположите номер страницы в правом нижнем углу (Выравнивание: по правому краю)');
      } else {
        this._pass('page-num-ok', 'п. 2.2.8', 'Нумерация страниц в нижнем колонтитуле ✓');
      }
    } else if (header.hasPageNumber) {
      this._add('critical', 'page-num-in-header', 'п. 2.2.8',
        'Номер страницы в верхнем колонтитуле. Требуется в правом нижнем углу',
        'Нумерация страниц',
        'Переместите номер страницы из шапки в подвал, в правый угол');
    } else {
      this._add('warning', 'no-page-number-field', 'п. 2.2.8',
        'Номер страницы в колонтитулах не обнаружен',
        'Нумерация страниц',
        'Вставьте номер страницы: Вставка → Номер страницы → Внизу страницы → Справа от полей');
    }
  }
}
