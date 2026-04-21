/**
 * App — главный контроллер приложения
 * Управляет загрузкой файлов, запуском проверок и отображением результатов
 */

class App {

  constructor() {
    this.docxFile = null;
    this.assignmentFile = null;
    this.drawingFiles = [];
    this.findings = [];
    this.docData = null;

    this.parser = new DocxParser();
    this.checker = new StpChecker();
    this.aiChecker = new AiChecker();

    this._initElements();
    this._initEvents();
    this._loadSettings();
  }

  // ============================================================
  // ИНИЦИАЛИЗАЦИЯ
  // ============================================================
  _initElements() {
    this.dropZone       = document.getElementById('drop-zone');
    this.docxInput      = document.getElementById('docx-input');
    this.fileInfo       = document.getElementById('file-info');
    this.fileName       = document.getElementById('file-name');
    this.fileSize       = document.getElementById('file-size');
    this.removeFileBtn  = document.getElementById('remove-file');

    this.drawingsZone   = document.getElementById('drawings-drop-zone');
    this.drawingsInput  = document.getElementById('drawings-input');
    this.drawingsList   = document.getElementById('drawings-list');

    this.assignmentZone = document.getElementById('assignment-drop-zone');
    this.assignmentInput = document.getElementById('assignment-input');
    this.assignmentInfo  = document.getElementById('assignment-info');
    this.assFileName     = document.getElementById('assignment-file-name');
    this.assFileSize     = document.getElementById('assignment-file-size');
    this.removeAssBtn    = document.getElementById('remove-assignment');

    this.useAiChk       = document.getElementById('use-ai');
    this.aiOptions      = document.getElementById('ai-options');
    this.apiKeyInput    = document.getElementById('api-key');
    this.apiKeyLabel    = document.getElementById('api-key-label');
    this.apiKeyLink     = document.getElementById('api-key-link');
    this.aiNote         = document.getElementById('ai-note');
    this.aiModelSelect  = document.getElementById('ai-model');
    this.aiProviderTabs = document.querySelectorAll('.ai-provider-tab');
    this.toggleKeyBtn   = document.getElementById('toggle-key');
    this.useVisualChk   = document.getElementById('use-visual');
    this._aiProvider    = 'gemini'; // текущий провайдер

    this.checkBtn       = document.getElementById('check-btn');
    this.checkHint      = document.getElementById('check-hint');

    this.progressSection  = document.getElementById('progress-section');
    this.progressBar      = document.getElementById('progress-bar');
    this.progressText     = document.getElementById('progress-text');
    this.progressSteps    = document.getElementById('progress-steps');

    this.resultsSection = document.getElementById('results-section');
    this.filterTabs     = document.querySelectorAll('.filter-tab');
    this.findingsContainer = document.getElementById('findings-container');
    this.previewSection = document.getElementById('preview-section');
    this.docPreview     = document.getElementById('doc-preview');

    this.exportPdfBtn   = document.getElementById('export-pdf-btn');
    this.copyBtn        = document.getElementById('copy-btn');
    this.recheckBtn     = document.getElementById('recheck-btn');

    this.themeToggle    = document.getElementById('theme-toggle');
    this.helpBtn        = document.getElementById('help-btn');
    this.helpModal      = document.getElementById('help-modal');
    this.modalClose     = document.getElementById('modal-close');

    // Чекбоксы параметров
    this.optChecks = {
      format:    document.getElementById('check-format'),
      structure: document.getElementById('check-structure'),
      headings:  document.getElementById('check-headings'),
      figures:   document.getElementById('check-figures'),
      formulas:  document.getElementById('check-formulas'),
      refs:      document.getElementById('check-refs'),
      text:      document.getElementById('check-text'),
    };
  }

  // ============================================================
  _initEvents() {
    // --- DOCX Drop Zone ---
    this.dropZone.addEventListener('click', () => this.docxInput.click());
    this.dropZone.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') this.docxInput.click(); });
    this.docxInput.addEventListener('change', (e) => { if (e.target.files[0]) this._setDocxFile(e.target.files[0]); });

    this.dropZone.addEventListener('dragover',  (e) => { e.preventDefault(); this.dropZone.classList.add('drag-over'); });
    this.dropZone.addEventListener('dragleave', () => this.dropZone.classList.remove('drag-over'));
    this.dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      this.dropZone.classList.remove('drag-over');
      const file = e.dataTransfer.files[0];
      if (file && file.name.endsWith('.docx')) this._setDocxFile(file);
      else this._toast('Пожалуйста, загрузите файл .docx', 'error');
    });

    // Удалить файл
    this.removeFileBtn.addEventListener('click', () => this._clearDocxFile());

    // --- Drawings ---
    this.drawingsZone.addEventListener('click', () => this.drawingsInput.click());
    this.drawingsInput.addEventListener('change', (e) => {
      for (const f of e.target.files) this._addDrawing(f);
    });
    this.drawingsZone.addEventListener('dragover',  (e) => { e.preventDefault(); this.drawingsZone.classList.add('drag-over'); });
    this.drawingsZone.addEventListener('dragleave', () => this.drawingsZone.classList.remove('drag-over'));
    this.drawingsZone.addEventListener('drop', (e) => {
      e.preventDefault();
      this.drawingsZone.classList.remove('drag-over');
      for (const f of e.dataTransfer.files) this._addDrawing(f);
    });

    // --- Assignment ---
    if (this.assignmentZone) {
      this.assignmentZone.addEventListener('click', () => this.assignmentInput.click());
      this.assignmentInput.addEventListener('change', (e) => {
        if (e.target.files[0]) this._setAssignment(e.target.files[0]);
      });
      this.assignmentZone.addEventListener('dragover',  (e) => { e.preventDefault(); this.assignmentZone.classList.add('drag-over'); });
      this.assignmentZone.addEventListener('dragleave', () => this.assignmentZone.classList.remove('drag-over'));
      this.assignmentZone.addEventListener('drop', (e) => {
        e.preventDefault();
        this.assignmentZone.classList.remove('drag-over');
        if (e.dataTransfer.files[0]) this._setAssignment(e.dataTransfer.files[0]);
      });
      if (this.removeAssBtn) {
        this.removeAssBtn.addEventListener('click', () => this._removeAssignment());
      }
    }

    // --- AI toggle ---
    this.useAiChk.addEventListener('change', () => {
      this.aiOptions.style.display = this.useAiChk.checked ? '' : 'none';
    });

    // --- AI provider tabs ---
    this.aiProviderTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        this.aiProviderTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        this._aiProvider = tab.dataset.provider;
        this._updateAiProviderUI(tab.dataset.provider);
      });
    });

    // --- API key toggle ---
    this.toggleKeyBtn.addEventListener('click', () => {
      const isPass = this.apiKeyInput.type === 'password';
      this.apiKeyInput.type = isPass ? 'text' : 'password';
      this.toggleKeyBtn.textContent = isPass ? '🙈' : '👁';
    });

    // Сохраняем API ключ в localStorage при вводе (ключ хранится отдельно на провайдера)
    this.apiKeyInput.addEventListener('input', () => {
      if (this.apiKeyInput.value.length > 10) {
        localStorage.setItem(`ai_key_${this._aiProvider}`, btoa(this.apiKeyInput.value));
      }
    });

    // --- Check button ---
    this.checkBtn.addEventListener('click', () => this._startCheck());

    // --- Filter tabs ---
    this.filterTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        this.filterTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        ReportGenerator.renderFindings(this.findings, tab.dataset.filter);
      });
    });

    // --- Export buttons ---
    this.exportPdfBtn.addEventListener('click', () => ReportGenerator.exportToPDF());
    this.copyBtn.addEventListener('click', async () => {
      const ok = await ReportGenerator.copyToClipboard(this.findings, { fileName: this.docxFile?.name });
      this._toast(ok ? '✅ Отчёт скопирован в буфер' : 'Ошибка копирования', ok ? 'success' : 'error');
    });
    this.recheckBtn.addEventListener('click', () => this._resetResults());

    // --- Theme ---
    this.themeToggle.addEventListener('click', () => {
      const isDark = document.body.getAttribute('data-theme') === 'dark';
      document.body.setAttribute('data-theme', isDark ? '' : 'dark');
      this.themeToggle.textContent = isDark ? '🌙' : '☀️';
      localStorage.setItem('theme', isDark ? 'light' : 'dark');
    });

    // --- Help modal ---
    this.helpBtn.addEventListener('click', () => this.helpModal.style.display = 'flex');
    this.modalClose.addEventListener('click', () => this.helpModal.style.display = 'none');
    this.helpModal.addEventListener('click', (e) => { if (e.target === this.helpModal) this.helpModal.style.display = 'none'; });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') this.helpModal.style.display = 'none'; });

    // Paste DOCX from clipboard (Ctrl+V)
    document.addEventListener('paste', (e) => {
      const items = e.clipboardData.items;
      for (const item of items) {
        if (item.kind === 'file' && (item.type.includes('word') || item.getAsFile()?.name?.endsWith('.docx'))) {
          const f = item.getAsFile();
          if (f) this._setDocxFile(f);
        }
      }
    });

    // --- Save Options to LocalStorage ---
    Object.keys(this.optChecks).forEach(key => {
      if (this.optChecks[key]) {
        this.optChecks[key].addEventListener('change', () => {
          try {
            const settings = JSON.parse(localStorage.getItem('stp_check_options') || '{}');
            settings[key] = this.optChecks[key].checked;
            localStorage.setItem('stp_check_options', JSON.stringify(settings));
          } catch (e) { /* ignore */ }
        });
      }
    });

    // Save AI visual check option
    if (this.useVisualChk) {
      this.useVisualChk.addEventListener('change', () => {
         localStorage.setItem('stp_ai_visual', this.useVisualChk.checked);
      });
    }
  }

  // ============================================================
  _loadSettings() {
    // Тема
    const savedTheme = localStorage.getItem('theme') || 'light';
    if (savedTheme === 'dark') {
      document.body.setAttribute('data-theme', 'dark');
      this.themeToggle.textContent = '☀️';
    }

    // API ключ (для текущего провайдера)
    try {
      const savedProvider = localStorage.getItem('ai_provider') || 'gemini';
      this._aiProvider = savedProvider;
      // Активируем нужную вкладку
      this.aiProviderTabs.forEach(t => {
        t.classList.toggle('active', t.dataset.provider === savedProvider);
      });
      this._updateAiProviderUI(savedProvider);

      const enc = localStorage.getItem(`ai_key_${savedProvider}`);
      if (!enc) {
        // Fallback: старый ключ Gemini
        const oldEnc = localStorage.getItem('gemini_api_key');
        if (oldEnc) this.apiKeyInput.value = atob(oldEnc);
      } else if (enc.length > 5) {
        this.apiKeyInput.value = atob(enc);
      }
    } catch (e) { /* ignore */ }

    // Загрузка состояния чекбоксов
    try {
      const settings = JSON.parse(localStorage.getItem('stp_check_options'));
      if (settings) {
        Object.keys(this.optChecks).forEach(key => {
          if (settings[key] !== undefined && this.optChecks[key]) {
             this.optChecks[key].checked = !!settings[key];
          }
        });
      }
      // Загрузка состояния визуальной проверки
      if (this.useVisualChk) {
         const vis = localStorage.getItem('stp_ai_visual');
         if (vis !== null) this.useVisualChk.checked = (vis === 'true');
      }
    } catch (e) { /* ignore */ }
  }

  // ============================================================
  // ПЕРЕКЛЮЧЕНИЕ ПРОВАЙДЕРА ИИ
  // ============================================================
  _updateAiProviderUI(provider) {
    const cfg = AiChecker.PROVIDERS[provider];
    if (!cfg) return;

    localStorage.setItem('ai_provider', provider);

    // Обновляем подпись и placeholder
    if (this.apiKeyLabel) this.apiKeyLabel.textContent = `API ключ ${cfg.name}:`;
    if (this.apiKeyInput) this.apiKeyInput.placeholder = cfg.keyPlaceholder;

    // Обновляем ссылку на получение ключа
    if (this.apiKeyLink) {
      this.apiKeyLink.href = cfg.keyLink;
      this.apiKeyLink.textContent = `🔑 ${cfg.keyLinkText} →`;
    }

    // Обновляем заметку
    if (this.aiNote) this.aiNote.textContent = cfg.note;

    // Обновляем модели в select
    if (this.aiModelSelect) {
      this.aiModelSelect.innerHTML = cfg.models.map(m =>
        `<option value="${m.id}">${m.label}</option>`
      ).join('');
    }

    // Показываем/скрываем визуальную проверку (только для провайдеров с Vision)
    const visualOpt = document.getElementById('visual-check-option');
    if (visualOpt) visualOpt.style.display = cfg.supportsVision ? '' : 'none';

    // Восстанавливаем ключ для этого провайдера
    try {
      const enc = localStorage.getItem(`ai_key_${provider}`);
      if (enc && enc.length > 5 && this.apiKeyInput) {
        this.apiKeyInput.value = atob(enc);
      } else if (this.apiKeyInput && provider !== 'gemini') {
        this.apiKeyInput.value = '';
      }
    } catch (e) { /* ignore */ }
  }

  // ============================================================
  // УПРАВЛЕНИЕ ФАЙЛАМИ
  // ============================================================
  _setDocxFile(file) {
    if (!file.name.toLowerCase().endsWith('.docx')) {
      this._toast('Пожалуйста, выберите файл .docx', 'error');
      return;
    }

    this.docxFile = file;
    this.fileName.textContent = file.name;
    this.fileSize.textContent = this._formatSize(file.size);

    this.dropZone.style.display = 'none';
    this.fileInfo.style.display = 'flex';

    this.checkBtn.disabled = false;
    this.checkHint.textContent = `Готово к проверке: ${file.name}`;

    this._toast(`Файл загружен: ${file.name}`, 'success');
  }

  _clearDocxFile() {
    this.docxFile = null;
    this.docxInput.value = '';
    this.dropZone.style.display = '';
    this.fileInfo.style.display = 'none';
    this.checkBtn.disabled = true;
    this.checkHint.textContent = 'Загрузите файл для начала проверки';
    this._resetResults(true);
  }

  _addDrawing(file) {
    const allowed = ['image/jpeg', 'image/png', 'image/jpg', 'application/pdf'];
    if (!allowed.includes(file.type) && !file.name.match(/\.(pdf|png|jpg|jpeg)$/i)) {
      this._toast(`Неподдерживаемый формат: ${file.name}`, 'error');
      return;
    }
    if (this.drawingFiles.some(f => f.name === file.name)) return;

    this.drawingFiles.push(file);
    this._renderDrawingsList();
    this._toast(`Добавлен чертёж: ${file.name}`, 'info');
  }

  _removeDrawing(name) {
    this.drawingFiles = this.drawingFiles.filter(f => f.name !== name);
    this._renderDrawingsList();
  }

  _renderDrawingsList() {
    this.drawingsList.innerHTML = this.drawingFiles.map(f => `
      <div class="drawing-chip">
        ${f.type.includes('pdf') ? '📐' : '🖼'} ${this._escHtml(f.name)}
        <button class="drawing-chip-remove" onclick="app._removeDrawing('${this._escHtml(f.name)}')">✕</button>
      </div>
    `).join('');
  }

  _setAssignment(file) {
    if (!file.name.toLowerCase().endsWith('.docx')) {
      this._toast(`Требуется файл .docx (загружен ${file.name})`, 'error');
      return;
    }
    this.assignmentFile = file;
    
    if (this.assFileName) this.assFileName.textContent = file.name;
    if (this.assFileSize) this.assFileSize.textContent = this._formatSize(file.size);
    
    if (this.assignmentZone) this.assignmentZone.style.display = 'none';
    if (this.assignmentInfo) this.assignmentInfo.style.display = 'flex';

    this._toast(`Задание загружено: ${file.name}`, 'success');
  }

  _removeAssignment() {
    this.assignmentFile = null;
    this.assignmentInput.value = '';
    
    if (this.assignmentZone) this.assignmentZone.style.display = 'flex';
    if (this.assignmentInfo) this.assignmentInfo.style.display = 'none';
  }

  // ============================================================
  // ЗАПУСК ПРОВЕРКИ
  // ============================================================
  async _startCheck() {
    if (!this.docxFile) return;

    this._showProgress();
    this.findings = [];

    const useAi     = this.useAiChk.checked;
    const useVisual = this.useVisualChk?.checked || false;
    const apiKey    = this.apiKeyInput.value.trim();
    const aiProvider = this._aiProvider || 'gemini';
    const aiModel   = this.aiModelSelect?.value || 'gemini-2.0-flash';
    const hasDrawings = this.drawingFiles.length > 0;

    const providerName = AiChecker.PROVIDERS[aiProvider]?.name || aiProvider;
    const steps = [
      { id: 'parse',     label: 'Разбор DOCX файла' },
      { id: 'format',    label: 'Проверка форматирования (поля, шрифт, отступы)' },
      { id: 'structure', label: 'Проверка структуры документа' },
      { id: 'headings',  label: 'Проверка заголовков и содержания' },
      { id: 'figures',   label: 'Проверка рисунков и таблиц' },
      { id: 'refs',      label: 'Проверка списка источников' },
      ...(this.assignmentFile ? [{ id: 'assignment', label: 'Сверка пояснительной с Заданием' }] : []),
      ...(useAi ? [{ id: 'ai',  label: `ИИ-анализ (${providerName} / ${aiModel})` }] : []),
      ...(useVisual ? [{ id: 'visual', label: 'Визуальная проверка страниц' }] : []),
      ...(hasDrawings ? [{ id: 'drawings', label: 'Проверка чертежей' }] : []),
      { id: 'report',    label: 'Формирование отчёта' },
    ];

    this._renderProgressSteps(steps);

    try {
      // --- Шаг 1: Парсинг DOCX ---
      this._setStep('parse', 'running', 5);
      this.docData = await this.parser.parse(this.docxFile);
      this._setStep('parse', 'done', 15);

      // --- Шаг 2–6: Программная проверка ---
      const opts = {
        checkFormat:    this.optChecks.format?.checked ?? true,
        checkStructure: this.optChecks.structure?.checked ?? true,
        checkHeadings:  this.optChecks.headings?.checked ?? true,
        checkFigures:   this.optChecks.figures?.checked ?? true,
        checkFormulas:  this.optChecks.formulas?.checked ?? true,
        checkRefs:      this.optChecks.refs?.checked ?? true,
        checkText:      this.optChecks.text?.checked ?? true,
      };

      this._setStep('format', 'running', 20);
      this.findings = this.checker.check(this.docData, opts);
      this._setStep('format', 'done', 50);
      this._setStep('structure', 'done', 55);
      this._setStep('headings',  'done', 60);
      this._setStep('figures',   'done', 65);
      this._setStep('refs',      'done', 70);

      // --- Шаг 7: Сверка с Заданием (без ИИ) ---
      if (this.assignmentFile) {
        this._setStep('assignment', 'running', 71);
        try {
          const assignmentDocData = await this.parser.parse(this.assignmentFile);
          const assFindings = this.checker.checkAssignmentCompliance(this.docData, assignmentDocData);
          
          // Проверяем, были ли найдены ошибки / предупреждения
          const hasErrors = assFindings.some(f => f.severity === 'warning' || f.severity === 'critical');
          
          if (!hasErrors) {
             assFindings.unshift({
                severity: 'info',
                id: 'assignment-success',
                section: 'Сверка с Заданием',
                description: 'Сверка с Заданием прошла успешно! Тема, обязательные разделы и упомянутые чертежи строго соблюдены в тексте записки.',
                location: 'Задание',
                recommendation: 'Ошибок или несоответствий не найдено.',
                source: 'programmatic'
             });
          } else {
             assFindings.unshift({
                severity: 'info',
                id: 'assignment-executed',
                section: 'Сверка с Заданием',
                description: 'Алгоритм успешно проанализировал Задание на проектирование и выявил некоторые расхождения (см. Ошибки/Предупреждения ниже).',
                location: 'Задание',
                recommendation: 'Обязательно устраните критические ошибки, связанные с заданием.',
                source: 'programmatic'
             });
          }
          
          this.findings.push(...assFindings);
          this._setStep('assignment', 'done', 72);
        } catch (assErr) {
          this._setStep('assignment', 'done', 72);
          this._toast(`Ошибка проверки задания: ${assErr.message}`, 'error');
        }
      }

      // --- Шаг 8: ИИ анализ ---
      if (useAi && apiKey) {
        this._setStep('ai', 'running', 75);
        try {
          const aiFindings = await this.aiChecker.analyzeDocument(apiKey, this.docData, aiProvider, aiModel);
          this.findings.push(...aiFindings);
          this._setStep('ai', 'done', 82);
        } catch (aiErr) {
          this._setStep('ai', 'done', 82);
          const friendlyMsg = this._translateAiError(aiErr.message);
          this._toast(`ИИ-ошибка: ${friendlyMsg}`, 'error');
          this.findings.push({
            severity: 'info',
            id: 'ai-error',
            section: 'ИИ-анализ',
            description: `ИИ-проверка не выполнена: ${friendlyMsg}`,
            location: 'API Провайдер',
            recommendation: 'Если вы используете бесплатный Gemini — подождите 1 минуту. Если платный OpenAI/DeepSeek — проверьте баланс.',
            source: 'ai'
          });
        }
      }

      // --- Шаг 9: Визуальная проверка ---
      if (useVisual && apiKey && this.docData.htmlContent) {
        this._setStep('visual', 'running', 84);
        await this._sleep(4000); // Пауза для обхода квот (15 RPM)
        try {
          const base64 = await AiChecker.captureDocumentPage(this.docData.htmlContent);
          if (base64) {
            const visFindings = await this.aiChecker.analyzeVisual(apiKey, base64, 1, aiProvider, aiModel);
            this.findings.push(...visFindings);
          }
          this._setStep('visual', 'done', 88);
        } catch (visErr) {
          this._setStep('visual', 'done', 88);
          const friendlyMsg = this._translateAiError(visErr.message);
          this._toast(`Визуальная проверка: ${friendlyMsg}`, 'error');
        }
      }

      // --- Шаг 9: Чертежи ---
      if (hasDrawings && apiKey) {
        this._setStep('drawings', 'running', 88);
        for (const drFile of this.drawingFiles) {
          try {
            let base64;
            if (drFile.type === 'application/pdf') {
              base64 = await AiChecker.pdfFirstPageToBase64(drFile);
            } else {
              base64 = await AiChecker.fileToBase64(drFile);
            }
            await this._sleep(4000); // Пауза между чертежами
            const drFindings = await this.aiChecker.analyzeDrawing(apiKey, base64, drFile.name, aiProvider, aiModel);
            // Добавляем метку чертежа
            drFindings.forEach(f => f.location = `Чертёж: ${drFile.name}`);
            this.findings.push(...drFindings);
          } catch (e) {
            this._toast(`Ошибка проверки чертежа: ${drFile.name}`, 'error');
          }
        }
        this._setStep('drawings', 'done', 95);
      }

      // --- Финал: Рендеринг отчёта ---
      this._setStep('report', 'running', 97);

      // Предпросмотр HTML — постраничный вид
      if (this.docData.htmlContent) {
        this.previewSection.style.display = '';
        this._renderDocumentPreview(this.docData.htmlContent);
      }

      this._setProgress(100, 'Проверка завершена!');
      this._setStep('report', 'done', 100);

      // Небольшая задержка для UX
      await this._sleep(500);

      this._showResults();

    } catch (err) {
      this._showError(err.message || String(err));
    }
  }

  // ============================================================
  // ПРОГРЕСС
  // ============================================================
  _showProgress() {
    document.getElementById('upload-section').style.opacity = '0.6';
    document.getElementById('options-section').style.opacity = '0.6';
    this.progressSection.style.display = '';
    this.resultsSection.style.display = 'none';
    this.checkBtn.disabled = true;
  }

  _setProgress(pct, text) {
    this.progressBar.style.width = pct + '%';
    this.progressText.textContent = text;
  }

  _renderProgressSteps(steps) {
    this.progressSteps.innerHTML = steps.map(s => `
      <div class="progress-step" id="step-${s.id}">
        <div class="step-status">○</div>
        <span>${s.label}</span>
      </div>
    `).join('');
  }

  _setStep(id, status, pct) {
    const el = document.getElementById(`step-${id}`);
    if (!el) return;
    el.className = `progress-step ${status}`;
    el.querySelector('.step-status').textContent = status === 'done' ? '✓' : '↻';
    this._setProgress(pct, el.querySelector('span').textContent + '...');
  }

  // ============================================================
  // ПРЕДПРОСМОТР ДОКУМЕНТА (постраничный вид, имитация A4)
  // ============================================================
  _renderDocumentPreview(htmlContent) {
    const CHARS_PER_PAGE = 2500; // примерно ~40 строк × ~65 символов

    // Парсим HTML-контент в элементы
    const temp = document.createElement('div');
    temp.innerHTML = htmlContent;
    const elements = Array.from(temp.children);

    // Разбиваем на страницы по объёму текста
    const pages = [];
    let curPage = [];
    let curChars = 0;

    for (const el of elements) {
      const isHeading = /^H[1-6]$/.test(el.tagName);
      const charWeight = el.textContent.length + (isHeading ? 200 : 0);

      // Заголовок 1-го уровня всегда начинает новую страницу
      if (isHeading && el.tagName === 'H1' && curPage.length > 0) {
        pages.push(curPage.splice(0));
        curChars = 0;
      } else if (curChars + charWeight > CHARS_PER_PAGE && curPage.length > 0) {
        pages.push(curPage.splice(0));
        curChars = 0;
      }

      curPage.push(el.outerHTML);
      curChars += charWeight;
    }
    if (curPage.length > 0) pages.push(curPage);

    const totalPages = pages.length;

    this.docPreview.innerHTML = `
      <div class="preview-viewer">
        ${pages.map((pageEls, i) => `
          <div class="preview-page" id="preview-page-${i + 1}">
            <div class="preview-page-inner">${pageEls.join('')}</div>
            <div class="preview-page-num">${i + 1} / ${totalPages}</div>
          </div>`).join('')}
      </div>`;
  }

  // ============================================================
  // НАВИГАЦИЯ К СТРАНИЦЕ В ПРЕДПРОСМОТРЕ
  // ============================================================
  navigateToPage(pageNum) {
    // Раскрываем секцию предпросмотра
    if (this.previewSection) {
      this.previewSection.open = true;
      this.previewSection.style.display = '';
    }

    const pages = document.querySelectorAll('.preview-page');
    if (!pages || pages.length === 0) {
      this._toast('Предпросмотр текста не загружен.', 'info');
      return;
    }

    // Ограничиваем запрошенную страницу фактически существующим количеством,
    // так как номера рассчитываются приблизительно. 
    let targetNum = parseInt(pageNum, 10);
    if (isNaN(targetNum) || targetNum < 1) targetNum = 1;
    if (targetNum > pages.length) targetNum = pages.length;

    const pageEl = document.getElementById(`preview-page-${targetNum}`);
    if (!pageEl) return;

    // Прокручиваем к странице
    pageEl.scrollIntoView({ behavior: 'smooth', block: 'center' });

    // Мигающая подсветка
    pageEl.classList.remove('preview-page-highlight');
    void pageEl.offsetWidth; // reflow
    pageEl.classList.add('preview-page-highlight');
    setTimeout(() => pageEl.classList.remove('preview-page-highlight'), 2200);
  }

  // ============================================================
  // РЕЗУЛЬТАТЫ
  // ============================================================
  _showResults() {
    this.progressSection.style.display = 'none';
    document.getElementById('upload-section').style.opacity = '1';
    document.getElementById('options-section').style.opacity = '1';
    this.checkBtn.disabled = false;

    ReportGenerator.updateSummary(this.findings);
    ReportGenerator.renderFindings(this.findings, 'all');

    // Сброс активного фильтра
    this.filterTabs.forEach(t => t.classList.remove('active'));
    this.filterTabs[0]?.classList.add('active');

    this.resultsSection.style.display = '';
    this.resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

    const { score, counts } = ReportGenerator.calculateScore(this.findings);
    this._toast(
      `Проверка завершена: ${counts.critical} критических, ${counts.warning} предупреждений. Оценка: ${score}/10`,
      counts.critical > 0 ? 'error' : counts.warning > 3 ? 'info' : 'success'
    );
  }

  _resetResults(clearProgress = false) {
    this.findings = [];
    this.docData = null;
    this.resultsSection.style.display = 'none';
    if (clearProgress) this.progressSection.style.display = 'none';
    document.getElementById('upload-section').style.opacity = '1';
    document.getElementById('options-section').style.opacity = '1';
    if (this.docxFile) this.checkBtn.disabled = false;
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  _showError(msg) {
    this._setProgress(0, '');
    this.progressSection.style.display = 'none';
    document.getElementById('upload-section').style.opacity = '1';
    document.getElementById('options-section').style.opacity = '1';
    this.checkBtn.disabled = false;
    this._toast(`Ошибка: ${msg}`, 'error');
    console.error('Check error:', msg);
  }

  // ============================================================
  // УТИЛИТЫ
  // ============================================================
  _toast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
      ${type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️'}
      <span>${this._escHtml(message)}</span>
    `;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.animation = 'fadeOut 0.3s ease forwards';
      setTimeout(() => toast.remove(), 350);
    }, 3000);
  }

  _escHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  _formatSize(bytes) {
    if (bytes < 1024) return bytes + ' Б';
    if (bytes < 1048576) return Math.round(bytes / 1024) + ' КБ';
    return (bytes / 1048576).toFixed(1) + ' МБ';
  }

  _sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  _translateAiError(msg) {
    if (!msg) return 'Неизвестная ошибка ИИ';
    const lowMsg = msg.toLowerCase();

    if (lowMsg.includes('insufficient balance') || lowMsg.includes('insufficient_quota')) {
      return 'Недостаточно средств на балансе API. Пополните счет или переключитесь на бесплатный Gemini.';
    }
    if (lowMsg.includes('quota exceeded') || lowMsg.includes('rate limit') || lowMsg.includes('429')) {
      return 'Превышен лимит запросов (Quota Exceeded). Подождите 1-5 минут или попробуйте другой API-ключ.';
    }
    if (lowMsg.includes('invalid api key') || lowMsg.includes('api key not found')) {
      return 'Неверный API-ключ. Проверьте правильность ввода в настройках.';
    }
    if (lowMsg.includes('model not found') || lowMsg.includes('not supported')) {
      return 'Выбранная модель недоступна для вашего ключа.';
    }
    if (lowMsg.includes('safety') || lowMsg.includes('blocked')) {
      return 'Запрос заблокирован фильтром безопасности ИИ (попробуйте другой фрагмент текста).';
    }

    return msg; // Возвращаем оригинал если не узнали
  }
}

// ============================================================
// ИНИЦИАЛИЗАЦИЯ
// ============================================================
let app;
document.addEventListener('DOMContentLoaded', () => {
  try {
    app = new App();
    console.log('✅ БГУИР Нормоконтроль v1.0 инициализирован');
  } catch (e) {
    console.error('Ошибка инициализации:', e);
    document.body.innerHTML = `
      <div style="padding:40px;text-align:center;font-family:sans-serif">
        <h2>Ошибка загрузки приложения</h2>
        <p>${e.message}</p>
        <p>Откройте консоль браузера (F12) для подробностей</p>
      </div>
    `;
  }
});
