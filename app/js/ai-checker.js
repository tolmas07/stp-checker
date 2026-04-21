/**
 * AI Checker — проверка через различные AI API
 * Поддерживаемые провайдеры:
 *   - Google Gemini  (https://ai.google.dev)
 *   - DeepSeek       (https://platform.deepseek.com)
 *   - OpenAI         (https://platform.openai.com)
 */

class AiChecker {

  constructor() {
    this.maxTextLength = 12000;
  }

  // ============================================================
  // Конфигурации провайдеров и их моделей (статический справочник)
  // ============================================================
  static get PROVIDERS() {
    return {
      gemini: {
        name: 'Google Gemini',
        keyPlaceholder: 'AIzaSy...',
        keyLink: 'https://aistudio.google.com/app/apikey',
        keyLinkText: 'Получить бесплатный ключ — Google AI Studio',
        note: 'Бесплатно: 15 запросов/мин, 1 500 000 токенов/день.',
        supportsVision: true,
        models: [
          { id: 'gemini-2.0-flash',      label: 'Gemini 2.0 Flash (рекомендуется, бесплатно)' },
          { id: 'gemini-2.0-flash-lite',  label: 'Gemini 2.0 Flash Lite (быстрее)' },
          { id: 'gemini-1.5-pro-latest',  label: 'Gemini 1.5 Pro (умнее, медленнее)' },
        ]
      },
      deepseek: {
        name: 'DeepSeek',
        keyPlaceholder: 'sk-...',
        keyLink: 'https://platform.deepseek.com/api_keys',
        keyLinkText: 'Получить ключ — DeepSeek Platform',
        note: 'Платный API. DeepSeek-V3 — мощная модель по низкой цене.',
        supportsVision: false,
        models: [
          { id: 'deepseek-chat',     label: 'DeepSeek-V3 Chat (рекомендуется)' },
          { id: 'deepseek-reasoner', label: 'DeepSeek-R1 Reasoner (глубокий анализ)' },
        ]
      },
      openai: {
        name: 'OpenAI',
        keyPlaceholder: 'sk-...',
        keyLink: 'https://platform.openai.com/api-keys',
        keyLinkText: 'Получить ключ — OpenAI Platform',
        note: 'Платный API. GPT-4o mini — оптимальный выбор по цене/качеству.',
        supportsVision: true,
        models: [
          { id: 'gpt-4o-mini',    label: 'GPT-4o mini (рекомендуется, дёшево)' },
          { id: 'gpt-4o',         label: 'GPT-4o (лучшее качество)' },
          { id: 'gpt-3.5-turbo',  label: 'GPT-3.5 Turbo (дёшевле)' },
        ]
      }
    };
  }

  // ============================================================
  // Текстовый анализ документа
  // ============================================================
  async analyzeDocument(apiKey, docData, provider = 'gemini', model = 'gemini-2.0-flash') {
    if (!apiKey) throw new Error('API ключ не указан');
    const docText = this._prepareText(docData.fullText);
    const prompt  = this._buildPrompt(docText, docData);
    const rawText = await this._callProvider(apiKey, provider, model, prompt);
    return this._parseJsonResponse(rawText, 'ИИ-анализ');
  }

  // ============================================================
  // Визуальная проверка страниц (Gemini Vision / OpenAI Vision)
  // ============================================================
  async analyzeVisual(apiKey, imageBase64, pageNum = 1, provider = 'gemini', model = 'gemini-2.0-flash') {
    if (!apiKey) throw new Error('API ключ не указан');
    const prompt = this._buildVisualPrompt(pageNum);
    let rawText;

    if (provider === 'gemini') {
      const body = {
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: 'image/jpeg', data: imageBase64 } }
        ]}],
        generationConfig: { temperature: 0.1, maxOutputTokens: 2000 }
      };
      rawText = await this._callGemini(apiKey, model, body);
    } else if (provider === 'openai') {
      rawText = await this._callOpenAIVision(apiKey, model, imageBase64, prompt);
    } else {
      throw new Error('Визуальная проверка поддерживается только для Gemini и OpenAI');
    }

    return this._parseJsonResponse(rawText, `Визуальная проверка стр. ${pageNum}`)
      .map(item => ({ ...item, location: `Стр. ${pageNum} (визуальная)` }));
  }

  // ============================================================
  // Анализ чертежа
  // ============================================================
  async analyzeDrawing(apiKey, imageBase64, fileName = '', provider = 'gemini', model = 'gemini-2.0-flash') {
    if (!apiKey) throw new Error('API ключ не указан');
    const prompt = `Ты — эксперт по нормоконтролю документации по ЕСКД.
Проверь данный чертёж/схему. Имя файла: ${fileName}

Проверь: рамку, основную надпись, линии, текст, общее качество.

Выведи ТОЛЬКО JSON массив:
[{"severity":"critical|warning|info|pass","description":"...","recommendation":"...или null"}]`;

    let rawText;
    if (provider === 'gemini') {
      rawText = await this._callGemini(apiKey, model, {
        contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: imageBase64 } }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 2000 }
      });
    } else if (provider === 'openai') {
      rawText = await this._callOpenAIVision(apiKey, model, imageBase64, prompt);
    } else {
      throw new Error('Анализ чертежей поддерживается только для Gemini и OpenAI');
    }

    return this._parseJsonResponse(rawText, `Чертёж: ${fileName}`);
  }

  // ============================================================
  // Анализ соответствия Заданию на дипломное проектирование
  // ============================================================
  async analyzeAssignmentCompliance(apiKey, imageBase64, docData, provider = 'gemini', model = 'gemini-2.0-flash') {
    if (!apiKey) throw new Error('API ключ не указан');
    
    // Берем начало документа (Введение, Содержание, Общие разделы), так как там кроется суть
    const docTextChunk = docData.fullText.slice(0, 15000); 

    const prompt = `Ты — строгий председатель комиссии и нормоконтролёр БГУИР.
Тебе предоставлены:
1. Фотография / скан "Задания на дипломное проектирование" (изображение).
2. Текст начала пояснительной записки студента (Содержание, Введение, первые разделы).

ТВОЯ ЗАДАЧА: Сверить Задание с фактическим текстом пояснительной записки.
Обязательно обрати внимание на:
- Соответствует ли тема дипломного проекта?
- Все ли пункты из "Перечня подлежащих разработке вопросов" отражены в содержании и тексте записки?
- Упоминается ли заявленный графический материал?
- Соответствуют ли исходные данные?

Выведи подробный анализ СТРОГО в формате JSON-массива:
[
  {
    "severity": "critical|warning|info|pass",
    "description": "конкретное описание совпадения или расхождения",
    "recommendation": "что нужно исправить (если есть ошибка)"
  }
]
Если всё совпадает идеально, выведи минимум один объект с severity="pass".`;

    let rawText;
    if (provider === 'gemini') {
      rawText = await this._callGemini(apiKey, model, {
        contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: imageBase64 } }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 2500 }
      });
    } else if (provider === 'openai') {
      rawText = await this._callOpenAIVision(apiKey, model, imageBase64, prompt);
    } else {
      throw new Error('Сверка с заданием поддерживается только для Gemini и OpenAI (Vision)');
    }

    return this._parseJsonResponse(rawText, 'Сверка с Заданием');
  }

  // ============================================================
  // ПРИВАТНЫЕ МЕТОДЫ
  // ============================================================

  async _callProvider(apiKey, provider, model, prompt) {
    if (provider === 'gemini') {
      return this._callGemini(apiKey, model, {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 3000 }
      });
    }
    const baseUrl = provider === 'openai'
      ? 'https://api.openai.com/v1/chat/completions'
      : 'https://api.deepseek.com/v1/chat/completions';
    return this._callOpenAICompat(apiKey, model, prompt, baseUrl);
  }

  async _callGemini(apiKey, model, body) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`;
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (!response.ok) {
      const errorText = await response.text();
      let msg = `Gemini API ошибка (${response.status})`;
      try { msg = JSON.parse(errorText).error?.message || msg; } catch (e) { /* ignore */ }
      throw new Error(msg);
    }
    const data = await response.json();
    if (!data.candidates?.length) throw new Error('Пустой ответ от Gemini API');
    const text = data.candidates[0]?.content?.parts?.[0]?.text;
    if (!text) throw new Error('Gemini вернул пустой текст');
    return text;
  }

  async _callOpenAICompat(apiKey, model, prompt, baseUrl) {
    const response = await fetch(baseUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model,
        messages: [
          { role: 'system', content: 'Ты — нормоконтролёр дипломных проектов БГУИР. Отвечаешь СТРОГО в формате JSON-массива.' },
          { role: 'user', content: prompt }
        ],
        max_tokens: 3000,
        temperature: 0.1
      })
    });
    if (!response.ok) {
      const errorText = await response.text();
      let msg = `API ошибка (${response.status})`;
      try { msg = JSON.parse(errorText).error?.message || msg; } catch (e) { /* ignore */ }
      throw new Error(msg);
    }
    const data = await response.json();
    const text = data.choices?.[0]?.message?.content;
    if (!text) throw new Error('Пустой ответ от API');
    return text;
  }

  async _callOpenAIVision(apiKey, model, imageBase64, prompt) {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
      body: JSON.stringify({
        model,
        messages: [{
          role: 'user',
          content: [
            { type: 'text', text: prompt },
            { type: 'image_url', image_url: { url: `data:image/jpeg;base64,${imageBase64}`, detail: 'high' } }
          ]
        }],
        max_tokens: 2000
      })
    });
    if (!response.ok) {
      const errorText = await response.text();
      let msg = `OpenAI Vision ошибка (${response.status})`;
      try { msg = JSON.parse(errorText).error?.message || msg; } catch (e) { /* ignore */ }
      throw new Error(msg);
    }
    const data = await response.json();
    return data.choices?.[0]?.message?.content || '';
  }

  _prepareText(text) {
    if (text.length <= this.maxTextLength) return text;
    const third = Math.floor(this.maxTextLength / 3);
    const start = text.slice(0, third);
    const mid   = text.slice(Math.floor(text.length / 2) - third / 2, Math.floor(text.length / 2) + third / 2);
    const end   = text.slice(-third);
    return start + '\n\n[... СЕРЕДИНА ДОКУМЕНТА ...]\n\n' + mid + '\n\n[... КОНЕЦ ДОКУМЕНТА ...]\n\n' + end;
  }

  _buildPrompt(docText, docData) {
    const headings = docData.paragraphs
      .filter(p => p.isHeading)
      .map(p => `  ${'  '.repeat(p.headingLevel - 1)}[H${p.headingLevel}] ${p.textTrimmed}`)
      .join('\n');

    return `Ты — нормоконтролёр-эксперт по дипломным проектам БГУИР (СТП 01-2024).
Твоя задача — проанализировать СОДЕРЖАНИЕ и СТИЛЬ изложения (НЕ форматирование — оно уже проверено автоматически).

СТРУКТУРА ЗАГОЛОВКОВ ДОКУМЕНТА:
${headings || '(заголовки не найдены)'}

ТЕКСТ ДОКУМЕНТА (фрагмент):
${docText}

ПРОВЕРЬ СЛЕДУЮЩЕЕ:
1. ВВЕДЕНИЕ: наличие цели, задач, объекта, предмета исследования, актуальности
2. ЗАКЛЮЧЕНИЕ: конкретные результаты (числа, «разработан», «получено»), не общие слова
3. СТИЛЬ: запрещены «я», «мы», «наш» — нужно «автором», «в работе», «получено»
4. АББРЕВИАТУРЫ: расшифровки при первом появлении
5. ССЫЛКИ: утверждения без источников («исследования показали»)
6. СТРУКТУРА: все обязательные разделы (введение, разделы, заключение, список литературы)
7. СОГЛАСОВАННОСТЬ: заголовки соответствуют содержимому

НЕ ПРОВЕРЯЙ шрифт, поля, отступы, рисунки, таблицы — это уже проверено.

ВЫВЕДИ СТРОГО в формате JSON-массива (максимум 12 объектов):
[
  {
    "severity": "critical|warning|info|pass",
    "section": "п. X.X.X СТП 01-2024 или 'Стиль изложения'",
    "description": "конкретное описание — что найдено в тексте",
    "location": "раздел или цитата до 80 символов",
    "recommendation": "конкретная правка или null для pass"
  }
]

Если нет реальных проблем — верни 1-2 объекта с severity=pass.`;
  }

  _buildVisualPrompt(pageNum) {
    return `Нормоконтролёр БГУИР. Проверь страницу ${pageNum} пояснительной записки (СТП 01-2024).
Поля 30/15/20/20 мм, шрифт TNR 14пт, отступ 1.25 см, заголовки ПРОПИСНЫЕ без точки, рисунки «Рисунок N – Название».
Выведи ТОЛЬКО JSON: [{"severity":"critical|warning|info|pass","description":"...","recommendation":"..."}]`;
  }

  _parseJsonResponse(text, source) {
    const findings = [];
    let jsonStr = text.trim()
      .replace(/^```(?:json)?\s*/i, '')
      .replace(/\s*```\s*$/, '');

    const match = jsonStr.match(/\[[\s\S]*\]/);
    if (match) jsonStr = match[0];

    try {
      const items = JSON.parse(jsonStr);
      if (!Array.isArray(items)) throw new Error('Not an array');
      for (const item of items) {
        if (!item.severity || !item.description) continue;
        findings.push({
          severity:       item.severity,
          id:             `ai-${Math.random().toString(36).slice(2, 8)}`,
          section:        item.section || 'ИИ-анализ',
          description:    item.description,
          location:       item.location || source,
          recommendation: item.recommendation || null,
          source:         'ai'
        });
      }
    } catch (e) {
      console.warn('Could not parse AI JSON:', e.message, '\nRaw:', text.slice(0, 200));
      findings.push({
        severity: 'info', id: 'ai-raw', section: 'ИИ-анализ',
        description: 'ИИ предоставил анализ в свободном формате (см. рекомендацию)',
        location: source,
        recommendation: text.slice(0, 500),
        source: 'ai'
      });
    }

    return findings;
  }

  // Утилиты изображений
  static async fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = () => resolve(reader.result.split(',')[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  static async pdfFirstPageToBase64(file) {
    if (!window.pdfjsLib) return AiChecker.fileToBase64(file);
    try {
      const ab   = await file.arrayBuffer();
      const pdf  = await pdfjsLib.getDocument({ data: ab }).promise;
      const page = await pdf.getPage(1);
      const vp   = page.getViewport({ scale: 1.5 });
      const c    = document.createElement('canvas');
      c.width = vp.width; c.height = vp.height;
      await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      return c.toDataURL('image/jpeg', 0.85).split(',')[1];
    } catch (e) {
      return AiChecker.fileToBase64(file);
    }
  }

  static async captureDocumentPage(htmlContent) {
    if (!window.html2canvas) return null;
    const c = document.getElementById('render-area');
    if (!c) return null;
    c.style.cssText = 'position:absolute;left:-9999px;top:0;width:794px;min-height:1123px;background:white;padding:76px 57px 76px 113px;font-family:Times New Roman,serif;font-size:14pt;line-height:1.5;color:black;';
    c.innerHTML = htmlContent;
    try {
      const canvas = await html2canvas(c, { width: 794, height: 1123, scale: 1.2, useCORS: true, logging: false });
      return canvas.toDataURL('image/jpeg', 0.8).split(',')[1];
    } catch (e) { return null; }
    finally { c.innerHTML = ''; }
  }
}
