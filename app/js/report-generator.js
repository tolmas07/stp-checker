/**
 * Report Generator — генерирует HTML-отчёт из результатов проверки
 * и управляет отображением/экспортом
 */

class ReportGenerator {

  // ============================================================
  // Генерация HTML карточки одной находки
  // ============================================================
  static renderFinding(finding, index) {
    const icons = {
      critical: '❌',
      warning: '⚠️',
      info: 'ℹ️',
      pass: '✅'
    };

    const labels = {
      critical: 'Критическая ошибка',
      warning: 'Предупреждение',
      info: 'Замечание',
      pass: 'Соответствует'
    };

    const icon = icons[finding.severity] || 'ℹ️';
    const aiTag = finding.source === 'ai'
      ? '<span class="finding-ai-badge">🤖 ИИ</span>'
      : '';

    // Извлекаем номер страницы
    const pageMatch = finding.location && finding.location.match(/~стр\.\s*(\d+)/);
    const pageNum = finding.pageNum || (pageMatch ? parseInt(pageMatch[1]) : null);

    // Бейдж: страница ИЛИ «Весь документ» для глобальных проверок
    let pageBadge = '';
    if (pageNum) {
      pageBadge = `<span class="finding-page-badge">стр. ~${pageNum}</span>`;
    } else if (finding.severity !== 'pass' && finding.source !== 'ai') {
      pageBadge = `<span class="finding-page-badge finding-page-global">📄 весь документ</span>`;
    }

    // Кнопка перехода к странице в предпросмотре (только если есть номер)
    const gotoBtn = (pageNum && finding.severity !== 'pass')
      ? `<button class="finding-goto-btn" onclick="event.stopPropagation(); app.navigateToPage(${pageNum})" title="Показать эту страницу в предпросмотре">📌 Перейти</button>`
      : '';

    const bodyHtml = finding.severity !== 'pass' ? `
      <div class="finding-body">
        <div class="finding-desc">${this._escHtml(finding.description)}</div>
        ${finding.location ? `<div class="finding-location-full">📍 ${this._escHtml(finding.location)}</div>` : ''}
        ${finding.recommendation ? `
          <div class="finding-rec">
            <strong>Рекомендация:</strong> ${this._escHtml(finding.recommendation)}
          </div>
        ` : ''}
      </div>
    ` : '';

    return `
      <div class="finding finding-${finding.severity}" data-severity="${finding.severity}" id="finding-${index}">
        <div class="finding-header" onclick="this.closest('.finding').classList.toggle('open')">
          <span class="finding-icon">${icon}</span>
          <div class="finding-main">
            <div class="finding-title">${this._escHtml(finding.description).slice(0, 100)}${aiTag}</div>
            <div class="finding-meta">
              <span class="finding-section">${this._escHtml(finding.section || '')}</span>
              ${pageBadge}
              ${gotoBtn}
            </div>
          </div>
          ${finding.severity !== 'pass' ? '<span class="finding-toggle">▾</span>' : ''}
        </div>
        ${bodyHtml}
      </div>
    `;
  }

  // ============================================================
  // Расчёт итогового балла готовности (0–10)
  // ============================================================
  static calculateScore(findings) {
    const counts = {
      critical: findings.filter(f => f.severity === 'critical').length,
      warning:  findings.filter(f => f.severity === 'warning').length,
      info:     findings.filter(f => f.severity === 'info').length,
      pass:     findings.filter(f => f.severity === 'pass').length,
    };

    // Базовый балл — 10
    let score = 10;

    // Критические ошибки сильно снижают балл
    score -= Math.min(counts.critical * 1.5, 6);

    // Предупреждения умеренно снижают
    score -= Math.min(counts.warning * 0.4, 3);

    // Замечания немного снижают
    score -= Math.min(counts.info * 0.15, 1);

    score = Math.max(0, Math.min(10, Math.round(score * 10) / 10));

    return { score, counts };
  }

  // ============================================================
  // Описание оценки
  // ============================================================
  static getScoreDescription(score, counts) {
    if (counts.critical > 5) return '🔴 Работа не готова к нормоконтролю';
    if (counts.critical > 2) return '🔴 Критические ошибки требуют исправления';
    if (counts.critical > 0) return '🟠 Есть критические нарушения СТП';
    if (score >= 8.5) return '🟢 Отличное оформление, готова к защите';
    if (score >= 7)   return '🟡 Хорошее оформление, исправьте предупреждения';
    if (score >= 5)   return '🟠 Удовлетворительно, требуются исправления';
    return '🔴 Документ требует значительной доработки';
  }

  // ============================================================
  // Генерация полного текстового отчёта (для копирования)
  // ============================================================
  static generateTextReport(findings, docInfo) {
    const { score, counts } = this.calculateScore(findings);
    const lines = [];

    lines.push('═══════════════════════════════════════════════');
    lines.push('  ОТЧЁТ НОРМОКОНТРОЛЯ — СТП 01-2024 БГУИР');
    lines.push('═══════════════════════════════════════════════');
    lines.push(`Файл: ${docInfo?.fileName || 'Неизвестно'}`);
    lines.push(`Дата: ${new Date().toLocaleString('ru-RU')}`);
    lines.push('');
    lines.push(`ИТОГ: ${score}/10 — ${this.getScoreDescription(score, counts)}`);
    lines.push(`  ❌ Критических: ${counts.critical}`);
    lines.push(`  ⚠️  Предупреждений: ${counts.warning}`);
    lines.push(`  ℹ️  Замечаний: ${counts.info}`);
    lines.push(`  ✅ Пройдено: ${counts.pass}`);
    lines.push('');

    const sections = [
      { key: 'critical', title: 'КРИТИЧЕСКИЕ ОШИБКИ (необходимо исправить)', icon: '❌' },
      { key: 'warning',  title: 'ПРЕДУПРЕЖДЕНИЯ', icon: '⚠️' },
      { key: 'info',     title: 'ЗАМЕЧАНИЯ', icon: 'ℹ️' },
      { key: 'pass',     title: 'ПРОЙДЕНО ✓', icon: '✅' },
    ];

    for (const section of sections) {
      const items = findings.filter(f => f.severity === section.key);
      if (items.length === 0) continue;

      lines.push('───────────────────────────────────────────────');
      lines.push(`${section.icon} ${section.title} (${items.length})`);
      lines.push('');

      for (const item of items) {
        lines.push(`[${item.section}] ${item.description}`);
        if (item.location) lines.push(`  📍 ${item.location}`);
        if (item.recommendation) lines.push(`  💡 ${item.recommendation}`);
        lines.push('');
      }
    }

    lines.push('═══════════════════════════════════════════════');
    lines.push('Проверено приложением БГУИР Нормоконтроль v1.0');

    return lines.join('\n');
  }

  // ============================================================
  // Обновить счётчики в шапке результатов
  // ============================================================
  static updateSummary(findings) {
    const { score, counts } = this.calculateScore(findings);

    document.getElementById('critical-count').textContent = counts.critical;
    document.getElementById('warning-count').textContent  = counts.warning;
    document.getElementById('info-count').textContent     = counts.info;
    document.getElementById('pass-count').textContent     = counts.pass;

    const scoreVal  = document.getElementById('score-val');
    const scoreBar  = document.getElementById('score-bar');
    const scoreDesc = document.getElementById('score-desc');
    const scoreCircle = document.getElementById('score-circle');

    scoreVal.textContent = score.toFixed(1);
    scoreDesc.textContent = this.getScoreDescription(score, counts);
    scoreBar.style.width = `${score * 10}%`;

    // Цвет оценки
    scoreCircle.className = 'score-circle';
    scoreBar.style.background = '';
    if (score >= 7.5) {
      scoreCircle.classList.add('score-high');
      scoreBar.style.background = 'var(--pass-icon)';
    } else if (score >= 5) {
      scoreCircle.classList.add('score-mid');
      scoreBar.style.background = 'var(--warning-icon)';
    } else {
      scoreCircle.classList.add('score-low');
      scoreBar.style.background = 'var(--critical-icon)';
    }
  }

  // ============================================================
  // Отрисовать все находки в контейнере
  // ============================================================
  static renderFindings(findings, filter = 'all') {
    const container = document.getElementById('findings-container');
    if (!container) return;

    const filtered = filter === 'all'
      ? findings
      : findings.filter(f => f.severity === filter);

    if (filtered.length === 0) {
      container.innerHTML = `
        <div class="no-findings">
          <div class="big-icon">${filter === 'pass' ? '🎉' : '🔍'}</div>
          <div>${filter === 'pass' ? 'Нет пройденных проверок' : 'Нарушений этого типа не найдено'}</div>
        </div>
      `;
      return;
    }

    // Сортировка: сначала критические, потом warnings, info, pass
    const order = { critical: 0, warning: 1, info: 2, pass: 3 };
    const sorted = [...filtered].sort((a, b) => (order[a.severity] || 0) - (order[b.severity] || 0));

    container.innerHTML = sorted.map((f, i) => this.renderFinding(f, i)).join('');
  }

  // ============================================================
  // Экспорт в PDF через Print
  // ============================================================
  static exportToPDF() {
    window.print();
  }

  // ============================================================
  // Копировать в буфер обмена
  // ============================================================
  static async copyToClipboard(findings, docInfo) {
    const text = this.generateTextReport(findings, docInfo);
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      // Fallback
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      return true;
    }
  }

  // ============================================================
  // Экранирование HTML
  // ============================================================
  static _escHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
}
