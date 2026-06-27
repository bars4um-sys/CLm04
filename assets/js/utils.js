/* ============================================================
 * Сверка — общие утилиты (utils.js)
 * Переиспользуемое ядро для всех функций-модулей.
 * Доступно глобально как window.AccProgUtils
 * ============================================================ */
(function () {
  'use strict';

  /* ---------- Чтение файлов ---------- */

  // Читает Excel/CSV как матрицу строк (header:1). Возвращает Promise.
  function readSheetMatrix(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          if (typeof XLSX === 'undefined') {
            reject(new Error('Библиотека XLSX не загружена'));
            return;
          }
          var data = new Uint8Array(e.target.result);
          var wb = XLSX.read(data, { type: 'array' });
          var name = wb.SheetNames[0];
          var sheet = wb.Sheets[name];
          var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
          resolve({ rows: rows, sheetName: name });
        } catch (err) { reject(err); }
      };
      reader.onerror = function () { reject(reader.error || new Error('Ошибка чтения файла')); };
      reader.readAsArrayBuffer(file);
    });
  }

  // Читает Excel/CSV как массив объектов (с заголовками). Возвращает Promise.
  function readSheetObjects(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      var lower = (file.name || '').toLowerCase();
      reader.onload = function (e) {
        try {
          var json;
          if (lower.endsWith('.csv')) {
            json = parseCSV(String(e.target.result));
          } else {
            if (typeof XLSX === 'undefined') {
              reject(new Error('Библиотека XLSX не загружена'));
              return;
            }
            var data = new Uint8Array(e.target.result);
            var wb = XLSX.read(data, { type: 'array' });
            var sheet = wb.Sheets[wb.SheetNames[0]];
            json = XLSX.utils.sheet_to_json(sheet);
          }
          resolve(json);
        } catch (err) { reject(err); }
      };
      reader.onerror = function () { reject(reader.error || new Error('Ошибка чтения файла')); };
      if (lower.endsWith('.csv')) reader.readAsText(file, 'utf-8');
      else reader.readAsArrayBuffer(file);
    });
  }

  // Простой парсер CSV -> массив объектов
  function parseCSV(text) {
    var lines = text.trim().split('\n');
    if (lines.length < 2) throw new Error('Нужны заголовок и минимум одна строка данных');
    var headers = lines[0].split(',').map(function (h) { return h.trim(); });
    var out = [];
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      var vals = line.split(',').map(function (v) { return v.trim(); });
      var row = {};
      headers.forEach(function (h, idx) {
        var v = vals[idx] || '';
        if (v !== '' && !isNaN(v)) v = parseFloat(v);
        row[h] = v;
      });
      out.push(row);
    }
    return out;
  }

  // Преобразует массив строк (header:1) в массив объектов как sheet_to_json
  // Первая строка — заголовки, остальные — данные
  function sheetRowsToObjects(rows) {
    if (!rows || rows.length < 2) return [];
    var headers = rows[0].map(function (h) { return String(h == null ? '' : h).trim(); });
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var obj = {};
      var empty = true;
      for (var j = 0; j < headers.length; j++) {
        var val = (j < row.length) ? row[j] : null;
        if (val !== null && val !== '' && val !== undefined) empty = false;
        obj[headers[j]] = val;
      }
      if (!empty) result.push(obj);
    }
    return result;
  }

  /* ---------- Нормализация ---------- */

  function normalizeSum(value) {
    if (value === null || value === undefined || value === '') return null;
    var str = String(value).trim().replace(/\s+/g, '').replace(',', '.');
    str = str.replace(/[^\d.\-]/g, '');
    if (!str || str === '-' || str === '.') return null;
    var num = parseFloat(str);
    if (isNaN(num)) return null;
    return Math.round(num * 100) / 100;
  }

  function normalizeDate(value) {
    if (value === null || value === undefined || value === '') return null;
    if (value instanceof Date && !isNaN(value.getTime())) return fmtDate(value);
    var str = String(value).trim();
    if (!str) return null;

    // Excel serial date
    if (/^\d+(\.\d+)?$/.test(str)) {
      var serial = parseFloat(str);
      if (serial > 0 && serial < 100000) {
        var d = new Date(Math.round((serial - 25569) * 86400 * 1000));
        if (!isNaN(d.getTime())) return fmtDate(d);
      }
    }
    // DD.MM.YYYY
    var m1 = str.match(/^(\d{1,2})[./\-\s](\d{1,2})[./\-\s](\d{2,4})$/);
    if (m1) {
      var dd = +m1[1], mm = +m1[2], yy = +m1[3];
      if (yy < 100) yy += (yy >= 70 ? 1900 : 2000);
      if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31)
        return yy + '-' + pad(mm) + '-' + pad(dd);
    }
    // ISO
    var m2 = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m2) {
      var y = +m2[1], mo = +m2[2], da = +m2[3];
      if (mo >= 1 && mo <= 12 && da >= 1 && da <= 31)
        return y + '-' + pad(mo) + '-' + pad(da);
    }
    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) return fmtDate(parsed);
    return null;
  }

  function pad(n) { return String(n).padStart(2, '0'); }
  function fmtDate(d) {
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }

  // Если первая строка пустая — шапка во второй (двухуровневая шапка)
  function detectHeaderRow(rows) {
    if (!rows || rows.length === 0) return 0;
    var first = rows[0] || [];
    var empty = true;
    for (var i = 0; i < first.length; i++) {
      if (!(first[i] === null || first[i] === undefined || String(first[i]).trim() === '')) {
        empty = false; break;
      }
    }
    return (empty && rows.length > 1) ? 1 : 0;
  }

  // Отбрасывает итоговые/заголовочные/пустые строки (для объектного формата)
  function cleanData(jsonData) {
    var skip = ['итого', 'всего', 'total', 'sum', 'сумма', 'grand total', 'конец',
      'end', 'footer', 'подитог', 'subtotal', 'наименование', 'название', 'компания'];
    return jsonData.filter(function (row) {
      if (!row || Object.keys(row).length === 0) return false;
      var firstText = '';
      for (var k in row) {
        if (typeof row[k] === 'string') { firstText = row[k].trim().toLowerCase(); break; }
      }
      for (var i = 0; i < skip.length; i++) if (firstText.includes(skip[i])) return false;
      var hasNum = false;
      for (var k2 in row) {
        if (typeof row[k2] === 'number' && !isNaN(row[k2])) { hasNum = true; break; }
      }
      return hasNum && firstText !== '';
    });
  }

  /* ---------- Форматирование / HTML ---------- */

  function formatNumber(num) {
    if (num === null || num === undefined) return '0';
    return Number(num).toLocaleString('ru-RU', {
      minimumFractionDigits: 2, maximumFractionDigits: 2
    });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* ---------- Рендер таблицы ---------- */

  // headers: массив строк; rows: массив массивов значений
  function renderTable(headers, rows, opts) {
    opts = opts || {};
    var html = '<div class="table-wrapper"><table><thead><tr>';
    headers.forEach(function (h, idx) {
      var marker = (opts.markCols && opts.markCols[idx])
        ? ' <span class="' + (opts.markCols[idx].cls || 'col-marker') + '">' +
          opts.markCols[idx].mark + '</span>' : '';
      html += '<th>' + escapeHtml(String(h == null ? '' : h)) + marker + '</th>';
    });
    html += '</tr></thead><tbody>';
    rows.forEach(function (row) {
      html += '<tr>';
      for (var c = 0; c < headers.length; c++) {
        var v = row[c];
        html += '<td>' + escapeHtml(v === null || v === undefined ? '' : String(v)) + '</td>';
      }
      html += '</tr>';
    });
    html += '</tbody></table></div>';
    return html;
  }

  /* ---------- Экспорт ---------- */

  // data: массив массивов (первая строка — заголовки)
  function exportXlsx(data, sheetName, fileBase) {
    var ws = XLSX.utils.aoa_to_sheet(data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Результат');
    var ts = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, (fileBase || 'result') + '_' + ts + '.xlsx');
  }

  /* ---------- Чтение из URL (для тестовых таблиц) ---------- */

  // Загружает xlsx/csv по URL и возвращает тот же формат, что readSheetMatrix: { rows, sheetName }
  function loadSheetFromUrl(url) {
    return fetch(url)
      .then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status + ' при загрузке ' + url);
        return r.arrayBuffer();
      })
      .then(function (buf) {
        if (typeof XLSX === 'undefined') throw new Error('Библиотека XLSX не загружена');
        var data = new Uint8Array(buf);
        var wb = XLSX.read(data, { type: 'array' });
        var name = wb.SheetNames[0];
        var sheet = wb.Sheets[name];
        var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
        return { rows: rows, sheetName: name, fileName: url.split('/').pop() };
      });
  }

  /* ---------- Предпросмотр таблиц (модальное окно) ---------- */

  // Показывает модальное окно с данными таблицы.
  // data — массив объектов (как из readSheetObjects) или массив массивов (как из readSheetMatrix)
  function showPreview(data, title) {
    var overlay = document.getElementById('modalOverlay');
    var body = document.getElementById('modalBody');
    var titleEl = document.getElementById('modalTitle');
    if (!overlay || !body) return;

    titleEl.textContent = title || 'Предпросмотр данных';

    // Определяем формат и строим заголовки
    var headers = [];
    var rows = [];
    if (data && data.length > 0) {
      if (Array.isArray(data[0])) {
        // Матрица строк (как readSheetMatrix)
        rows = data;
        if (rows.length > 0) {
          for (var j = 0; j < rows[0].length; j++) {
            headers.push('#' + (j + 1));
          }
        }
      } else {
        // Массив объектов (как readSheetObjects)
        headers = Object.keys(data[0]);
        rows = data.map(function (obj) {
          return headers.map(function (k) { return obj[k]; });
        });
      }
    }

    var html = '<div class="summary">Всего строк: <b>' + rows.length + '</b></div>';
    html += '<div class="table-wrapper"><table><thead><tr>';
    headers.forEach(function (h) {
      html += '<th>' + escapeHtml(String(h == null ? '' : h)) + '</th>';
    });
    html += '</tr></thead><tbody>';

    // Показываем максимум 100 строк
    var maxShow = Math.min(rows.length, 100);
    for (var i = 0; i < maxShow; i++) {
      html += '<tr>';
      var row = rows[i] || [];
      for (var j = 0; j < headers.length; j++) {
        var v = j < row.length ? row[j] : '';
        html += '<td>' + escapeHtml(v == null ? '' : String(v)) + '</td>';
      }
      html += '</tr>';
    }
    if (rows.length > 100) {
      html += '<tr><td colspan="' + headers.length + '" style="text-align:center;color:var(--text-muted)">… и ещё ' + (rows.length - 100) + ' строк …</td></tr>';
    }
    html += '</tbody></table></div>';

    body.innerHTML = html;
    overlay.style.display = 'flex';
  }

  // Закрытие модального окна
  document.addEventListener('DOMContentLoaded', function () {
    var overlay = document.getElementById('modalOverlay');
    var closeBtn = document.getElementById('modalClose');
    if (overlay && closeBtn) {
      closeBtn.addEventListener('click', function () { overlay.style.display = 'none'; });
      overlay.addEventListener('click', function (e) {
        if (e.target === overlay) overlay.style.display = 'none';
      });
    }
  });

  /* ---------- Экспорт API ---------- */
  window.AccProgUtils = {
    readSheetMatrix: readSheetMatrix,
    readSheetObjects: readSheetObjects,
    sheetRowsToObjects: sheetRowsToObjects,
    loadSheetFromUrl: loadSheetFromUrl,
    parseCSV: parseCSV,
    normalizeSum: normalizeSum,
    normalizeDate: normalizeDate,
    detectHeaderRow: detectHeaderRow,
    cleanData: cleanData,
    formatNumber: formatNumber,
    escapeHtml: escapeHtml,
    renderTable: renderTable,
    exportXlsx: exportXlsx,
    showPreview: showPreview
  };
})();
