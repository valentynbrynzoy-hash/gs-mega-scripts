/**
 * ═══════════════════════════════════════════════════════════════════════════
 * 🛠️ GS MEGA SCRIPTS — ДУБЛИ + РАЗДЕЛЕНИЕ КОНТАКТОВ + ОБРАБОТКА ССЫЛОК
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * Google Apps Script для обработки данных об инфлюенсерах в Google Sheets.
 *
 * ФУНКЦИИ В МЕНЮ:
 *   🧹 Удалить дубли: Связь
 *   🧹 Удалить дубли: Январь 2026
 *   📂 Разделить контакты: Связь
 *   📂 Разделить контакты: Январь 2026
 *   🔗 Обработать ссылки: Связь        ← колонка B → B
 *   🔗 Обработать ссылки: Январь 2026  ← колонка D → D
 *   ⚙️ Показать настройки
 *
 * ТРЕБОВАНИЯ:
 *   Google Sheets + Google Apps Script (бесплатно)
 *
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════════════════
// ⚙️ НАСТРОЙКИ — МЕНЯЙ ЗДЕСЬ
// ══════════════════════════════════════════════════════════════════════════

var НАСТРОЙКИ = {

  // ── Дубли: лист "Связь" ───────────────────────────────────────────────
  СВЯЗЬ_ДУБЛИ: {
    НАЗВАНИЕ_ЛИСТА  : "Связь",
    КОЛОНКА_ССЫЛОК  : "B",
    НАЧАЛЬНАЯ_СТРОКА: 1,
  },

  // ── Дубли: лист "Январь 2026" ─────────────────────────────────────────
  ЯНВАРЬ_ДУБЛИ: {
    НАЗВАНИЕ_ЛИСТА  : "Январь 2026",
    КОЛОНКА_ССЫЛОК  : "D",
    НАЧАЛЬНАЯ_СТРОКА: 1,
  },

  // ── Разделение контактов: лист "Связь" ────────────────────────────────
  СВЯЗЬ_КОНТАКТЫ: {
    НАЗВАНИЕ_ЛИСТА     : "Связь",
    КОЛОНКА_НОМЕРОВ    : "G",
    КОЛОНКА_ПОЧТ       : "G",
    НАЧАЛЬНАЯ_СТРОКА   : 2,
    ДИАПАЗОН_ОТ        : "B",
    ДИАПАЗОН_ДО        : "H",
    ЛИСТ_НОМЕРА        : "📞 Номера (Связь)",
    ЛИСТ_ПОЧТЫ         : "📧 Почты (Связь)",
    ЛИСТ_ПУСТЫЕ        : "⬜ Без контактов (Связь)",
  },

  // ── Разделение контактов: лист "Январь 2026" ──────────────────────────
  ЯНВАРЬ_КОНТАКТЫ: {
    НАЗВАНИЕ_ЛИСТА     : "Январь 2026",
    КОЛОНКА_НОМЕРОВ    : "O",
    КОЛОНКА_ПОЧТ       : "O",
    НАЧАЛЬНАЯ_СТРОКА   : 2,
    ДИАПАЗОН_ОТ        : "B",
    ДИАПАЗОН_ДО        : "R",
    ЛИСТ_НОМЕРА        : "📞 Номера (Январь)",
    ЛИСТ_ПОЧТЫ         : "📧 Почты (Январь)",
    ЛИСТ_ПУСТЫЕ        : "⬜ Без контактов (Январь)",
  },

  // ── Обработка ссылок: лист "Связь" ────────────────────────────────────
  СВЯЗЬ_ССЫЛКИ: {
    НАЗВАНИЕ_ЛИСТА  : "Связь",
    КОЛОНКА        : "B",
    НАЧАЛЬНАЯ_СТРОКА: 2,
  },

  // ── Обработка ссылок: лист "Январь 2026" ──────────────────────────────
  ЯНВАРЬ_ССЫЛКИ: {
    НАЗВАНИЕ_ЛИСТА  : "Январь 2026",
    КОЛОНКА        : "D",
    НАЧАЛЬНАЯ_СТРОКА: 2,
  },

};

// ══════════════════════════════════════════════════════════════════════════
// 📋 МЕНЮ
// ══════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Мои скрипты')
    .addItem('🧹 Удалить дубли: Связь',             'удалитьДублиСвязь')
    .addItem('🧹 Удалить дубли: Январь 2026',        'удалитьДублиЯнварь')
    .addSeparator()
    .addItem('📂 Разделить контакты: Связь',         'разделитьКонтактыСвязь')
    .addItem('📂 Разделить контакты: Январь 2026',   'разделитьКонтактыЯнварь')
    .addSeparator()
    .addItem('🔗 Обработать ссылки: Связь',          'обработатьСсылкиСвязь')
    .addItem('🔗 Обработать ссылки: Январь 2026',    'обработатьСсылкиЯнварь')
    .addSeparator()
    .addItem('⚙️ Показать настройки',               'показатьНастройки')
    .addToUi();
}

// ══════════════════════════════════════════════════════════════════════════
// 🧹 ДУБЛИ
// ══════════════════════════════════════════════════════════════════════════

function удалитьДублиСвязь() {
  var c = НАСТРОЙКИ.СВЯЗЬ_ДУБЛИ;
  _дубли(c.НАЗВАНИЕ_ЛИСТА, c.КОЛОНКА_ССЫЛОК, c.НАЧАЛЬНАЯ_СТРОКА);
}

function удалитьДублиЯнварь() {
  var c = НАСТРОЙКИ.ЯНВАРЬ_ДУБЛИ;
  _дубли(c.НАЗВАНИЕ_ЛИСТА, c.КОЛОНКА_ССЫЛОК, c.НАЧАЛЬНАЯ_СТРОКА);
}

function _дубли(имяЛиста, колонка, начСтрока) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(имяЛиста);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      '❌ Лист "' + имяЛиста + '" не найден!\n\n'
      + 'Проверь НАСТРОЙКИ — имя вкладки должно совпадать точно.'
    );
    return;
  }

  var lastRow = sheet.getLastRow();
  if (начСтрока > lastRow) {
    SpreadsheetApp.getUi().alert('⚠️ Нет данных — лист пустой или начальная строка больше последней.');
    return;
  }

  var colIdx   = _col(колонка) - 1;
  var numRows  = lastRow - начСтрока + 1;
  var range    = sheet.getRange(начСтрока, 1, numRows, colIdx + 1);
  var values   = range.getValues();
  var formulas = range.getFormulas();

  var seen = {}, toDelete = [], примеры = [];

  for (var i = 0; i < values.length; i++) {
    var val = _cellValue(String(values[i][colIdx] || ''), formulas[i][colIdx] || '');
    if (!val) continue;
    var key = val.trim().toLowerCase();
    if (seen[key] !== undefined) {
      toDelete.push(начСтрока + i);
      if (примеры.length < 5) примеры.push('  • стр.' + (начСтрока + i) + ': ' + val);
    } else {
      seen[key] = начСтрока + i;
    }
  }

  if (toDelete.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Дубли не найдены!\n\nЛист: ' + имяЛиста + '\nПроверено строк: ' + numRows
    );
    return;
  }

  for (var j = toDelete.length - 1; j >= 0; j--) sheet.deleteRow(toDelete[j]);

  SpreadsheetApp.getUi().alert(
    '✅ ГОТОВО!\n\n'
    + '📋 Лист: '            + имяЛиста           + '\n'
    + '📊 Проверено строк: ' + numRows             + '\n'
    + '🗑️ Удалено дублей: ' + toDelete.length     + '\n'
    + '💾 Осталось: '        + (numRows - toDelete.length) + '\n'
    + (примеры.length
        ? '\nПримеры удалённых:\n' + примеры.join('\n')
          + (toDelete.length > 5 ? '\n  ... и ещё ' + (toDelete.length - 5) : '')
        : '')
  );
}

// ══════════════════════════════════════════════════════════════════════════
// 📂 РАЗДЕЛЕНИЕ КОНТАКТОВ
// ══════════════════════════════════════════════════════════════════════════

function разделитьКонтактыСвязь()  { _разделить(НАСТРОЙКИ.СВЯЗЬ_КОНТАКТЫ);  }
function разделитьКонтактыЯнварь() { _разделить(НАСТРОЙКИ.ЯНВАРЬ_КОНТАКТЫ); }

function _разделить(cfg) {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var src = ss.getSheetByName(cfg.НАЗВАНИЕ_ЛИСТА);
  if (!src) { ui.alert('❌ Лист "' + cfg.НАЗВАНИЕ_ЛИСТА + '" не найден!'); return; }

  var lastRow      = src.getLastRow();
  var defaultStart = cfg.НАЧАЛЬНАЯ_СТРОКА;
  if (defaultStart > lastRow) { ui.alert('⚠️ Нет данных для обработки.'); return; }

  var ответ = ui.prompt(
    '📂 Разделить контакты — выбор строк',
    '📋 Лист: "' + cfg.НАЗВАНИЕ_ЛИСТА + '"\nВсего строк: ' + lastRow + '\n\n'
    + 'Введи диапазон (или оставь пустым — весь лист):\n\n'
    + '  "5:200"  — с 5 по 200\n  "50:"    — с 50 до конца\n'
    + '  ":100"   — с начала по 100\n  пусто / "всё" — с ' + defaultStart + ' по ' + lastRow,
    ui.ButtonSet.OK_CANCEL
  );
  if (ответ.getSelectedButton() !== ui.Button.OK) return;

  var ввод = ответ.getResponseText().trim().toLowerCase();
  var startRow, endRow;

  if (ввод === '' || ввод === 'всё' || ввод === 'все') {
    startRow = defaultStart; endRow = lastRow;
  } else {
    var части = ввод.split(':');
    if (части.length !== 2) { ui.alert('❌ Неверный формат. Пример: "5:200"'); return; }
    startRow = части[0] === '' ? defaultStart : parseInt(части[0], 10);
    endRow   = части[1] === '' ? lastRow      : parseInt(части[1], 10);
    if (isNaN(startRow) || isNaN(endRow)) { ui.alert('❌ Введи числа.'); return; }
    if (startRow < 1) startRow = 1;
    if (endRow > lastRow) endRow = lastRow;
    if (startRow > endRow) { ui.alert('❌ Начало больше конца.'); return; }
  }

  var colFrom   = _col(cfg.ДИАПАЗОН_ОТ);
  var colTo     = _col(cfg.ДИАПАЗОН_ДО);
  var totalCols = colTo - colFrom + 1;
  var numRows   = endRow - startRow + 1;

  var idxN = _col(cfg.КОЛОНКА_НОМЕРОВ) - colFrom;
  var idxP = _col(cfg.КОЛОНКА_ПОЧТ)   - colFrom;
  var однаКолонка = (cfg.КОЛОНКА_НОМЕРОВ === cfg.КОЛОНКА_ПОЧТ);

  var data = src.getRange(startRow, colFrom, numRows, totalCols).getValues();
  var srcRowsN = [], srcRowsP = [], srcRowsEmpty = [];

  for (var i = 0; i < numRows; i++) {
    var row = data[i];
    var естьН = false, естьП = false;
    if (однаКолонка) {
      String(row[idxN] || '').split(/[\n,;]+/).forEach(function(part) {
        part = part.trim();
        if (_почта(part)) естьП = true;
        if (_номер(part)) естьН = true;
      });
    } else {
      String(row[idxN] || '').split(/[\n,;]+/).forEach(function(p) { if (_номер(p.trim())) естьН = true; });
      String(row[idxP] || '').split(/[\n,;]+/).forEach(function(p) { if (_почта(p.trim())) естьП = true; });
    }
    var srcRow = startRow + i;
    if      (естьН && естьП) { srcRowsN.push(srcRow); srcRowsP.push(srcRow); }
    else if (естьН)           { srcRowsN.push(srcRow); }
    else if (естьП)           { srcRowsP.push(srcRow); }
    else                      { srcRowsEmpty.push(srcRow); }
  }

  var hasHeader = (startRow > 1);
  var sheetN = _лист(ss, cfg.ЛИСТ_НОМЕРА);
  var sheetP = _лист(ss, cfg.ЛИСТ_ПОЧТЫ);
  var sheetE = _лист(ss, cfg.ЛИСТ_ПУСТЫЕ);
  sheetN.clearContents(); sheetP.clearContents(); sheetE.clearContents();

  function _записатьЛист(targetSheet, sourceRows, цветШапки) {
    if (sourceRows.length === 0) return;

    var dataOffset = 1;

    if (hasHeader) {
      src.getRange(1, colFrom, 1, totalCols)
         .copyTo(targetSheet.getRange(1, 2, 1, totalCols));
      targetSheet.getRange(1, 1).setValue('\u2116');
      targetSheet.getRange(1, 1, 1, totalCols + 1)
        .setBackground(цветШапки).setFontColor('#ffffff').setFontWeight('bold');
      dataOffset = 2;
    }

    var nums = [];
    for (var r = 0; r < sourceRows.length; r++) {
      src.getRange(sourceRows[r], colFrom, 1, totalCols)
         .copyTo(targetSheet.getRange(dataOffset + r, 2, 1, totalCols));
      nums.push([r + 1]);
    }
    targetSheet.getRange(dataOffset, 1, nums.length, 1).setValues(nums);

    var fullDataRange = targetSheet.getRange(dataOffset, 1, sourceRows.length, totalCols + 1);
    fullDataRange.setBackground('#ffffff').setFontWeight('normal');

    var колКонтакт = idxN + 2;
    for (var r = 0; r < sourceRows.length; r++) {
      var cell = targetSheet.getRange(dataOffset + r, колКонтакт);
      var raw  = String(cell.getValue() || '').trim();
      if (raw.indexOf('\n') !== -1 || raw.indexOf(';') !== -1) {
        var parts = raw.split(/[\n;]+/)
          .map(function(p) { return p.trim(); })
          .filter(function(p) { return p !== ''; });
        cell.setValue(parts.join(', '));
      }
    }

    var dataCols    = totalCols + 1;
    var dataRegion  = targetSheet.getRange(dataOffset, 1, sourceRows.length, dataCols);
    var allRichText = dataRegion.getRichTextValues();

    var colorGrid = [];
    for (var row = 0; row < allRichText.length; row++) {
      var colorRow = [];
      for (var col = 0; col < allRichText[row].length; col++) {
        var rt    = allRichText[row][col];
        var color = '#000000';
        if (rt) {
          var runs = rt.getRuns();
          for (var k = 0; k < runs.length; k++) {
            if (runs[k].getLinkUrl()) { color = '#1155cc'; break; }
          }
        }
        colorRow.push(color);
      }
      colorGrid.push(colorRow);
    }

    dataRegion.setFontColors(colorGrid);
  }

  _записатьЛист(sheetN, srcRowsN, '#1a73e8');
  _записатьЛист(sheetP, srcRowsP, '#34a853');
  _записатьЛист(sheetE, srcRowsEmpty, '#9e9e9e');

  var стОба = 0, setN = {};
  srcRowsN.forEach(function(r) { setN[r] = true; });
  srcRowsP.forEach(function(r) { if (setN[r]) стОба++; });

  ui.alert(
    '✅ ГОТОВО!\n\n'
    + '📋 Источник: '  + cfg.НАЗВАНИЕ_ЛИСТА + '\n'
    + '📊 Строки: '    + startRow + '–' + endRow + ' (' + numRows + ' строк)\n\n'
    + '📞 "' + cfg.ЛИСТ_НОМЕРА + '": ' + srcRowsN.length + ' строк\n'
    + '📧 "' + cfg.ЛИСТ_ПОЧТЫ  + '": ' + srcRowsP.length + ' строк\n'
    + '⬜ "' + cfg.ЛИСТ_ПУСТЫЕ + '": ' + srcRowsEmpty.length + ' строк\n'
    + (стОба > 0 ? '\n⚠️ Строк с обоими контактами: ' + стОба : '')
  );
}

// ══════════════════════════════════════════════════════════════════════════
// 🔗 ОБРАБОТКА ССЫЛОК
// ══════════════════════════════════════════════════════════════════════════

function обработатьСсылкиСвязь()  { _обработатьСсылки(НАСТРОЙКИ.СВЯЗЬ_ССЫЛКИ);  }
function обработатьСсылкиЯнварь() { _обработатьСсылки(НАСТРОЙКИ.ЯНВАРЬ_ССЫЛКИ); }

function _обработатьСсылки(cfg) {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(cfg.НАЗВАНИЕ_ЛИСТА);
  if (!sheet) { ui.alert('❌ Лист "' + cfg.НАЗВАНИЕ_ЛИСТА + '" не найден!'); return; }

  var lastRow      = sheet.getLastRow();
  var defaultStart = cfg.НАЧАЛЬНАЯ_СТРОКА;
  if (defaultStart > lastRow) { ui.alert('⚠️ Нет данных для обработки.'); return; }

  var ответ = ui.prompt(
    '🔗 Обработать ссылки — выбор строк',
    '📋 Лист: "' + cfg.НАЗВАНИЕ_ЛИСТА + '"\n'
    + 'Колонка: ' + cfg.КОЛОНКА + '\n'
    + 'Всего строк: ' + lastRow + '\n\n'
    + 'Введи диапазон (или оставь пустым — весь лист):\n\n'
    + '  "5:200"  — с 5 по 200\n  "50:"    — с 50 до конца\n'
    + '  ":100"   — с начала по 100\n  пусто / "всё" — с ' + defaultStart + ' по ' + lastRow,
    ui.ButtonSet.OK_CANCEL
  );
  if (ответ.getSelectedButton() !== ui.Button.OK) return;

  var ввод = ответ.getResponseText().trim().toLowerCase();
  var startRow, endRow;

  if (ввод === '' || ввод === 'всё' || ввод === 'все') {
    startRow = defaultStart; endRow = lastRow;
  } else {
    var части = ввод.split(':');
    if (части.length !== 2) { ui.alert('❌ Неверный формат. Пример: "5:200"'); return; }
    startRow = части[0] === '' ? defaultStart : parseInt(части[0], 10);
    endRow   = части[1] === '' ? lastRow      : parseInt(части[1], 10);
    if (isNaN(startRow) || isNaN(endRow)) { ui.alert('❌ Введи числа.'); return; }
    if (startRow < 1) startRow = 1;
    if (endRow > lastRow) endRow = lastRow;
    if (startRow > endRow) { ui.alert('❌ Начало больше конца.'); return; }
  }

  var numRows  = endRow - startRow + 1;
  var colIdx   = _col(cfg.КОЛОНКА);

  var range    = sheet.getRange(startRow, colIdx, numRows, 1);
  var values   = range.getValues();
  var formulas = range.getFormulas();

  var richTextValues = [];
  var fontColors     = [];
  var countWithNames = 0;
  var countUrlsOnly  = 0;
  var countSkipped   = 0;

  for (var i = 0; i < numRows; i++) {
    var rawValue   = values[i][0];
    var rawFormula = formulas[i][0] || '';

    if (!rawValue && !rawFormula) {
      richTextValues.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
      fontColors.push(['#000000']);
      countSkipped++;
      continue;
    }

    var cellText = _cellValue(String(rawValue || ''), rawFormula);
    if (!cellText || cellText.trim() === '') {
      richTextValues.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
      fontColors.push(['#000000']);
      countSkipped++;
      continue;
    }

    var url  = '';
    var name = '';

    var urlMatch = cellText.match(/https?:\/\/[^\s\)\]>]+/);
    url = urlMatch ? urlMatch[0] : cellText;

    url = url.trim()
      .replace(/^[_\s]+|[_\s]+$/g, '')
      .replace(/_+$/, '')
      .replace(/#.*$/, '')
      .replace(/\/$/, '');

    if (url.includes('instagram.com')) {
      var m = url.match(/instagram\.com\/([^\/\?#]+)/);
      if (m && m[1]) {
        var u = m[1];
        var servicePages = ['p', 'reel', 'reels', 'explore', 'stories', 'tv', 'accounts'];
        if (!servicePages.includes(u.toLowerCase())) name = u;
      }

    } else if (url.includes('facebook.com')) {
      if (!url.includes('profile.php?id=')) {
        var fbParts = url.split('/');
        var last    = fbParts[fbParts.length - 1].split('?')[0];
        var serviceWords = ['following', 'followers', 'about', 'photos', 'videos', 'posts'];
        if (last && !serviceWords.includes(last.toLowerCase())) name = last;
      }

    } else if (url.includes('youtube.com') || url.includes('youtu.be')) {
      var ymAt = url.match(/youtube\.com\/@([^\/\?#]+)/);
      if (ymAt && ymAt[1]) {
        name = ymAt[1];
      } else {
        var ymC = url.match(/youtube\.com\/(?:c|user)\/([^\/\?#]+)/);
        if (ymC && ymC[1]) name = ymC[1];
      }

    } else if (url.includes('tiktok.com')) {
      var tt = url.match(/tiktok\.com\/@([^\/\?#]+)/);
      if (tt && tt[1]) name = '@' + tt[1];

    } else if (url.includes('t.me') || url.includes('telegram.me')) {
      var tg = url.match(/(?:t\.me|telegram\.me)\/([^\/\?#]+)/);
      if (tg && tg[1]) name = tg[1];
    }

    if (name && name !== '') {
      richTextValues.push([
        SpreadsheetApp.newRichTextValue()
          .setText(String(name))
          .setLinkUrl(String(url))
          .build()
      ]);
      fontColors.push(['#1155cc']);
      countWithNames++;
    } else {
      richTextValues.push([
        SpreadsheetApp.newRichTextValue()
          .setText(String(url))
          .build()
      ]);
      fontColors.push(['#000000']);
      countUrlsOnly++;
    }
  }

  var writeRange = sheet.getRange(startRow, colIdx, numRows, 1);
  writeRange.setRichTextValues(richTextValues);
  writeRange.setFontColors(fontColors);

  ui.alert(
    '✅ ГОТОВО!\n\n'
    + '📋 Лист: '                        + cfg.НАЗВАНИЕ_ЛИСТА + '\n'
    + '📊 Строки: '                       + startRow + '–' + endRow + ' (' + numRows + ' строк)\n'
    + '🔗 Кликабельных с именем: '        + countWithNames   + '\n'
    + '📎 Оставлено как URL (без имени): ' + countUrlsOnly   + '\n'
    + '⬜ Пустых / пропущено: '            + countSkipped    + '\n\n'
    + '💡 Строки "без имени" — это profile.php, /channel/UC...\n'
    + '   и другие ссылки, где имя не угадать. Доделай вручную.'
  );
}

// ══════════════════════════════════════════════════════════════════════════
// ⚙️ ПОКАЗАТЬ НАСТРОЙКИ
// ══════════════════════════════════════════════════════════════════════════

function показатьНастройки() {
  var c = НАСТРОЙКИ;
  SpreadsheetApp.getUi().alert(
    '⚙️ ТЕКУЩИЕ НАСТРОЙКИ\n'
    + '══════════════════════════════\n\n'

    + '🧹 ДУБЛИ — Связь:\n'
    + '  Вкладка: "' + c.СВЯЗЬ_ДУБЛИ.НАЗВАНИЕ_ЛИСТА   + '"\n'
    + '  Колонка: '  + c.СВЯЗЬ_ДУБЛИ.КОЛОНКА_ССЫЛОК   + '\n'
    + '  Старт:   '  + c.СВЯЗЬ_ДУБЛИ.НАЧАЛЬНАЯ_СТРОКА + '\n\n'

    + '🧹 ДУБЛИ — Январь:\n'
    + '  Вкладка: "' + c.ЯНВАРЬ_ДУБЛИ.НАЗВАНИЕ_ЛИСТА   + '"\n'
    + '  Колонка: '  + c.ЯНВАРЬ_ДУБЛИ.КОЛОНКА_ССЫЛОК   + '\n'
    + '  Старт:   '  + c.ЯНВАРЬ_ДУБЛИ.НАЧАЛЬНАЯ_СТРОКА + '\n\n'

    + '📂 КОНТАКТЫ — Связь:\n'
    + '  Вкладка: "' + c.СВЯЗЬ_КОНТАКТЫ.НАЗВАНИЕ_ЛИСТА + '"\n'
    + '  Колонки: '  + c.СВЯЗЬ_КОНТАКТЫ.КОЛОНКА_НОМЕРОВ + ' / ' + c.СВЯЗЬ_КОНТАКТЫ.КОЛОНКА_ПОЧТ + '\n'
    + '  Диапазон: ' + c.СВЯЗЬ_КОНТАКТЫ.ДИАПАЗОН_ОТ + '→' + c.СВЯЗЬ_КОНТАКТЫ.ДИАПАЗОН_ДО + '\n\n'

    + '📂 КОНТАКТЫ — Январь:\n'
    + '  Вкладка: "' + c.ЯНВАРЬ_КОНТАКТЫ.НАЗВАНИЕ_ЛИСТА + '"\n'
    + '  Колонки: '  + c.ЯНВАРЬ_КОНТАКТЫ.КОЛОНКА_НОМЕРОВ + ' / ' + c.ЯНВАРЬ_КОНТАКТЫ.КОЛОНКА_ПОЧТ + '\n'
    + '  Диапазон: ' + c.ЯНВАРЬ_КОНТАКТЫ.ДИАПАЗОН_ОТ + '→' + c.ЯНВАРЬ_КОНТАКТЫ.ДИАПАЗОН_ДО + '\n\n'

    + '🔗 ССЫЛКИ — Связь:\n'
    + '  Вкладка: "' + c.СВЯЗЬ_ССЫЛКИ.НАЗВАНИЕ_ЛИСТА + '"\n'
    + '  Колонка: '  + c.СВЯЗЬ_ССЫЛКИ.КОЛОНКА + '\n\n'

    + '🔗 ССЫЛКИ — Январь:\n'
    + '  Вкладка: "' + c.ЯНВАРЬ_ССЫЛКИ.НАЗВАНИЕ_ЛИСТА + '"\n'
    + '  Колонка: '  + c.ЯНВАРЬ_ССЫЛКИ.КОЛОНКА + '\n\n'

    + '💡 Чтобы изменить — открой скрипт и отредактируй объект НАСТРОЙКИ.'
  );
}

// ══════════════════════════════════════════════════════════════════════════
// 🔧 ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ══════════════════════════════════════════════════════════════════════════

function _col(letter) {
  letter = letter.toUpperCase();
  var n = 0;
  for (var i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}

function _cellValue(display, formula) {
  if (formula) {
    var m = formula.match(/HYPERLINK\("([^"]+)"/i);
    if (m) return m[1];
    var u = formula.match(/[",]"?__?([a-zA-Z0-9._]+)__?"?\)/);
    if (u) return u[1];
  }
  return display ? display.replace(/^__+|__+$/g, '').trim() : '';
}

function _почта(s) {
  return s.length > 4 && /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s);
}

function _номер(s) {
  if (!s) return false;
  if (/^\d{1,2}[.\-\/]\d{1,2}[.\-\/]\d{2,4}$/.test(s.trim())) return false;
  var d = s.replace(/[\s\-\(\)\+\.]/g, '');
  return /^\d{7,15}$/.test(d);
}

function _лист(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  return sh;
}
