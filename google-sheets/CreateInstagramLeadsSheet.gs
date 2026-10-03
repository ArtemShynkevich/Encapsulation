/**
 * Создаёт таблицу «Instagram-лиды fakechat.pro».
 * Запуск: Выполнить createInstagramLeadsSpreadsheet() в редакторе Apps Script.
 */
function createInstagramLeadsSpreadsheet() {
  var ss = SpreadsheetApp.create('Instagram-лиды fakechat.pro');
  ss.setSpreadsheetLocale('ru_RU');
  ss.setSpreadsheetTimeZone('Europe/Minsk');

  var leads = ss.getSheets()[0];
  leads.setName('Лиды');
  setupLeadsSheet_(leads);

  var stats = ss.insertSheet('Статистика');
  setupStatsSheet_(stats, ss);

  var charts = ss.insertSheet('Графики');
  setupChartsSheet_(charts, ss);

  var templates = ss.insertSheet('Шаблоны');
  setupTemplatesSheet_(templates, ss);

  var settings = ss.insertSheet('Настройки');
  setupSettingsSheet_(settings);

  addTestRows_(leads);

  Logger.log('Готово: ' + ss.getUrl());
  return ss.getUrl();
}

function formulaSep_(ss) {
  var loc = ss.getSpreadsheetLocale() || 'ru_RU';
  return loc.indexOf('ru') === 0 || loc.indexOf('de') === 0 || loc.indexOf('fr') === 0 ? ';' : ',';
}

function setupLeadsSheet_(sheet) {
  var ss = sheet.getParent();
  var sep = formulaSep_(ss);
  var headers = [
    'ID', 'Дата добавления', 'Ссылка на профиль', 'Ник', 'Сегмент', 'Подписчики', 'Страна',
    'Почему подходит', 'Текст сообщения', 'Статус', 'Моя заметка', 'Лайков поставлено',
    'Дата и время отправки', 'Ответил (да/нет)', 'Дата ответа', 'Текст ответа', 'Результат', 'Шаблон'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  sheet.getRange('A2').setFormula(
    '=ARRAYFORMULA(IF(LEN(C2:C)=0' + sep + '""' + sep + ' ROW(C2:C)-1))'
  );

  var segments = ['Рилсы с переписками', 'Мем-паблик', 'SMM', 'Блогер-эксперт', 'Дизайнер/демо'];
  var countries = ['Россия', 'Беларусь'];
  var statuses = ['на одобрении', 'одобрено', 'отклонено', 'отправлено', 'ответил', 'ошибка'];
  var results = ['—', 'интерес', 'зарегистрировался', 'купил Pro', 'отказ'];
  var replied = ['да', 'нет'];

  setValidation_(sheet, 'E2:E', segments);
  setValidation_(sheet, 'G2:G', countries);
  setValidation_(sheet, 'J2:J', statuses);
  setValidation_(sheet, 'Q2:Q', results);
  setValidation_(sheet, 'N2:N', replied);

  sheet.setColumnWidth(9, 400);
  sheet.getRange('I:I').setWrap(true);

  var rules = [
    { status: 'на одобрении', color: '#fff9c4' },
    { status: 'одобрено', color: '#b3e5fc' },
    { status: 'отправлено', color: '#e0e0e0' },
    { status: 'ответил', color: '#c8e6c9' },
    { status: 'отклонено', color: '#ffcdd2' },
    { status: 'ошибка', color: '#ffcdd2' }
  ];
  var cfRules = [];
  rules.forEach(function (r) {
    cfRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$J2="' + r.status + '"')
        .setBackground(r.color)
        .setRanges([sheet.getRange('A2:R10000')])
        .build()
    );
  });
  sheet.setConditionalFormatRules(cfRules);

  sheet.getRange('A:A').protect().setWarningOnly(true).setDescription('Колонка ID — правка с предупреждением');
  sheet.getRange('M:M').protect().setWarningOnly(true).setDescription('Дата отправки — правка с предупреждением');
}

function setupStatsSheet_(sheet, ss) {
  var leads = ss.getSheetByName('Лиды').getName();
  var sep = formulaSep_(ss);
  var L = "'" + leads + "'!";

  sheet.getRange('A1').setValue('KPI').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A2:J2').setValues([[
    'Всего лидов', 'На одобрении', 'Одобрено', 'Отправлено', 'Ответили', '% ответов',
    'Зарегистрировались', 'Купили Pro', 'Конверсия в регистрацию', 'Конверсия в Pro'
  ]]).setFontWeight('bold').setWrap(true);

  sheet.getRange('A3').setFormula('=COUNTA(' + L + 'C2:C)');
  sheet.getRange('B3').setFormula('=COUNTIF(' + L + 'J:J' + sep + '"на одобрении")');
  sheet.getRange('C3').setFormula('=COUNTIF(' + L + 'J:J' + sep + '"одобрено")');
  sheet.getRange('D3').setFormula(
    '=COUNTIF(' + L + 'J:J' + sep + '"отправлено")+COUNTIF(' + L + 'J:J' + sep + '"ответил")'
  );
  sheet.getRange('E3').setFormula('=COUNTIF(' + L + 'J:J' + sep + '"ответил")');
  sheet.getRange('F3').setFormula('=IF(D3=0' + sep + '""' + sep + 'E3/D3)').setNumberFormat('0.0%');
  sheet.getRange('G3').setFormula('=COUNTIF(' + L + 'Q:Q' + sep + '"зарегистрировался")');
  sheet.getRange('H3').setFormula('=COUNTIF(' + L + 'Q:Q' + sep + '"купил Pro")');
  sheet.getRange('I3').setFormula('=IF(D3=0' + sep + '""' + sep + 'G3/D3)').setNumberFormat('0.0%');
  sheet.getRange('J3').setFormula('=IF(D3=0' + sep + '""' + sep + 'H3/D3)').setNumberFormat('0.0%');

  sheet.getRange('A5').setValue('По дням (последние 30)').setFontWeight('bold');
  sheet.getRange('A6:E6').setValues([['Дата', 'Добавлено', 'Отправлено', 'Ответов', '% ответов']])
    .setFontWeight('bold');
  sheet.getRange('A7').setFormula('=TODAY()-29');
  sheet.getRange('A8').setFormula('=A7+1');
  sheet.getRange('A9:A36').setFormula('=A8+ROW(A9:A36)-ROW(A9)');
  sheet.getRange('A7:A36').setNumberFormat('dd.mm.yyyy');

  sheet.getRange('B7').setFormula(
    '=ARRAYFORMULA(IF(A7:A36=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'B:B' + sep + '">="&A7:A36' + sep + L + 'B:B' + sep + '"<"&A7:A36+1)))'
  );
  sheet.getRange('C7').setFormula(
    '=ARRAYFORMULA(IF(A7:A36=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'M:M' + sep + '">="&A7:A36' + sep + L + 'M:M' + sep + '"<"&A7:A36+1)))'
  );
  sheet.getRange('D7').setFormula(
    '=ARRAYFORMULA(IF(A7:A36=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'O:O' + sep + '">="&A7:A36' + sep + L + 'O:O' + sep + '"<"&A7:A36+1)))'
  );
  sheet.getRange('E7').setFormula(
    '=ARRAYFORMULA(IF(A7:A36=""' + sep + '""' + sep + ' IF(C7:C36=0' + sep + '""' + sep + ' D7:D36/C7:C36)))'
  );
  sheet.getRange('E7:E36').setNumberFormat('0.0%');

  sheet.getRange('A38').setValue('По неделям').setFontWeight('bold');
  sheet.getRange('A39:E39').setValues([['Неделя (пн)', 'Добавлено', 'Отправлено', 'Ответов', '% ответов']])
    .setFontWeight('bold');
  sheet.getRange('A40').setFormula('=A7-WEEKDAY(A7' + sep + '2)+1');
  sheet.getRange('A41:A52').setFormula('=A40+7*(ROW(A41:A52)-ROW(A41))');
  sheet.getRange('A40:A52').setNumberFormat('dd.mm.yyyy');

  sheet.getRange('B40').setFormula(
    '=ARRAYFORMULA(IF(A40:A52=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'B:B' + sep + '">="&A40:A52' + sep + L + 'B:B' + sep + '"<"&A40:A52+7)))'
  );
  sheet.getRange('C40').setFormula(
    '=ARRAYFORMULA(IF(A40:A52=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'M:M' + sep + '">="&A40:A52' + sep + L + 'M:M' + sep + '"<"&A40:A52+7)))'
  );
  sheet.getRange('D40').setFormula(
    '=ARRAYFORMULA(IF(A40:A52=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'O:O' + sep + '">="&A40:A52' + sep + L + 'O:O' + sep + '"<"&A40:A52+7)))'
  );
  sheet.getRange('E40').setFormula(
    '=ARRAYFORMULA(IF(A40:A52=""' + sep + '""' + sep + ' IF(C40:C52=0' + sep + '""' + sep + ' D40:D52/C40:C52)))'
  );
  sheet.getRange('E40:E52').setNumberFormat('0.0%');

  sheet.getRange('A54').setValue('По сегментам').setFontWeight('bold');
  sheet.getRange('A55:G55').setValues([[
    'Сегмент', 'Отправлено', 'Ответов', '% ответов', 'Регистраций', 'Pro', 'Всего лидов'
  ]]).setFontWeight('bold');
  var segList = ['Рилсы с переписками', 'Мем-паблик', 'SMM', 'Блогер-эксперт', 'Дизайнер/демо'];
  sheet.getRange(56, 1, segList.length, 1).setValues(segList.map(function (s) { return [s]; }));

  sheet.getRange('B56').setFormula(
    '=ARRAYFORMULA(IF(A56:A60=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'E:E' + sep + 'A56:A60' + sep + L + 'J:J' + sep + '"отправлено")+COUNTIFS(' + L + 'E:E' + sep + 'A56:A60' + sep + L + 'J:J' + sep + '"ответил")))'
  );
  sheet.getRange('C56').setFormula(
    '=ARRAYFORMULA(IF(A56:A60=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'E:E' + sep + 'A56:A60' + sep + L + 'J:J' + sep + '"ответил")))'
  );
  sheet.getRange('D56').setFormula(
    '=ARRAYFORMULA(IF(A56:A60=""' + sep + '""' + sep + ' IF(B56:B60=0' + sep + '""' + sep + ' C56:C60/B56:B60)))'
  );
  sheet.getRange('D56:D60').setNumberFormat('0.0%');
  sheet.getRange('E56').setFormula(
    '=ARRAYFORMULA(IF(A56:A60=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'E:E' + sep + 'A56:A60' + sep + L + 'Q:Q' + sep + '"зарегистрировался")))'
  );
  sheet.getRange('F56').setFormula(
    '=ARRAYFORMULA(IF(A56:A60=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'E:E' + sep + 'A56:A60' + sep + L + 'Q:Q' + sep + '"купил Pro")))'
  );
  sheet.getRange('G56').setFormula(
    '=ARRAYFORMULA(IF(A56:A60=""' + sep + '""' + sep + ' COUNTIF(' + L + 'E:E' + sep + 'A56:A60)))'
  );

  sheet.getRange('A62').setValue('По странам').setFontWeight('bold');
  sheet.getRange('A63:F63').setValues([['Страна', 'Отправлено', 'Ответов', '% ответов', 'Регистраций', 'Pro']])
    .setFontWeight('bold');
  sheet.getRange('A64:A65').setValues([['Россия'], ['Беларусь']]);
  sheet.getRange('B64').setFormula(
    '=ARRAYFORMULA(IF(A64:A65=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'G:G' + sep + 'A64:A65' + sep + L + 'J:J' + sep + '"отправлено")+COUNTIFS(' + L + 'G:G' + sep + 'A64:A65' + sep + L + 'J:J' + sep + '"ответил")))'
  );
  sheet.getRange('C64').setFormula(
    '=ARRAYFORMULA(IF(A64:A65=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'G:G' + sep + 'A64:A65' + sep + L + 'J:J' + sep + '"ответил")))'
  );
  sheet.getRange('D64').setFormula(
    '=ARRAYFORMULA(IF(A64:A65=""' + sep + '""' + sep + ' IF(B64:B65=0' + sep + '""' + sep + ' C64:C65/B64:B65)))'
  );
  sheet.getRange('D64:D65').setNumberFormat('0.0%');
  sheet.getRange('E64').setFormula(
    '=ARRAYFORMULA(IF(A64:A65=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'G:G' + sep + 'A64:A65' + sep + L + 'Q:Q' + sep + '"зарегистрировался")))'
  );
  sheet.getRange('F64').setFormula(
    '=ARRAYFORMULA(IF(A64:A65=""' + sep + '""' + sep + ' COUNTIFS(' + L + 'G:G' + sep + 'A64:A65' + sep + L + 'Q:Q' + sep + '"купил Pro")))'
  );

  sheet.getRange('A67').setValue('Средние').setFontWeight('bold');
  sheet.getRange('A68').setValue('Среднее время до ответа (часы)');
  sheet.getRange('B68').setFormula(
    '=IFERROR(AVERAGE(FILTER((' + L + 'O:O-' + L + 'M:M)*24' + sep + '(' + L + 'O:O<>"")*(' + L + 'M:M<>"")))' + sep + '"")'
  );
  sheet.getRange('A69').setValue('Среднее лайков на лида');
  sheet.getRange('B69').setFormula(
    '=IFERROR(AVERAGE(FILTER(' + L + 'L:L' + sep + L + 'L:L<>""))' + sep + '"")'
  );

  sheet.getRange('A71').setValue('Воронка (для графика)').setFontWeight('bold');
  sheet.getRange('A72:B72').setValues([['Этап', 'Количество']]).setFontWeight('bold');
  sheet.getRange('A73:A78').setValues([
    ['Добавлено'], ['Одобрено'], ['Отправлено'], ['Ответил'], ['Регистрация'], ['Pro']
  ]);
  sheet.getRange('B73').setFormula('=A3');
  sheet.getRange('B74').setFormula(
    '=COUNTIF(' + L + 'J:J' + sep + '"одобрено")+D3'
  );
  sheet.getRange('B75').setFormula('=D3');
  sheet.getRange('B76').setFormula('=E3');
  sheet.getRange('B77').setFormula('=G3');
  sheet.getRange('B78').setFormula('=H3');

  sheet.setColumnWidths(1, 7, 120);
}

function setupChartsSheet_(sheet, ss) {
  var stats = ss.getSheetByName('Статистика');
  sheet.getRange('A1').setValue('Графики (данные — лист «Статистика»)').setFontWeight('bold');

  sheet.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(stats.getRange('A6:D36'))
    .setPosition(2, 1, 0, 0)
    .setOption('title', 'Отправлено и ответов по дням')
    .setOption('series', {
      0: { labelInLegend: 'Дата' },
      1: { labelInLegend: 'Отправлено' },
      2: { labelInLegend: 'Ответов' }
    })
    .setNumHeaders(1)
    .build());

  sheet.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(stats.getRange('A55:D60'))
    .setPosition(2, 8, 0, 0)
    .setOption('title', '% ответов по сегментам')
    .setNumHeaders(1)
    .build());

  sheet.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(stats.getRange('A72:B78'))
    .setPosition(18, 1, 0, 0)
    .setOption('title', 'Воронка: добавлено → Pro')
    .setNumHeaders(1)
    .build());

  sheet.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(stats.getRange('A56:A60'))
    .addRange(stats.getRange('G56:G60'))
    .setPosition(18, 8, 0, 0)
    .setOption('title', 'Распределение лидов по сегментам')
    .setNumHeaders(0)
    .build());

  sheet.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(stats.getRange('A39:C52'))
    .setPosition(34, 1, 0, 0)
    .setOption('title', 'Отправлено по неделям')
    .setNumHeaders(1)
    .build());
}

function setupTemplatesSheet_(sheet, ss) {
  var leads = ss.getSheetByName('Лиды').getName();
  var sep = formulaSep_(sheet.getParent());
  sheet.getRange('A1:E1').setValues([['Сегмент', 'Название шаблона', 'Текст', 'Использован раз', '% ответов']])
    .setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('A2:C4').setValues([
    ['Рилсы с переписками', 'Демо переписки', 'Привет! Делаем fakechat.pro — переписки для рилсов за минуту. Показать демо?'],
    ['Мем-паблик', 'Мем-формат', 'Привет! Видели мемы с чатами? У нас конструктор для такого контента.'],
    ['SMM', 'Агентский оффер', 'Коллеги, экономим время на визуале переписок для клиентов — fakechat.pro.']
  ]);
  sheet.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(B2:B=""' + sep + '""' + sep + ' COUNTIF(\'' + leads + '\'!R:R' + sep + 'B2:B)))'
  );
  sheet.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(B2:B=""' + sep + '""' + sep + ' IF(D2:D=0' + sep + '""' + sep + ' COUNTIFS(\'' + leads + '\'!R:R' + sep + 'B2:B' + sep + '\'' + leads + '\'!N:N' + sep + '"да")/D2:D)))'
  );
  sheet.getRange('E2:E').setNumberFormat('0.0%');
  sheet.setColumnWidth(3, 400);
  sheet.getRange('C:C').setWrap(true);
}

function setupSettingsSheet_(sheet) {
  sheet.getRange('A1:B1').setValues([['Параметр', 'Значение']]).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('A2:B6').setValues([
    ['Лимит сообщений в день', 15],
    ['Окно отправки (начало)', '10:00'],
    ['Окно отправки (конец)', '19:00'],
    ['Часовой пояс', 'Europe/Minsk'],
    ['Лайков на лида (мин–макс)', '1–5 из последних 15 постов']
  ]);
  sheet.getRange('A8').setValue('Активные сегменты').setFontWeight('bold');
  sheet.getRange('A9:A10').setValues([['Рилсы с переписками'], ['Мем-паблик']]);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 320);
}

function addTestRows_(sheet) {
  var today = new Date();
  var rows = [
    [today, 'https://instagram.com/test_reels_1', 'test_reels_1', 'Рилсы с переписками', 12500, 'Россия',
      'Рилсы с диалогами', 'Привет! ТЕСТ — fakechat.pro для ваших рилсов.', 'ответил', 'ТЕСТ', 3,
      new Date(today.getTime() - 2 * 86400000), 'да', today, 'Интересно, пришлите ссылку', 'интерес', 'Демо переписки'],
    [today, 'https://instagram.com/test_meme_2', 'test_meme_2', 'Мем-паблик', 89000, 'Беларусь',
      'Мемы с чатами', 'ТЕСТ: конструктор чатов для мемов.', 'отправлено', 'ТЕСТ', 2,
      new Date(today.getTime() - 1 * 86400000), 'нет', '', '', '—', 'Мем-формат'],
    [today, 'https://instagram.com/test_pending_3', 'test_pending_3', 'Рилсы с переписками', 4200, 'Россия',
      'Новый аккаунт рилсов', 'ТЕСТ на одобрении.', 'на одобрении', 'ТЕСТ', 1,
      '', '', '', '', '—', 'Демо переписки']
  ];
  sheet.getRange(2, 2, rows.length, 17).setValues(rows);
}

function setValidation_(sheet, a1Range, values) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(a1Range).setDataValidation(rule);
}
