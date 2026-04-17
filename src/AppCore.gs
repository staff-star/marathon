var APP_CONFIG = {
  MENU_NAME: '楽天イベント準備ツール',
  TIMEZONE: 'Asia/Tokyo',
  SHEETS: {
    INTRO: 'はじめに',
    SETTINGS: '設定',
    EVENTS: 'イベント一覧',
    IMPORT_RMS: '取込_価格表示結果',
    IMPORT_ITEM: '取込_IRアイテム',
    IMPORT_ITEMSUB: '取込_IRアイテムサブ',
    IMPORT_SELECTION: '取込_IRセレクション',
    WORK_SINGLE: '単品_作業',
    WORK_VARIATION: 'バリエーション_作業',
    OUTPUT: '出力確認',
    LOG: '処理ログ'
  },
  COLORS: {
    HEADER_BG: '#30475E',
    HEADER_FG: '#FFFFFF',
    INPUT_BG: '#FFF2CC',
    AUTO_BG: '#E7E6E6',
    WHITE: '#FFFFFF',
    NOTE_BG: '#F4F6F8'
  },
  SETTINGS_ROWS: {
    CURRENT_EVENT: 2,
    CURRENT_START_DATE: 3,
    CURRENT_END_DATE: 4,
    CURRENT_START_TIME: 5,
    CURRENT_END_TIME: 6,
    RESTORE_START_DATE: 7,
    RESTORE_END_DATE: 8,
    RESTORE_START_TIME: 9,
    RESTORE_END_TIME: 10,
    FLAG_PREFIX: 11,
    DOUBLE_PRICE_TEXT: 12,
    PRODUCT_NAME_MAX_LENGTH: 13,
    OUTPUT_ENCODING: 14
  },
  DEFAULTS: {
    FLAG_PREFIX: '二重価格OK',
    DOUBLE_PRICE_TEXT: '1',
    PRODUCT_NAME_MAX_LENGTH: 127,
    OUTPUT_ENCODING: 'Shift_JIS',
    CURRENT_START_TIME: '20:00',
    CURRENT_END_TIME: '01:59',
    RESTORE_START_TIME: '20:00',
    RESTORE_END_TIME: '01:59'
  },
  SINGLE_HEADERS: [
    '反映する',
    '割引率（%）',
    '商品コード',
    '元シート行番号',
    '商品名（元）',
    '商品名（更新後）',
    '表示価格（元）',
    '通常購入販売価格（元）',
    '通常購入販売価格（更新後）',
    '販売期間開始（更新後）',
    '販売期間終了（更新後）',
    '二重価格文言（更新後）',
    '処理状態'
  ],
  VARIATION_HEADERS: [
    '反映する',
    '割引率（%）',
    '商品コード',
    '商品名（元）',
    'SKU管理番号',
    '元シート行番号',
    '表示価格（元）',
    '通常購入販売価格（元）',
    '通常購入販売価格（更新後）',
    '処理状態'
  ]
};

var COL = {
  RMS_PRODUCT_CODE: columnToIndex_('A'),
  RMS_SKU_CODE: columnToIndex_('D'),
  RMS_RESULT: columnToIndex_('E'),
  RMS_DISPLAY_PRICE: columnToIndex_('G'),
  RMS_SALE_PRICE: columnToIndex_('H'),
  ITEM_PRODUCT_CODE: columnToIndex_('A'),
  ITEM_STOCK_TYPE: columnToIndex_('BQ'),
  ITEM_FLAG: columnToIndex_('BY'),
  ITEMSUB_PRODUCT_CODE: columnToIndex_('A'),
  ITEMSUB_NAME: columnToIndex_('I'),
  ITEMSUB_NORMAL_PRICE: columnToIndex_('L'),
  ITEMSUB_START_AT: columnToIndex_('BN'),
  ITEMSUB_END_AT: columnToIndex_('BO'),
  ITEMSUB_DISPLAY_PRICE: columnToIndex_('BS'),
  ITEMSUB_DOUBLE_PRICE_TEXT: columnToIndex_('BT'),
  SELECTION_PRODUCT_CODE: columnToIndex_('A'),
  SELECTION_NAME: columnToIndex_('B'),
  SELECTION_SKU_CODE: columnToIndex_('S'),
  SELECTION_NORMAL_PRICE: columnToIndex_('V'),
  SELECTION_DISPLAY_PRICE: columnToIndex_('W')
};

function createMenu_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(APP_CONFIG.MENU_NAME)
    .addSubMenu(
      ui.createMenu('初回設定')
        .addItem('シートを初期作成する', 'createInitialSheets')
        .addItem('説明書を再作成する', 'recreateIntroSheet')
    )
    .addSubMenu(
      ui.createMenu('設定')
        .addItem('イベント設定を開く', 'openEventSettings')
        .addItem('復旧設定を開く', 'openRestoreSettings')
    )
    .addSubMenu(
      ui.createMenu('CSV取込')
        .addItem('4つのCSVをまとめて取り込む', 'openBulkImportDialog')
        .addItem('価格表示結果CSVを取り込む', 'openImportRmsDialog')
        .addItem('IRアイテムCSVを取り込む', 'openImportItemDialog')
        .addItem('IRアイテムサブCSVを取り込む', 'openImportItemsubDialog')
        .addItem('IRセレクションCSVを取り込む', 'openImportSelectionDialog')
    )
    .addSubMenu(
      ui.createMenu('作業シート作成')
        .addItem('単品の作業シートを作成する', 'generateSingleWorkSheet')
        .addItem('バリエーションの作業シートを作成する', 'generateVariationWorkSheet')
    )
    .addSubMenu(
      ui.createMenu('更新処理')
        .addItem('IRアイテムの付箋を更新する', 'updateItemFlags')
        .addItem('単品商品の内容を反映する', 'applySingleUpdates')
        .addItem('バリエーション商品の内容を反映する', 'applyVariationUpdates')
        .addItem('すべての更新を実行する', 'runAllUpdates')
    )
    .addSubMenu(
      ui.createMenu('復旧処理')
        .addItem('単品を復旧する', 'restoreSingleProducts')
        .addItem('バリエーションを復旧する', 'restoreVariationProducts')
        .addItem('すべてを復旧する', 'restoreAllProducts')
    )
    .addSubMenu(
      ui.createMenu('CSV出力')
        .addItem('IRアイテムCSVを書き出す', 'exportItemCsv')
        .addItem('IRアイテムサブCSVを書き出す', 'exportItemsubCsv')
        .addItem('IRセレクションCSVを書き出す', 'exportSelectionCsv')
    )
    .addToUi();
}

function runMenuAction_(menuName, handler) {
  ensureBaseSheets_();
  try {
    var result = handler() || {};
    refreshOutputSummary_();
    appendProcessLog_(buildLogEntry_(menuName, result));
    SpreadsheetApp.getUi().alert(result.message || (menuName + ' を実行しました。'));
    return result;
  } catch (error) {
    appendProcessLog_({
      executedAt: new Date(),
      menuName: menuName,
      targetCount: 0,
      updatedCount: 0,
      restoredCount: 0,
      skippedCount: 0,
      errorCount: 1,
      message: truncateText_(String(error && error.message ? error.message : error), 500)
    });
    SpreadsheetApp.getUi().alert(menuName + ' でエラーが発生しました。\n' + error.message);
    throw error;
  }
}

function buildLogEntry_(menuName, result) {
  return {
    executedAt: new Date(),
    menuName: menuName,
    targetCount: Number(result.targetCount || 0),
    updatedCount: Number(result.updatedCount || 0),
    restoredCount: Number(result.restoredCount || 0),
    skippedCount: Number(result.skippedCount || 0),
    errorCount: Number(result.errorCount || 0),
    message: truncateText_(result.logMessage || result.message || '', 500)
  };
}

function ensureBaseSheets_(options) {
  options = options || {};
  var ss = getSpreadsheet_();
  setupIntroSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.INTRO), !!options.rebuildIntro);
  setupSettingsSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.SETTINGS));
  setupEventsSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.EVENTS));
  setupImportSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.IMPORT_RMS));
  setupImportSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.IMPORT_ITEM));
  setupImportSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.IMPORT_ITEMSUB));
  setupImportSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.IMPORT_SELECTION));
  setupWorkSheetShell_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.WORK_SINGLE), APP_CONFIG.SINGLE_HEADERS);
  setupWorkSheetShell_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.WORK_VARIATION), APP_CONFIG.VARIATION_HEADERS);
  setupOutputSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.OUTPUT));
  setupLogSheet_(getOrCreateSheet_(ss, APP_CONFIG.SHEETS.LOG));
}

function setupIntroSheet_(sheet, rebuild) {
  if (sheet.getLastRow() > 0 && !rebuild) {
    return;
  }
  var rows = [
    ['楽天イベント準備効率化ツール'],
    [''],
    ['目的'],
    ['楽天イベント向けの CSV 取込、作業シート生成、反映、復旧、CSV 出力を 1 つのスプレッドシートで行います。'],
    [''],
    ['毎回の操作順'],
    ['1. イベント設定を開く'],
    ['2. 価格表示結果CSVを取り込む'],
    ['3. IRアイテムCSVを取り込む'],
    ['4. IRアイテムサブCSVを取り込む'],
    ['5. IRセレクションCSVを取り込む'],
    ['6. IRアイテムの付箋を更新する'],
    ['7. 単品 / バリエーションの作業シートを作成する'],
    ['8. チェックボックスと割引率を入力する'],
    ['9. 単品 / バリエーションの反映を実行する'],
    ['10. 出力確認を見て CSV を書き出す'],
    [''],
    ['復旧の流れ'],
    ['1. 復旧設定を開く'],
    ['2. 単品 / バリエーションを復旧する'],
    ['3. 出力確認を見て CSV を書き出す'],
    [''],
    ['色の意味'],
    ['黄色：ユーザー入力欄'],
    ['灰色：自動計算 / 自動表示欄'],
    ['白色：通常表示'],
    [''],
    ['注意事項'],
    ['表示価格は変更しない'],
    ['通常購入販売価格のみ変更する'],
    ['CSVの列順を変えない'],
    ['CSVを手で並べ替えない'],
    ['復旧時、商品名の切り捨て前文字列は復元しない'],
    ['']
  ];
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sheet.getRange(3, 1, rows.length - 2, 1).setWrap(true);
  sheet.getRange(1, 1, rows.length, 1).setBackground(APP_CONFIG.COLORS.WHITE);
  sheet.autoResizeColumn(1);
}

function setupSettingsSheet_(sheet) {
  var existing = {};
  if (sheet.getLastRow() >= 14) {
    var currentValues = sheet.getRange(2, 2, 13, 1).getValues().flat();
    existing = {
      currentEvent: currentValues[0],
      currentStartDate: currentValues[1],
      currentEndDate: currentValues[2],
      currentStartTime: currentValues[3],
      currentEndTime: currentValues[4],
      restoreStartDate: currentValues[5],
      restoreEndDate: currentValues[6],
      restoreStartTime: currentValues[7],
      restoreEndTime: currentValues[8],
      flagPrefix: currentValues[9],
      doublePriceText: currentValues[10],
      productNameMaxLength: currentValues[11],
      outputEncoding: currentValues[12]
    };
  }

  var values = [
    ['項目', '値', '説明'],
    ['現在のイベント名', existing.currentEvent || '', 'イベント設定画面で更新'],
    ['現在の開始日', existing.currentStartDate || '', 'イベント設定画面で更新'],
    ['現在の終了日', existing.currentEndDate || '', 'イベント設定画面で更新'],
    ['現在の開始時刻', normalizeTimeCell_(existing.currentStartTime, APP_CONFIG.DEFAULTS.CURRENT_START_TIME), 'イベント一覧から反映'],
    ['現在の終了時刻', normalizeTimeCell_(existing.currentEndTime, APP_CONFIG.DEFAULTS.CURRENT_END_TIME), 'イベント一覧から反映'],
    ['復旧開始日', existing.restoreStartDate || '', '復旧設定画面で更新'],
    ['復旧終了日', existing.restoreEndDate || '', '復旧設定画面で更新'],
    ['復旧開始時刻', normalizeTimeCell_(existing.restoreStartTime, APP_CONFIG.DEFAULTS.RESTORE_START_TIME), '固定値'],
    ['復旧終了時刻', normalizeTimeCell_(existing.restoreEndTime, APP_CONFIG.DEFAULTS.RESTORE_END_TIME), '固定値'],
    ['付箋接頭辞', existing.flagPrefix || APP_CONFIG.DEFAULTS.FLAG_PREFIX, 'この接頭辞 + 日付で付箋更新'],
    ['二重価格文言固定値', existing.doublePriceText || APP_CONFIG.DEFAULTS.DOUBLE_PRICE_TEXT, '通常は 1'],
    ['商品名最大文字数', existing.productNameMaxLength || APP_CONFIG.DEFAULTS.PRODUCT_NAME_MAX_LENGTH, '先頭文言込みの最大文字数'],
    ['出力文字コード', existing.outputEncoding || APP_CONFIG.DEFAULTS.OUTPUT_ENCODING, 'CSV 書き出し時に使用']
  ];

  sheet.clear();
  sheet.getRange(1, 1, values.length, 3).setValues(values);
  styleHeaderRow_(sheet.getRange(1, 1, 1, 3));
  sheet.getRange(2, 1, values.length - 1, 1).setBackground(APP_CONFIG.COLORS.AUTO_BG);
  sheet.getRange(2, 2, values.length - 1, 1).setBackground(APP_CONFIG.COLORS.AUTO_BG);
  sheet.getRange(2, 3, values.length - 1, 1).setBackground(APP_CONFIG.COLORS.NOTE_BG);
  sheet.getRange(3, 2, 2, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(7, 2, 2, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(5, 2, 2, 1).setNumberFormat('hh:mm');
  sheet.getRange(9, 2, 2, 1).setNumberFormat('hh:mm');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 3, 220);
}

function setupEventsSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    var values = [
      ['イベント名', '開始時刻', '終了時刻', '使用する'],
      ['楽天マラソンセール', normalizeTimeCell_('', APP_CONFIG.DEFAULTS.CURRENT_START_TIME), normalizeTimeCell_('', APP_CONFIG.DEFAULTS.CURRENT_END_TIME), true],
      ['楽天スーパーSALE', normalizeTimeCell_('', APP_CONFIG.DEFAULTS.CURRENT_START_TIME), normalizeTimeCell_('', APP_CONFIG.DEFAULTS.CURRENT_END_TIME), true],
      ['その他イベント', normalizeTimeCell_('', APP_CONFIG.DEFAULTS.CURRENT_START_TIME), normalizeTimeCell_('', APP_CONFIG.DEFAULTS.CURRENT_END_TIME), true]
    ];
    sheet.getRange(1, 1, values.length, 4).setValues(values);
  }
  styleHeaderRow_(sheet.getRange(1, 1, 1, 4));
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 2).setNumberFormat('hh:mm');
  }
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 4, 180);
}

function setupImportSheet_(sheet) {
  if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
    sheet.clear();
  }
}

function setupWorkSheetShell_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  styleHeaderRow_(sheet.getRange(1, 1, 1, headers.length));
  sheet.setFrozenRows(1);
}

function setupOutputSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['項目', '件数', 'メモ']]);
  }
  styleHeaderRow_(sheet.getRange(1, 1, 1, 3));
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 3, 220);
}

function setupLogSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      '実行日時',
      '実行メニュー名',
      '対象件数',
      '更新件数',
      '復旧件数',
      'スキップ件数',
      'エラー件数',
      '主なエラー内容'
    ]]);
  }
  styleHeaderRow_(sheet.getRange(1, 1, 1, 8));
  sheet.setFrozenRows(1);
}

function styleHeaderRow_(range) {
  range
    .setBackground(APP_CONFIG.COLORS.HEADER_BG)
    .setFontColor(APP_CONFIG.COLORS.HEADER_FG)
    .setFontWeight('bold');
}

function appendProcessLog_(entry) {
  ensureBaseSheets_();
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.LOG);
  sheet.appendRow([
    entry.executedAt || new Date(),
    entry.menuName || '',
    Number(entry.targetCount || 0),
    Number(entry.updatedCount || 0),
    Number(entry.restoredCount || 0),
    Number(entry.skippedCount || 0),
    Number(entry.errorCount || 0),
    entry.message || ''
  ]);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function refreshOutputSummary_() {
  ensureBaseSheets_();
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.OUTPUT);
  var summary = calculateSummary_();
  var rows = [
    ['項目', '件数', 'メモ'],
    ['表示対象件数', summary.displayedCount, 'RMS で「表示」の行数'],
    ['単品対象件数', summary.singleTargetCount, '単品_作業の行数'],
    ['バリエーション対象件数', summary.variationTargetCount, 'バリエーション_作業の行数'],
    ['付箋更新件数', summary.flagTargetCount, '現在の表示対象と照合できる商品数'],
    ['単品反映予定件数', summary.singlePlannedCount, 'チェックONかつ割引率正常'],
    ['バリエーション反映予定件数', summary.variationPlannedCount, 'チェックONかつ割引率正常'],
    ['復旧対象件数', summary.restoreTargetCount, '付箋接頭辞で判定'],
    ['エラー件数', summary.errorCount, '重複と入力不備の合計'],
    ['重複件数', summary.duplicateCount, 'マスタ重複の件数'],
    ['未入力件数', summary.unfilledCount, 'チェックONだが割引率が不正な件数']
  ];
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  styleHeaderRow_(sheet.getRange(1, 1, 1, 3));
  sheet.getRange(2, 2, rows.length - 1, 1).setHorizontalAlignment('right');
  sheet.getRange(2, 1, rows.length - 1, 1).setBackground(APP_CONFIG.COLORS.AUTO_BG);
  sheet.getRange(2, 2, rows.length - 1, 1).setBackground(APP_CONFIG.COLORS.WHITE);
  sheet.getRange(2, 3, rows.length - 1, 1).setBackground(APP_CONFIG.COLORS.NOTE_BG);
  sheet.setFrozenRows(1);
}

function calculateSummary_() {
  var displayedCount = 0;
  var flagTargetCount = 0;
  var restoreTargetCount = 0;
  var duplicateCount = 0;

  try {
    var rmsRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_RMS, [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE], 'RMS価格表示結果CSV');
    displayedCount = collectDisplayedRms_(rmsRows).rawDisplayedCount;
  } catch (e) {}

  try {
    var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
    var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
    duplicateCount += itemIndex.duplicateCount;
    var settings = getSettingsValues_();
    try {
      var targetProducts = collectFlagTargetProductCodes_(
        getImportedValues_(APP_CONFIG.SHEETS.IMPORT_RMS, [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE], 'RMS価格表示結果CSV'),
        tryGetImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [COL.SELECTION_PRODUCT_CODE, COL.SELECTION_NAME, COL.SELECTION_SKU_CODE, COL.SELECTION_NORMAL_PRICE, COL.SELECTION_DISPLAY_PRICE], 'IR selection CSV')
      );
      targetProducts.productCodesInOrder.forEach(function (code) {
        if (itemIndex.map[code] && !itemIndex.duplicates[code]) {
          flagTargetCount++;
        }
      });
    } catch (ignored) {}
    Object.keys(itemIndex.map).forEach(function (code) {
      var flagValue = itemIndex.map[code].values[COL.ITEM_FLAG - 1];
      if (startsManagedFlag_(flagValue, settings.flagPrefix)) {
        restoreTargetCount++;
      }
    });
  } catch (e2) {}

  try {
    duplicateCount += indexUniqueRows_(
      getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, [
        COL.ITEMSUB_PRODUCT_CODE,
        COL.ITEMSUB_NAME,
        COL.ITEMSUB_NORMAL_PRICE,
        COL.ITEMSUB_START_AT,
        COL.ITEMSUB_END_AT,
        COL.ITEMSUB_DISPLAY_PRICE,
        COL.ITEMSUB_DOUBLE_PRICE_TEXT
      ], 'IR itemsub CSV'),
      COL.ITEMSUB_PRODUCT_CODE
    ).duplicateCount;
  } catch (e3) {}

  try {
    duplicateCount += indexUniqueRows_(
      getImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [
        COL.SELECTION_PRODUCT_CODE,
        COL.SELECTION_NAME,
        COL.SELECTION_SKU_CODE,
        COL.SELECTION_NORMAL_PRICE,
        COL.SELECTION_DISPLAY_PRICE
      ], 'IR selection CSV'),
      COL.SELECTION_SKU_CODE
    ).duplicateCount;
  } catch (e4) {}

  var singleStats = getWorkSheetStats_(APP_CONFIG.SHEETS.WORK_SINGLE, true);
  var variationStats = getWorkSheetStats_(APP_CONFIG.SHEETS.WORK_VARIATION, false);

  return {
    displayedCount: displayedCount,
    singleTargetCount: singleStats.totalRows,
    variationTargetCount: variationStats.totalRows,
    flagTargetCount: flagTargetCount,
    singlePlannedCount: singleStats.readyCount,
    variationPlannedCount: variationStats.readyCount,
    restoreTargetCount: restoreTargetCount,
    errorCount: duplicateCount + singleStats.invalidCount + variationStats.invalidCount,
    duplicateCount: duplicateCount,
    unfilledCount: singleStats.invalidCount + variationStats.invalidCount
  };
}

function getWorkSheetStats_(sheetName, isSingle) {
  var sheet = getSheetOrThrow_(sheetName);
  if (sheet.getLastRow() <= 1) {
    return { totalRows: 0, readyCount: 0, invalidCount: 0 };
  }
  var columnCount = isSingle ? APP_CONFIG.SINGLE_HEADERS.length : APP_CONFIG.VARIATION_HEADERS.length;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, columnCount).getValues();
  var totalRows = 0;
  var readyCount = 0;
  var invalidCount = 0;
  values.forEach(function (row) {
    if (!normalizeString_(row[2])) {
      return;
    }
    totalRows++;
    if (row[0] === true) {
      if (isValidDiscountInteger_(row[1])) {
        readyCount++;
      } else {
        invalidCount++;
      }
    }
  });
  return {
    totalRows: totalRows,
    readyCount: readyCount,
    invalidCount: invalidCount
  };
}

function getEventDialogPayload_() {
  var settings = getSettingsValues_();
  return {
    events: getEnabledEvents_(),
    currentEventName: settings.currentEvent,
    currentStartDate: formatDateForInput_(settings.currentStartDate),
    currentEndDate: formatDateForInput_(settings.currentEndDate),
    currentStartTime: formatTimeForDisplay_(settings.currentStartTime),
    currentEndTime: formatTimeForDisplay_(settings.currentEndTime)
  };
}

function getRestoreDialogPayload_() {
  var settings = getSettingsValues_();
  return {
    restoreStartDate: formatDateForInput_(settings.restoreStartDate),
    restoreEndDate: formatDateForInput_(settings.restoreEndDate),
    restoreStartTime: formatTimeForDisplay_(settings.restoreStartTime),
    restoreEndTime: formatTimeForDisplay_(settings.restoreEndTime)
  };
}

function saveEventSettings_(payload) {
  var eventName = normalizeString_(payload && payload.eventName);
  var startDate = parseDateInput_(payload && payload.startDate);
  var endDate = parseDateInput_(payload && payload.endDate);
  if (!eventName) {
    throw new Error('イベント名を選択してください。');
  }
  if (!startDate || !endDate) {
    throw new Error('開始日と終了日を入力してください。');
  }

  var events = getEnabledEvents_();
  var matched = events.filter(function (event) {
    return event.name === eventName;
  })[0];
  if (!matched) {
    throw new Error('イベント一覧に存在する有効なイベント名を選択してください。');
  }

  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.SETTINGS);
  sheet.getRange(APP_CONFIG.SETTINGS_ROWS.CURRENT_EVENT, 2).setValue(eventName);
  sheet.getRange(APP_CONFIG.SETTINGS_ROWS.CURRENT_START_DATE, 2).setValue(startDate);
  sheet.getRange(APP_CONFIG.SETTINGS_ROWS.CURRENT_END_DATE, 2).setValue(endDate);
  sheet.getRange(APP_CONFIG.SETTINGS_ROWS.CURRENT_START_TIME, 2).setValue(makeTimeValue_(matched.startTime));
  sheet.getRange(APP_CONFIG.SETTINGS_ROWS.CURRENT_END_TIME, 2).setValue(makeTimeValue_(matched.endTime));
  return {
    message: 'イベント設定を保存しました。',
    updatedCount: 1,
    targetCount: 1,
    errorCount: 0
  };
}

function saveRestoreSettings_(payload) {
  var startDate = parseDateInput_(payload && payload.startDate);
  var endDate = parseDateInput_(payload && payload.endDate);
  if (!startDate || !endDate) {
    throw new Error('復旧開始日と復旧終了日を入力してください。');
  }
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.SETTINGS);
  sheet.getRange(APP_CONFIG.SETTINGS_ROWS.RESTORE_START_DATE, 2).setValue(startDate);
  sheet.getRange(APP_CONFIG.SETTINGS_ROWS.RESTORE_END_DATE, 2).setValue(endDate);
  return {
    message: '復旧設定を保存しました。',
    updatedCount: 1,
    targetCount: 1,
    errorCount: 0
  };
}

function getEnabledEvents_() {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.EVENTS);
  if (sheet.getLastRow() <= 1) {
    return [];
  }
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return values.filter(function (row) {
    return normalizeString_(row[0]) && asBoolean_(row[3]);
  }).map(function (row) {
    return {
      name: normalizeString_(row[0]),
      startTime: formatTimeForDisplay_(row[1]),
      endTime: formatTimeForDisplay_(row[2])
    };
  });
}

function showImportDialog_(target) {
  ensureBaseSheets_();
  var config = getImportConfig_(target);
  var template = HtmlService.createTemplateFromFile('ImportDialog');
  template.title = config.dialogTitle;
  template.target = target;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(260),
    config.dialogTitle
  );
}

function getImportConfig_(target) {
  var map = {
    rms: {
      sheetName: APP_CONFIG.SHEETS.IMPORT_RMS,
      menuLabel: '価格表示結果CSVを取り込む',
      dialogTitle: '価格表示結果CSV取込',
      requiredColumns: [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE],
      label: 'RMS価格表示結果CSV'
    },
    item: {
      sheetName: APP_CONFIG.SHEETS.IMPORT_ITEM,
      menuLabel: 'IRアイテムCSVを取り込む',
      dialogTitle: 'IRアイテムCSV取込',
      requiredColumns: [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG],
      label: 'IR item CSV'
    },
    itemsub: {
      sheetName: APP_CONFIG.SHEETS.IMPORT_ITEMSUB,
      menuLabel: 'IRアイテムサブCSVを取り込む',
      dialogTitle: 'IRアイテムサブCSV取込',
      requiredColumns: [
        COL.ITEMSUB_PRODUCT_CODE,
        COL.ITEMSUB_NAME,
        COL.ITEMSUB_NORMAL_PRICE,
        COL.ITEMSUB_START_AT,
        COL.ITEMSUB_END_AT,
        COL.ITEMSUB_DISPLAY_PRICE,
        COL.ITEMSUB_DOUBLE_PRICE_TEXT
      ],
      label: 'IR itemsub CSV'
    },
    selection: {
      sheetName: APP_CONFIG.SHEETS.IMPORT_SELECTION,
      menuLabel: 'IRセレクションCSVを取り込む',
      dialogTitle: 'IRセレクションCSV取込',
      requiredColumns: [COL.SELECTION_PRODUCT_CODE, COL.SELECTION_NAME, COL.SELECTION_SKU_CODE, COL.SELECTION_NORMAL_PRICE, COL.SELECTION_DISPLAY_PRICE],
      label: 'IR selection CSV'
    }
  };
  if (!map[target]) {
    throw new Error('不明な取込種別です。');
  }
  return map[target];
}

function importCsvPayload_(payload) {
  var config = getImportConfig_(payload && payload.target);
  var fileName = normalizeString_(payload && payload.fileName);
  var base64 = normalizeString_(payload && payload.base64);
  if (!fileName || !base64) {
    throw new Error('CSV ファイルを選択してください。');
  }

  var decoded = decodeCsvBytes_(Utilities.base64Decode(stripDataUrlPrefix_(base64)));
  var rows = Utilities.parseCsv(stripUtf8Bom_(decoded.text));
  if (!rows.length) {
    throw new Error('CSV を読み込めませんでした。');
  }

  validateImportedSheetStructure_(rows, config.requiredColumns, config.label);
  var paddedRows = padRows_(rows);
  var sheet = getSheetOrThrow_(config.sheetName);
  sheet.clear();
  sheet.getRange(1, 1, paddedRows.length, paddedRows[0].length).setValues(paddedRows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, paddedRows[0].length).setFontWeight('bold');
  return {
    message: config.label + ' を取り込みました。\n件数: ' + Math.max(rows.length - 1, 0) + '件\n判定文字コード: ' + decoded.encoding,
    targetCount: Math.max(rows.length - 1, 0),
    updatedCount: Math.max(rows.length - 1, 0),
    restoredCount: 0,
    skippedCount: 0,
    errorCount: 0
  };
}

function importMultipleCsvPayloads_(payloads) {
  var list = (payloads || []).filter(function (payload) {
    return payload && payload.target;
  });
  if (!list.length) {
    throw new Error('少なくとも1つはCSVを選択してください。');
  }
  var messages = [];
  var totalCount = 0;
  list.forEach(function (payload) {
    var result = importCsvPayload_(payload);
    totalCount += Number(result.updatedCount || 0);
    messages.push(getImportConfig_(payload.target).label + ': ' + Number(result.updatedCount || 0) + '件');
  });
  return {
    message: '一括取込が完了しました。\n' + messages.join('\n'),
    targetCount: totalCount,
    updatedCount: totalCount,
    restoredCount: 0,
    skippedCount: 0,
    errorCount: 0
  };
}

function generateSingleWorkSheetCore_() {
  var rmsRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_RMS, [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE], 'RMS価格表示結果CSV');
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var itemsubRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, [
    COL.ITEMSUB_PRODUCT_CODE,
    COL.ITEMSUB_NAME,
    COL.ITEMSUB_NORMAL_PRICE,
    COL.ITEMSUB_START_AT,
    COL.ITEMSUB_END_AT,
    COL.ITEMSUB_DISPLAY_PRICE,
    COL.ITEMSUB_DOUBLE_PRICE_TEXT
  ], 'IR itemsub CSV');

  var displayed = collectDisplayedRms_(rmsRows);
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var itemsubIndex = indexUniqueRows_(itemsubRows, COL.ITEMSUB_PRODUCT_CODE);
  var records = [];
  var errorCount = itemIndex.duplicateCount + itemsubIndex.duplicateCount;
  var skippedCount = 0;

  displayed.productCodesInOrder.forEach(function (productCode) {
    var itemEntry = itemIndex.map[productCode];
    if (!itemEntry || itemIndex.duplicates[productCode]) {
      skippedCount++;
      return;
    }
    if (normalizeString_(itemEntry.values[COL.ITEM_STOCK_TYPE - 1]) !== '1') {
      return;
    }
    var itemsubEntry = itemsubIndex.map[productCode];
    if (!itemsubEntry || itemsubIndex.duplicates[productCode]) {
      skippedCount++;
      return;
    }
    var displayPrice = toNumber_(itemsubEntry.values[COL.ITEMSUB_DISPLAY_PRICE - 1]);
    if (displayPrice === null) {
      errorCount++;
      skippedCount++;
      return;
    }
    records.push([
      false,
      '',
      productCode,
      itemsubEntry.rowNumber,
      normalizeString_(itemsubEntry.values[COL.ITEMSUB_NAME - 1]),
      '',
      displayPrice,
      toSheetValue_(itemsubEntry.values[COL.ITEMSUB_NORMAL_PRICE - 1]),
      '',
      '',
      '',
      '',
      ''
    ]);
  });

  writeSingleWorkSheet_(records);
  return {
    targetCount: records.length,
    updatedCount: records.length,
    restoredCount: 0,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: '単品_作業シートを作成しました。\n対象: ' + records.length + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}

function writeSingleWorkSheet_(records) {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.WORK_SINGLE);
  sheet.clear();
  sheet.getRange(1, 1, 1, APP_CONFIG.SINGLE_HEADERS.length).setValues([APP_CONFIG.SINGLE_HEADERS]);
  styleHeaderRow_(sheet.getRange(1, 1, 1, APP_CONFIG.SINGLE_HEADERS.length));
  sheet.setFrozenRows(1);
  if (!records.length) {
    return;
  }
  sheet.getRange(2, 1, records.length, APP_CONFIG.SINGLE_HEADERS.length).setValues(records);
  sheet.getRange(2, 1, records.length, 1).insertCheckboxes();
  var nameFormulas = [];
  var priceFormulas = [];
  var startFormulas = [];
  var endFormulas = [];
  var doublePriceFormulas = [];
  var statusFormulas = [];
  for (var i = 0; i < records.length; i++) {
    var row = i + 2;
    nameFormulas.push([buildSingleNameFormula_(row)]);
    priceFormulas.push([buildSinglePriceFormula_(row)]);
    startFormulas.push([buildCurrentStartFormula_(row, 'I')]);
    endFormulas.push([buildCurrentEndFormula_(row, 'I')]);
    doublePriceFormulas.push([buildDoublePriceFormula_(row, 'I')]);
    statusFormulas.push([buildSingleStatusFormula_(row)]);
  }
  sheet.getRange(2, 6, records.length, 1).setFormulas(nameFormulas);
  sheet.getRange(2, 9, records.length, 1).setFormulas(priceFormulas);
  sheet.getRange(2, 10, records.length, 1).setFormulas(startFormulas);
  sheet.getRange(2, 11, records.length, 1).setFormulas(endFormulas);
  sheet.getRange(2, 12, records.length, 1).setFormulas(doublePriceFormulas);
  sheet.getRange(2, 13, records.length, 1).setFormulas(statusFormulas);
  sheet.setColumnWidths(1, 13, 130);
  sheet.getRange(2, 1, records.length, 2).setBackground(APP_CONFIG.COLORS.INPUT_BG);
  sheet.getRange(2, 3, records.length, 11).setBackground(APP_CONFIG.COLORS.AUTO_BG);
  sheet.getRange(2, 4, records.length, 1).setNumberFormat('0');
  sheet.getRange(2, 7, records.length, 3).setNumberFormat('0');
}

function generateVariationWorkSheetCore_() {
  var rmsRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_RMS, [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE], 'RMS価格表示結果CSV');
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var selectionRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [COL.SELECTION_PRODUCT_CODE, COL.SELECTION_NAME, COL.SELECTION_SKU_CODE, COL.SELECTION_NORMAL_PRICE, COL.SELECTION_DISPLAY_PRICE], 'IR selection CSV');

  var displayed = collectDisplayedRms_(rmsRows);
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var selectionIndex = indexUniqueRows_(selectionRows, COL.SELECTION_SKU_CODE);
  var records = [];
  var errorCount = itemIndex.duplicateCount + selectionIndex.duplicateCount;
  var skippedCount = 0;

  displayed.skuCodesInOrder.forEach(function (skuCode) {
    var selectionEntry = selectionIndex.map[skuCode];
    if (!selectionEntry || selectionIndex.duplicates[skuCode]) {
      skippedCount++;
      return;
    }
    var productCode = normalizeString_(selectionEntry.values[COL.SELECTION_PRODUCT_CODE - 1]);
    var itemEntry = itemIndex.map[productCode];
    if (!itemEntry || itemIndex.duplicates[productCode]) {
      skippedCount++;
      return;
    }
    if (normalizeString_(itemEntry.values[COL.ITEM_STOCK_TYPE - 1]) !== '2') {
      return;
    }
    var displayPrice = toNumber_(selectionEntry.values[COL.SELECTION_DISPLAY_PRICE - 1]);
    if (displayPrice === null) {
      errorCount++;
      skippedCount++;
      return;
    }
    records.push([
      false,
      '',
      productCode,
      normalizeString_(selectionEntry.values[COL.SELECTION_NAME - 1]),
      skuCode,
      selectionEntry.rowNumber,
      displayPrice,
      toSheetValue_(selectionEntry.values[COL.SELECTION_NORMAL_PRICE - 1]),
      '',
      ''
    ]);
  });

  writeVariationWorkSheet_(records);
  return {
    targetCount: records.length,
    updatedCount: records.length,
    restoredCount: 0,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: 'バリエーション_作業シートを作成しました。\n対象: ' + records.length + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}

function writeVariationWorkSheet_(records) {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.WORK_VARIATION);
  sheet.clear();
  sheet.getRange(1, 1, 1, APP_CONFIG.VARIATION_HEADERS.length).setValues([APP_CONFIG.VARIATION_HEADERS]);
  styleHeaderRow_(sheet.getRange(1, 1, 1, APP_CONFIG.VARIATION_HEADERS.length));
  sheet.setFrozenRows(1);
  if (!records.length) {
    return;
  }
  sheet.getRange(2, 1, records.length, APP_CONFIG.VARIATION_HEADERS.length).setValues(records);
  sheet.getRange(2, 1, records.length, 1).insertCheckboxes();
  var priceFormulas = [];
  var statusFormulas = [];
  for (var i = 0; i < records.length; i++) {
    var row = i + 2;
    priceFormulas.push([buildVariationPriceFormula_(row)]);
    statusFormulas.push([buildVariationStatusFormula_(row)]);
  }
  sheet.getRange(2, 9, records.length, 1).setFormulas(priceFormulas);
  sheet.getRange(2, 10, records.length, 1).setFormulas(statusFormulas);
  sheet.setColumnWidths(1, 10, 140);
  sheet.getRange(2, 1, records.length, 2).setBackground(APP_CONFIG.COLORS.INPUT_BG);
  sheet.getRange(2, 3, records.length, 8).setBackground(APP_CONFIG.COLORS.AUTO_BG);
  sheet.getRange(2, 6, records.length, 1).setNumberFormat('0');
  sheet.getRange(2, 7, records.length, 3).setNumberFormat('0');
}

function updateItemFlagsCore_() {
  var rmsRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_RMS, [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE], 'RMS価格表示結果CSV');
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var selectionRows = tryGetImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [COL.SELECTION_PRODUCT_CODE, COL.SELECTION_NAME, COL.SELECTION_SKU_CODE, COL.SELECTION_NORMAL_PRICE, COL.SELECTION_DISPLAY_PRICE], 'IR selection CSV');
  var targetProducts = collectFlagTargetProductCodes_(rmsRows, selectionRows);
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var settings = getSettingsValues_();
  var flagValue = buildFlagValue_(settings.flagPrefix);
  var itemValues = itemRows.values;
  var updatedCount = 0;
  var clearedCount = 0;

  for (var i = 1; i < itemValues.length; i++) {
    var row = itemValues[i];
    var productCode = normalizeString_(row[COL.ITEM_PRODUCT_CODE - 1]);
    if (!productCode || itemIndex.duplicates[productCode]) {
      continue;
    }
    if (targetProducts.productCodeSet[productCode]) {
      if (row[COL.ITEM_FLAG - 1] !== flagValue) {
        row[COL.ITEM_FLAG - 1] = flagValue;
        updatedCount++;
      }
    } else if (startsManagedFlag_(row[COL.ITEM_FLAG - 1], settings.flagPrefix)) {
      row[COL.ITEM_FLAG - 1] = '';
      clearedCount++;
    }
  }

  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, itemValues);
  return {
    targetCount: targetProducts.productCodesInOrder.length,
    updatedCount: updatedCount + clearedCount,
    restoredCount: 0,
    skippedCount: itemIndex.duplicateCount + targetProducts.duplicateCount,
    errorCount: itemIndex.duplicateCount + targetProducts.duplicateCount,
    message: 'IR item の付箋を更新しました。\n上書き: ' + updatedCount + '件\nクリア: ' + clearedCount + '件\n重複エラー: ' + (itemIndex.duplicateCount + targetProducts.duplicateCount) + '件'
  };
}

function applySingleUpdatesCore_() {
  var settings = ensureCurrentSettingsComplete_();
  var workRows = getWorkSheetRows_(APP_CONFIG.SHEETS.WORK_SINGLE, 13, '単品_作業');
  var itemsubRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, [
    COL.ITEMSUB_PRODUCT_CODE,
    COL.ITEMSUB_NAME,
    COL.ITEMSUB_NORMAL_PRICE,
    COL.ITEMSUB_START_AT,
    COL.ITEMSUB_END_AT,
    COL.ITEMSUB_DISPLAY_PRICE,
    COL.ITEMSUB_DOUBLE_PRICE_TEXT
  ], 'IR itemsub CSV');
  var values = itemsubRows.values;
  var startAt = formatDateTimeForCsv_(settings.currentStartDate, settings.currentStartTime);
  var endAt = formatDateTimeForCsv_(settings.currentEndDate, settings.currentEndTime);
  var updatedCount = 0;
  var skippedCount = 0;
  var errorCount = 0;
  var touched = {};

  workRows.forEach(function (row) {
    if (row[0] !== true) {
      skippedCount++;
      return;
    }
    if (!isValidDiscountInteger_(row[1])) {
      skippedCount++;
      errorCount++;
      return;
    }
    var sourceRowNumber = Number(row[3]);
    if (!sourceRowNumber || sourceRowNumber < 2 || sourceRowNumber > values.length || touched[sourceRowNumber]) {
      skippedCount++;
      errorCount++;
      return;
    }
    var targetRow = values[sourceRowNumber - 1];
    var displayPrice = toNumber_(targetRow[COL.ITEMSUB_DISPLAY_PRICE - 1]);
    if (displayPrice === null) {
      skippedCount++;
      errorCount++;
      return;
    }
    var newPrice = calculateDiscountedPrice_(displayPrice, Number(row[1]));
    targetRow[COL.ITEMSUB_NAME - 1] = buildSingleName_(settings.currentEvent, displayPrice, newPrice, normalizeString_(targetRow[COL.ITEMSUB_NAME - 1]), settings.productNameMaxLength);
    targetRow[COL.ITEMSUB_NORMAL_PRICE - 1] = newPrice;
    targetRow[COL.ITEMSUB_START_AT - 1] = startAt;
    targetRow[COL.ITEMSUB_END_AT - 1] = endAt;
    targetRow[COL.ITEMSUB_DOUBLE_PRICE_TEXT - 1] = settings.doublePriceText;
    touched[sourceRowNumber] = true;
    updatedCount++;
  });

  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, values);
  return {
    targetCount: workRows.length,
    updatedCount: updatedCount,
    restoredCount: 0,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: '単品商品の内容を反映しました。\n更新: ' + updatedCount + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}

function applyVariationUpdatesCore_() {
  getSettingsValues_();
  var workRows = getWorkSheetRows_(APP_CONFIG.SHEETS.WORK_VARIATION, APP_CONFIG.VARIATION_HEADERS.length, 'バリエーション_作業');
  var selectionRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [COL.SELECTION_PRODUCT_CODE, COL.SELECTION_NAME, COL.SELECTION_SKU_CODE, COL.SELECTION_NORMAL_PRICE, COL.SELECTION_DISPLAY_PRICE], 'IR selection CSV');
  var values = selectionRows.values;
  var updatedCount = 0;
  var skippedCount = 0;
  var errorCount = 0;
  var touched = {};

  workRows.forEach(function (row) {
    if (row[0] !== true) {
      skippedCount++;
      return;
    }
    if (!isValidDiscountInteger_(row[1])) {
      skippedCount++;
      errorCount++;
      return;
    }
    var sourceRowNumber = Number(row[5]);
    if (!sourceRowNumber || sourceRowNumber < 2 || sourceRowNumber > values.length || touched[sourceRowNumber]) {
      skippedCount++;
      errorCount++;
      return;
    }
    var targetRow = values[sourceRowNumber - 1];
    var displayPrice = toNumber_(targetRow[COL.SELECTION_DISPLAY_PRICE - 1]);
    if (displayPrice === null) {
      skippedCount++;
      errorCount++;
      return;
    }
    targetRow[COL.SELECTION_NORMAL_PRICE - 1] = calculateDiscountedPrice_(displayPrice, Number(row[1]));
    touched[sourceRowNumber] = true;
    updatedCount++;
  });

  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, values);
  return {
    targetCount: workRows.length,
    updatedCount: updatedCount,
    restoredCount: 0,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: 'バリエーション商品の内容を反映しました。\n更新: ' + updatedCount + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}

function restoreSingleProductsCore_() {
  var settings = ensureRestoreSettingsComplete_();
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var itemsubRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, [
    COL.ITEMSUB_PRODUCT_CODE,
    COL.ITEMSUB_NAME,
    COL.ITEMSUB_NORMAL_PRICE,
    COL.ITEMSUB_START_AT,
    COL.ITEMSUB_END_AT,
    COL.ITEMSUB_DISPLAY_PRICE,
    COL.ITEMSUB_DOUBLE_PRICE_TEXT
  ], 'IR itemsub CSV');
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var itemsubIndex = indexUniqueRows_(itemsubRows, COL.ITEMSUB_PRODUCT_CODE);
  var values = itemsubRows.values;
  var targetCodes = [];

  Object.keys(itemIndex.map).forEach(function (productCode) {
    if (itemIndex.duplicates[productCode]) {
      return;
    }
    var row = itemIndex.map[productCode].values;
    if (normalizeString_(row[COL.ITEM_STOCK_TYPE - 1]) === '1' && startsManagedFlag_(row[COL.ITEM_FLAG - 1], settings.flagPrefix)) {
      targetCodes.push(productCode);
    }
  });

  var restoreStart = formatDateTimeForCsv_(settings.restoreStartDate, settings.restoreStartTime);
  var restoreEnd = formatDateTimeForCsv_(settings.restoreEndDate, settings.restoreEndTime);
  var restoredCount = 0;
  var skippedCount = itemIndex.duplicateCount + itemsubIndex.duplicateCount;
  var errorCount = itemIndex.duplicateCount + itemsubIndex.duplicateCount;

  targetCodes.forEach(function (productCode) {
    var itemsubEntry = itemsubIndex.map[productCode];
    if (!itemsubEntry || itemsubIndex.duplicates[productCode]) {
      skippedCount++;
      errorCount++;
      return;
    }
    var row = values[itemsubEntry.rowNumber - 1];
    row[COL.ITEMSUB_NORMAL_PRICE - 1] = row[COL.ITEMSUB_DISPLAY_PRICE - 1];
    row[COL.ITEMSUB_NAME - 1] = stripSalePrefix_(normalizeString_(row[COL.ITEMSUB_NAME - 1]));
    row[COL.ITEMSUB_START_AT - 1] = restoreStart;
    row[COL.ITEMSUB_END_AT - 1] = restoreEnd;
    row[COL.ITEMSUB_DOUBLE_PRICE_TEXT - 1] = settings.doublePriceText;
    restoredCount++;
  });

  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, values);
  return {
    targetCount: targetCodes.length,
    updatedCount: 0,
    restoredCount: restoredCount,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: '単品を復旧しました。\n復旧: ' + restoredCount + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}

function restoreVariationProductsCore_() {
  var settings = ensureRestoreSettingsComplete_();
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var selectionRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [COL.SELECTION_PRODUCT_CODE, COL.SELECTION_NAME, COL.SELECTION_SKU_CODE, COL.SELECTION_NORMAL_PRICE, COL.SELECTION_DISPLAY_PRICE], 'IR selection CSV');
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var selectionIndex = indexUniqueRows_(selectionRows, COL.SELECTION_SKU_CODE);
  var values = selectionRows.values;
  var targetProductCodes = {};

  Object.keys(itemIndex.map).forEach(function (productCode) {
    if (itemIndex.duplicates[productCode]) {
      return;
    }
    var row = itemIndex.map[productCode].values;
    if (normalizeString_(row[COL.ITEM_STOCK_TYPE - 1]) === '2' && startsManagedFlag_(row[COL.ITEM_FLAG - 1], settings.flagPrefix)) {
      targetProductCodes[productCode] = true;
    }
  });

  var restoredCount = 0;
  var skippedCount = itemIndex.duplicateCount;
  var errorCount = itemIndex.duplicateCount + selectionIndex.duplicateCount;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var productCode = normalizeString_(row[COL.SELECTION_PRODUCT_CODE - 1]);
    var skuCode = normalizeString_(row[COL.SELECTION_SKU_CODE - 1]);
    if (!productCode || !targetProductCodes[productCode]) {
      continue;
    }
    if (selectionIndex.duplicates[skuCode]) {
      skippedCount++;
      continue;
    }
    row[COL.SELECTION_NORMAL_PRICE - 1] = row[COL.SELECTION_DISPLAY_PRICE - 1];
    restoredCount++;
  }

  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, values);
  return {
    targetCount: Object.keys(targetProductCodes).length,
    updatedCount: 0,
    restoredCount: restoredCount,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: 'バリエーションを復旧しました。\n復旧: ' + restoredCount + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}

function showExportDialog_(kind) {
  ensureBaseSheets_();
  var payload = buildExportPayload_(kind);
  var template = HtmlService.createTemplateFromFile('DownloadDialog');
  template.fileName = payload.fileName;
  template.base64 = payload.base64;
  template.description = payload.description;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(230),
    'CSVダウンロード'
  );
}

function buildExportPayload_(kind) {
  var configMap = {
    item: { sheetName: APP_CONFIG.SHEETS.IMPORT_ITEM, prefix: 'ir-item' },
    itemsub: { sheetName: APP_CONFIG.SHEETS.IMPORT_ITEMSUB, prefix: 'ir-itemsub' },
    selection: { sheetName: APP_CONFIG.SHEETS.IMPORT_SELECTION, prefix: 'ir-selection' }
  };
  var config = configMap[kind];
  if (!config) {
    throw new Error('不明な出力種別です。');
  }
  var sheet = getSheetOrThrow_(config.sheetName);
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    throw new Error('出力対象のデータがありません。');
  }
  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  var fileName = config.prefix + '_更新済み_' + Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss') + '.csv';
  var blob = Utilities.newBlob(buildCsvText_(values), 'Shift_JIS', fileName);
  return {
    fileName: fileName,
    base64: Utilities.base64Encode(blob.getBytes()),
    description: config.sheetName + ' を Shift_JIS / CRLF で出力します。'
  };
}

function buildCsvText_(rows) {
  return rows.map(function (row) {
    return row.map(csvEscape_).join(',');
  }).join('\r\n') + '\r\n';
}

function csvEscape_(value) {
  var text = value == null ? '' : String(value);
  if (/[",\r\n]/.test(text)) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('アクティブなスプレッドシートが見つかりません。');
  }
  return ss;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getSheetOrThrow_(name) {
  var sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + name);
  }
  return sheet;
}

function getImportedValues_(sheetName, requiredColumns, label) {
  var sheet = getSheetOrThrow_(sheetName);
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    throw new Error(label + ' が未取込です。');
  }
  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  validateImportedSheetStructure_(values, requiredColumns, label);
  return {
    values: values,
    header: values[0]
  };
}

function tryGetImportedValues_(sheetName, requiredColumns, label) {
  try {
    return getImportedValues_(sheetName, requiredColumns, label);
  } catch (error) {
    return null;
  }
}

function writeBackImportedValues_(sheetName, values) {
  var sheet = getSheetOrThrow_(sheetName);
  sheet.clear();
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, values[0].length).setFontWeight('bold');
}

function validateImportedSheetStructure_(rows, requiredColumns, label) {
  if (!rows || !rows.length) {
    throw new Error(label + ' が空です。');
  }
  var header = rows[0];
  requiredColumns.forEach(function (columnIndex) {
    if (header.length < columnIndex || normalizeString_(header[columnIndex - 1]) === '') {
      throw new Error(label + ' の必須列が見つかりません。列: ' + indexToColumn_(columnIndex));
    }
  });
}

function collectDisplayedRms_(rmsRows) {
  var productCodeSet = {};
  var skuSet = {};
  var productCodesInOrder = [];
  var skuCodesInOrder = [];
  var rawDisplayedCount = 0;
  rmsRows.values.slice(1).forEach(function (row) {
    if (normalizeString_(row[COL.RMS_RESULT - 1]) !== '表示') {
      return;
    }
    rawDisplayedCount++;
    var productCode = normalizeString_(row[COL.RMS_PRODUCT_CODE - 1]);
    var skuCode = normalizeString_(row[COL.RMS_SKU_CODE - 1]);
    if (productCode && !productCodeSet[productCode]) {
      productCodeSet[productCode] = true;
      productCodesInOrder.push(productCode);
    }
    if (skuCode && !skuSet[skuCode]) {
      skuSet[skuCode] = true;
      skuCodesInOrder.push(skuCode);
    }
  });
  return {
    rawDisplayedCount: rawDisplayedCount,
    productCodesInOrder: productCodesInOrder,
    skuCodesInOrder: skuCodesInOrder,
    productCodeSet: productCodeSet,
    skuSet: skuSet
  };
}

function collectFlagTargetProductCodes_(rmsRows, selectionRows) {
  var displayed = collectDisplayedRms_(rmsRows);
  var productCodeSet = {};
  var productCodesInOrder = [];
  var duplicateCount = 0;

  displayed.productCodesInOrder.forEach(function (productCode) {
    if (!productCodeSet[productCode]) {
      productCodeSet[productCode] = true;
      productCodesInOrder.push(productCode);
    }
  });

  if (selectionRows) {
    var selectionIndex = indexUniqueRows_(selectionRows, COL.SELECTION_SKU_CODE);
    duplicateCount += selectionIndex.duplicateCount;
    displayed.skuCodesInOrder.forEach(function (skuCode) {
      if (selectionIndex.duplicates[skuCode]) {
        return;
      }
      var selectionEntry = selectionIndex.map[skuCode];
      if (!selectionEntry) {
        return;
      }
      var productCode = normalizeString_(selectionEntry.values[COL.SELECTION_PRODUCT_CODE - 1]);
      if (productCode && !productCodeSet[productCode]) {
        productCodeSet[productCode] = true;
        productCodesInOrder.push(productCode);
      }
    });
  }

  return {
    productCodeSet: productCodeSet,
    productCodesInOrder: productCodesInOrder,
    duplicateCount: duplicateCount
  };
}

function indexUniqueRows_(sheetRows, keyColumn) {
  var map = {};
  var duplicates = {};
  var duplicateCount = 0;
  sheetRows.values.slice(1).forEach(function (row, index) {
    var key = normalizeString_(row[keyColumn - 1]);
    if (!key) {
      return;
    }
    if (map[key]) {
      if (!duplicates[key]) {
        duplicateCount++;
      }
      duplicates[key] = true;
      return;
    }
    map[key] = {
      rowNumber: index + 2,
      values: row
    };
  });
  return {
    map: map,
    duplicates: duplicates,
    duplicateCount: duplicateCount
  };
}

function getWorkSheetRows_(sheetName, expectedColumns) {
  var sheet = getSheetOrThrow_(sheetName);
  if (sheet.getLastRow() <= 1) {
    return [];
  }
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, expectedColumns).getValues().filter(function (row) {
    return normalizeString_(row[2]) !== '';
  });
}

function ensureCurrentSettingsComplete_() {
  var settings = getSettingsValues_();
  if (!settings.currentEvent) {
    throw new Error('イベント設定が未入力です。');
  }
  if (!settings.currentStartDate || !settings.currentEndDate) {
    throw new Error('現在の開始日 / 終了日が未入力です。');
  }
  return settings;
}

function ensureRestoreSettingsComplete_() {
  var settings = getSettingsValues_();
  if (!settings.restoreStartDate || !settings.restoreEndDate) {
    throw new Error('復旧設定が未入力です。');
  }
  return settings;
}

function getSettingsValues_() {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.SETTINGS);
  var values = sheet.getRange(2, 2, 13, 1).getValues().flat();
  return {
    currentEvent: normalizeString_(values[0]),
    currentStartDate: asDateOnly_(values[1]),
    currentEndDate: asDateOnly_(values[2]),
    currentStartTime: normalizeTimeCell_(values[3], APP_CONFIG.DEFAULTS.CURRENT_START_TIME),
    currentEndTime: normalizeTimeCell_(values[4], APP_CONFIG.DEFAULTS.CURRENT_END_TIME),
    restoreStartDate: asDateOnly_(values[5]),
    restoreEndDate: asDateOnly_(values[6]),
    restoreStartTime: normalizeTimeCell_(values[7], APP_CONFIG.DEFAULTS.RESTORE_START_TIME),
    restoreEndTime: normalizeTimeCell_(values[8], APP_CONFIG.DEFAULTS.RESTORE_END_TIME),
    flagPrefix: normalizeString_(values[9]) || APP_CONFIG.DEFAULTS.FLAG_PREFIX,
    doublePriceText: normalizeString_(values[10]) || APP_CONFIG.DEFAULTS.DOUBLE_PRICE_TEXT,
    productNameMaxLength: Number(values[11] || APP_CONFIG.DEFAULTS.PRODUCT_NAME_MAX_LENGTH),
    outputEncoding: normalizeString_(values[12]) || APP_CONFIG.DEFAULTS.OUTPUT_ENCODING
  };
}

function buildSingleName_(eventName, displayPrice, newPrice, originalName, maxLength) {
  var prefix = '[' + eventName + ' ' + displayPrice + '円→' + newPrice + '円]';
  var remain = Math.max(Number(maxLength || 127) - prefix.length, 0);
  return prefix + originalName.slice(0, remain);
}

function stripSalePrefix_(name) {
  return name.replace(/^\[[^\]]+ (?:\d+円→\d+円|\d+%OFF)\]/, '');
}

function calculateDiscountedPrice_(displayPrice, discount) {
  return Math.floor(Number(displayPrice) * (1 - Number(discount) / 100));
}

function buildFlagValue_(prefix) {
  return prefix + '_' + Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd');
}

function startsManagedFlag_(value, prefix) {
  return normalizeString_(value).indexOf(prefix + '_') === 0;
}

function formatDateTimeForCsv_(dateValue, timeValue) {
  var hm = formatTimeForDisplay_(timeValue).split(':');
  var merged = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate(), Number(hm[0]), Number(hm[1]), 0, 0);
  return Utilities.formatDate(merged, APP_CONFIG.TIMEZONE, 'yyyyMMddHHmm');
}

function buildSingleNameFormula_(row) {
  var prefix = buildPrefixFormula_(row, 'G', 'B');
  return '=IF(OR($B' + row + '="",NOT(ISNUMBER($B' + row + ')),$B' + row + '<>INT($B' + row + '),$B' + row + '<=0,$B' + row + '>=100),"",' +
    prefix + '&LEFT($E' + row + ',MAX(0,' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.PRODUCT_NAME_MAX_LENGTH) + '-LEN(' + prefix + '))))';
}

function buildSinglePriceFormula_(row) {
  return '=IF(OR($B' + row + '="",NOT(ISNUMBER($B' + row + ')),$B' + row + '<>INT($B' + row + '),$B' + row + '<=0,$B' + row + '>=100),"",ROUNDDOWN($G' + row + '*(1-$B' + row + '/100),0))';
}

function buildVariationPriceFormula_(row) {
  return '=IF(OR($B' + row + '="",NOT(ISNUMBER($B' + row + ')),$B' + row + '<>INT($B' + row + '),$B' + row + '<=0,$B' + row + '>=100),"",ROUNDDOWN($G' + row + '*(1-$B' + row + '/100),0))';
}

function buildCurrentStartFormula_(row, priceColumn) {
  return '=IF($' + priceColumn + row + '="","",TEXT(' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.CURRENT_START_DATE) + '+' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.CURRENT_START_TIME) + ',"yyyy/mm/dd hh:mm"))';
}

function buildCurrentEndFormula_(row, priceColumn) {
  return '=IF($' + priceColumn + row + '="","",TEXT(' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.CURRENT_END_DATE) + '+' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.CURRENT_END_TIME) + ',"yyyy/mm/dd hh:mm"))';
}

function buildDoublePriceFormula_(row, priceColumn) {
  return '=IF($' + priceColumn + row + '="","",' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.DOUBLE_PRICE_TEXT) + ')';
}

function buildSingleStatusFormula_(row) {
  return '=IF($C' + row + '="","",IF($A' + row + '=FALSE,"未選択",IF($B' + row + '="","割引率未入力",IF(NOT(ISNUMBER($B' + row + ')),"割引率不正",IF($B' + row + '<>INT($B' + row + '),"整数で入力",IF(OR($B' + row + '<=0,$B' + row + '>=100),"割引率範囲外","反映予定"))))))';
}

function buildVariationStatusFormula_(row) {
  return '=IF($C' + row + '="","",IF($A' + row + '=FALSE,"未選択",IF($B' + row + '="","割引率未入力",IF(NOT(ISNUMBER($B' + row + ')),"割引率不正",IF($B' + row + '<>INT($B' + row + '),"整数で入力",IF(OR($B' + row + '<=0,$B' + row + '>=100),"割引率範囲外","反映予定"))))))';
}

function buildPrefixFormula_(row, priceColumn, discountColumn) {
  return '"["&' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.CURRENT_EVENT) + '&" "&TEXT($' + priceColumn + row + ',"0")&"円→"&TEXT(ROUNDDOWN($' + priceColumn + row + '*(1-$' + discountColumn + row + '/100),0),"0")&"円]"';
}

function settingsRef_(rowNumber) {
  return "'" + APP_CONFIG.SHEETS.SETTINGS + "'!$B$" + rowNumber;
}

function decodeCsvBytes_(bytes) {
  var utf8 = Utilities.newBlob(bytes).getDataAsString('UTF-8');
  var shiftJis = Utilities.newBlob(bytes).getDataAsString('Shift_JIS');
  return scoreDecodedText_(shiftJis) > scoreDecodedText_(utf8)
    ? { text: shiftJis, encoding: 'Shift_JIS' }
    : { text: utf8, encoding: 'UTF-8' };
}

function scoreDecodedText_(text) {
  if (!text) {
    return -10000;
  }
  return (text.match(/[ぁ-んァ-ヶ一-龯]/g) || []).length * 3
    - (text.match(/\uFFFD/g) || []).length * 50
    - (text.match(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g) || []).length * 10;
}

function padRows_(rows) {
  var width = rows.reduce(function (max, row) {
    return Math.max(max, row.length);
  }, 0);
  return rows.map(function (row) {
    var clone = row.slice();
    while (clone.length < width) {
      clone.push('');
    }
    return clone;
  });
}

function stripUtf8Bom_(text) {
  return text.replace(/^\uFEFF/, '');
}

function stripDataUrlPrefix_(base64) {
  return base64.replace(/^data:.*?;base64,/, '');
}

function normalizeString_(value) {
  return value == null ? '' : String(value).trim();
}

function truncateText_(text, maxLength) {
  var normalized = normalizeString_(text);
  return normalized.length <= maxLength ? normalized : normalized.slice(0, maxLength);
}

function toNumber_(value) {
  if (value === '' || value == null) {
    return null;
  }
  var numberValue = Number(String(value).replace(/,/g, '').trim());
  return isNaN(numberValue) ? null : numberValue;
}

function toSheetValue_(value) {
  var numeric = toNumber_(value);
  return numeric === null ? normalizeString_(value) : numeric;
}

function asBoolean_(value) {
  if (value === true || value === false) {
    return value;
  }
  return normalizeString_(value).toUpperCase() === 'TRUE';
}

function isValidDiscountInteger_(value) {
  if (value === '' || value == null) {
    return false;
  }
  var numeric = Number(value);
  return !isNaN(numeric) && numeric === Math.floor(numeric) && numeric > 0 && numeric < 100;
}

function parseDateInput_(value) {
  var text = normalizeString_(value);
  if (!text) {
    return null;
  }
  var parts = text.split('-');
  var parsed = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 0, 0, 0, 0);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function asDateOnly_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) {
    return null;
  }
  return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0, 0);
}

function normalizeTimeCell_(value, fallback) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  return makeTimeValue_(normalizeString_(value) || fallback);
}

function makeTimeValue_(hhmm) {
  var parts = normalizeString_(hhmm).split(':');
  return new Date(1899, 11, 30, Number(parts[0] || 0), Number(parts[1] || 0), 0, 0);
}

function formatDateForInput_(dateValue) {
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime())) {
    return '';
  }
  return Utilities.formatDate(dateValue, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function formatTimeForDisplay_(timeValue) {
  if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
    return Utilities.formatDate(timeValue, APP_CONFIG.TIMEZONE, 'HH:mm');
  }
  return normalizeString_(timeValue);
}

function columnToIndex_(columnLabel) {
  var index = 0;
  var text = String(columnLabel).toUpperCase();
  for (var i = 0; i < text.length; i++) {
    index = index * 26 + (text.charCodeAt(i) - 64);
  }
  return index;
}

function indexToColumn_(index) {
  var result = '';
  var current = Number(index);
  while (current > 0) {
    var remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

function generateSingleWorkSheetCore_() {
  var rmsRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_RMS, [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE], 'RMS価格表示結果CSV');
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var itemsubRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, [
    COL.ITEMSUB_PRODUCT_CODE,
    COL.ITEMSUB_NAME,
    COL.ITEMSUB_NORMAL_PRICE,
    COL.ITEMSUB_DISPLAY_PRICE,
    COL.ITEMSUB_DOUBLE_PRICE_TEXT
  ], 'IR itemsub CSV');

  var displayed = collectDisplayedRms_(rmsRows);
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var itemsubIndex = indexUniqueRows_(itemsubRows, COL.ITEMSUB_PRODUCT_CODE);
  var records = [];
  var errorCount = itemIndex.duplicateCount + itemsubIndex.duplicateCount;
  var skippedCount = 0;

  displayed.productCodesInOrder.forEach(function (productCode) {
    var itemEntry = itemIndex.map[productCode];
    if (!itemEntry || itemIndex.duplicates[productCode]) {
      skippedCount++;
      return;
    }
    if (normalizeString_(itemEntry.values[COL.ITEM_STOCK_TYPE - 1]) !== '1') {
      return;
    }
    var itemsubEntry = itemsubIndex.map[productCode];
    if (!itemsubEntry || itemsubIndex.duplicates[productCode]) {
      skippedCount++;
      return;
    }
    var displayPrice = toNumber_(itemsubEntry.values[COL.ITEMSUB_DISPLAY_PRICE - 1]);
    if (displayPrice === null) {
      errorCount++;
      skippedCount++;
      return;
    }
    records.push([
      false,
      '',
      productCode,
      itemsubEntry.rowNumber,
      normalizeString_(itemsubEntry.values[COL.ITEMSUB_NAME - 1]),
      '',
      displayPrice,
      toSheetValue_(itemsubEntry.values[COL.ITEMSUB_NORMAL_PRICE - 1]),
      '',
      '',
      '',
      '',
      ''
    ]);
  });

  writeSingleWorkSheet_(records);
  return {
    targetCount: records.length,
    updatedCount: records.length,
    restoredCount: 0,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: '単品_作業シートを生成しました。\n対象: ' + records.length + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}
