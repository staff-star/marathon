function onOpen(e) {
  createMenu_();
}

function onInstall(e) {
  onOpen(e);
}

function createInitialSheets() {
  return runMenuAction_('シートを初期作成する', function () {
    ensureBaseSheets_({ rebuildIntro: true });
    return { message: '初期シートを作成しました。' };
  });
}

function recreateIntroSheet() {
  return runMenuAction_('説明書を再作成する', function () {
    ensureBaseSheets_({ rebuildIntro: true });
    return { message: 'はじめにシートを再作成しました。' };
  });
}

function openEventSettings() {
  ensureBaseSheets_();
  var template = HtmlService.createTemplateFromFile('EventSettingsDialog');
  template.payload = getEventDialogPayload_();
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(460).setHeight(420),
    'イベント設定'
  );
}

function openRestoreSettings() {
  ensureBaseSheets_();
  var template = HtmlService.createTemplateFromFile('RestoreSettingsDialog');
  template.payload = getRestoreDialogPayload_();
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(360),
    '復旧設定'
  );
}

function openBulkImportDialog() {
  ensureBaseSheets_();
  var template = HtmlService.createTemplateFromFile('BulkImportDialog');
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(520).setHeight(420),
    'CSV一括取込'
  );
}

function saveEventSettingsFromDialog(payload) {
  ensureBaseSheets_();
  var result = saveEventSettings_(payload);
  refreshOutputSummary_();
  appendProcessLog_(buildLogEntry_('イベント設定を開く', result));
  return result;
}

function saveRestoreSettingsFromDialog(payload) {
  ensureBaseSheets_();
  var result = saveRestoreSettings_(payload);
  refreshOutputSummary_();
  appendProcessLog_(buildLogEntry_('復旧設定を開く', result));
  return result;
}

function openImportRmsDialog() {
  showImportDialog_('rms');
}

function openImportItemDialog() {
  showImportDialog_('item');
}

function openImportItemsubDialog() {
  showImportDialog_('itemsub');
}

function openImportSelectionDialog() {
  showImportDialog_('selection');
}

function importCsvFileFromDialog(payload) {
  ensureBaseSheets_();
  var result = importCsvPayload_(payload);
  refreshOutputSummary_();
  appendProcessLog_(buildLogEntry_(getImportConfig_(payload.target).menuLabel, result));
  return result;
}

function importMultipleCsvFilesFromDialog(payloads) {
  ensureBaseSheets_();
  var results = importMultipleCsvPayloads_(payloads);
  refreshOutputSummary_();
  appendProcessLog_(buildLogEntry_('4つのCSVをまとめて取り込む', results));
  return results;
}

function generateSingleWorkSheet() {
  return runMenuAction_('単品の作業シートを作成する', function () {
    return generateSingleWorkSheetCore_();
  });
}

function generateVariationWorkSheet() {
  return runMenuAction_('バリエーションの作業シートを作成する', function () {
    return generateVariationWorkSheetCore_();
  });
}

function updateItemFlags() {
  return runMenuAction_('IRアイテムの付箋を更新する', function () {
    return updateItemFlagsCore_();
  });
}

function applySingleUpdates() {
  return runMenuAction_('単品商品の内容を反映する', function () {
    return applySingleUpdatesCore_();
  });
}

function applyVariationUpdates() {
  return runMenuAction_('バリエーション商品の内容を反映する', function () {
    return applyVariationUpdatesCore_();
  });
}

function runAllUpdates() {
  return runMenuAction_('すべての更新を実行する', function () {
    var flagResult = updateItemFlagsCore_();
    var singleResult = applySingleUpdatesCore_();
    var variationResult = applyVariationUpdatesCore_();
    return {
      targetCount: flagResult.targetCount + singleResult.targetCount + variationResult.targetCount,
      updatedCount: flagResult.updatedCount + singleResult.updatedCount + variationResult.updatedCount,
      restoredCount: 0,
      skippedCount: flagResult.skippedCount + singleResult.skippedCount + variationResult.skippedCount,
      errorCount: flagResult.errorCount + singleResult.errorCount + variationResult.errorCount,
      message:
        'すべての更新を実行しました。\n' +
        '付箋: ' + flagResult.updatedCount + '件\n' +
        '単品反映: ' + singleResult.updatedCount + '件\n' +
        'バリエーション反映: ' + variationResult.updatedCount + '件'
    };
  });
}

function restoreSingleProducts() {
  return runMenuAction_('単品を復旧する', function () {
    return restoreSingleProductsCore_();
  });
}

function restoreVariationProducts() {
  return runMenuAction_('バリエーションを復旧する', function () {
    return restoreVariationProductsCore_();
  });
}

function restoreAllProducts() {
  return runMenuAction_('すべてを復旧する', function () {
    var singleResult = restoreSingleProductsCore_();
    var variationResult = restoreVariationProductsCore_();
    return {
      targetCount: singleResult.targetCount + variationResult.targetCount,
      updatedCount: 0,
      restoredCount: singleResult.restoredCount + variationResult.restoredCount,
      skippedCount: singleResult.skippedCount + variationResult.skippedCount,
      errorCount: singleResult.errorCount + variationResult.errorCount,
      message:
        'すべてを復旧しました。\n' +
        '単品復旧: ' + singleResult.restoredCount + '件\n' +
        'バリエーション復旧: ' + variationResult.restoredCount + '件'
    };
  });
}

function exportItemCsv() {
  showExportDialog_('item');
}

function exportItemsubCsv() {
  showExportDialog_('itemsub');
}

function exportSelectionCsv() {
  showExportDialog_('selection');
}
