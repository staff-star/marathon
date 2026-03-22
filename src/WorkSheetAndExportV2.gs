function generateSingleWorkSheetV2Core_() {
  var result = generateSingleWorkSheetCore_();
  patchSingleWorkSheetPreviewV2_();
  return result;
}

function generateVariationWorkSheetV2Core_() {
  var result = generateVariationWorkSheetByProductCodeCore_();
  patchVariationWorkSheetPreviewV2_();
  return result;
}

function patchSingleWorkSheetPreviewV2_() {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.WORK_SINGLE);
  if (sheet.getLastRow() <= 1) {
    return;
  }
  var rowCount = sheet.getLastRow() - 1;
  var nameFormulas = [];
  var priceFormulas = [];
  for (var i = 0; i < rowCount; i++) {
    var row = i + 2;
    nameFormulas.push([buildSingleNameFormulaV2_(row)]);
    priceFormulas.push([buildSinglePriceFormulaV2_(row)]);
  }
  sheet.getRange(2, 6, rowCount, 1).setFormulas(nameFormulas);
  sheet.getRange(2, 9, rowCount, 1).setFormulas(priceFormulas);
}

function patchVariationWorkSheetPreviewV2_() {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.WORK_VARIATION);
  if (sheet.getLastRow() <= 1) {
    return;
  }
  var rowCount = sheet.getLastRow() - 1;
  var priceFormulas = [];
  var nameFormulas = [];
  sheet.getRange(1, 11).setValue('商品名（更新後）');
  styleHeaderRow_(sheet.getRange(1, 11, 1, 1));
  sheet.getRange(2, 11, rowCount, 1).setBackground(APP_CONFIG.COLORS.AUTO_BG);
  sheet.setColumnWidth(11, 320);
  for (var i = 0; i < rowCount; i++) {
    var row = i + 2;
    priceFormulas.push([buildVariationPriceFormulaV2_(row)]);
    nameFormulas.push([buildVariationNameFormulaV2_(row)]);
  }
  sheet.getRange(2, 9, rowCount, 1).setFormulas(priceFormulas);
  sheet.getRange(2, 11, rowCount, 1).setFormulas(nameFormulas);
}

function buildSingleNameFormulaV2_(row) {
  var prefix = buildPrefixFormulaV2_(row, 'G', 'B');
  return '=IF(OR($B' + row + '="",NOT(ISNUMBER($B' + row + ')),$B' + row + '<>INT($B' + row + '),$B' + row + '<=0,$B' + row + '>=100),"",' +
    prefix + '&LEFT($E' + row + ',MAX(0,' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.PRODUCT_NAME_MAX_LENGTH) + '-LEN(' + prefix + '))))';
}

function buildVariationNameFormulaV2_(row) {
  var prefix = buildPrefixFormulaV2_(row, 'G', 'B');
  return '=IF(OR($B' + row + '="",NOT(ISNUMBER($B' + row + ')),$B' + row + '<>INT($B' + row + '),$B' + row + '<=0,$B' + row + '>=100),"",' +
    prefix + '&LEFT($D' + row + ',MAX(0,' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.PRODUCT_NAME_MAX_LENGTH) + '-LEN(' + prefix + '))))';
}

function buildSinglePriceFormulaV2_(row) {
  return '=IF(OR($B' + row + '="",NOT(ISNUMBER($B' + row + ')),$B' + row + '<>INT($B' + row + '),$B' + row + '<=0,$B' + row + '>=100),"",ROUNDDOWN($G' + row + '*(1-$B' + row + '/100),-1))';
}

function buildVariationPriceFormulaV2_(row) {
  return '=IF(OR($B' + row + '="",NOT(ISNUMBER($B' + row + ')),$B' + row + '<>INT($B' + row + '),$B' + row + '<=0,$B' + row + '>=100),"",ROUNDDOWN($G' + row + '*(1-$B' + row + '/100),-1))';
}

function buildPrefixFormulaV2_(row, priceColumn, discountColumn) {
  return '"【"&' + settingsRef_(APP_CONFIG.SETTINGS_ROWS.CURRENT_EVENT) + '&" "&TEXT($' + priceColumn + row + ',"0")&"円→"&TEXT(ROUNDDOWN($' + priceColumn + row + '*(1-$' + discountColumn + row + '/100),-1),"0")&"円】"';
}

function buildSingleNameV2_(eventName, displayPrice, newPrice, originalName, maxLength) {
  var prefix = '【' + eventName + ' ' + displayPrice + '円→' + newPrice + '円】';
  var remain = Math.max(Number(maxLength || 127) - prefix.length, 0);
  return prefix + normalizeString_(originalName).slice(0, remain);
}

function calculateDiscountedPriceV2_(displayPrice, discount) {
  var raw = Number(displayPrice) * (1 - Number(discount) / 100);
  return Math.floor(raw / 10) * 10;
}

function stripSalePrefixV2_(name) {
  return normalizeString_(name).replace(/^[\[【][^\]】]+ (?:\d+円→\d+円|\d+%OFF)[\]】]/, '');
}

function applySingleUpdatesV2Core_() {
  var settings = ensureCurrentSettingsComplete_();
  var workRows = getWorkSheetRows_(APP_CONFIG.SHEETS.WORK_SINGLE, 13);
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
    var newPrice = calculateDiscountedPriceV2_(displayPrice, Number(row[1]));
    targetRow[COL.ITEMSUB_NAME - 1] = buildSingleNameV2_(
      settings.currentEvent,
      displayPrice,
      newPrice,
      normalizeString_(targetRow[COL.ITEMSUB_NAME - 1]),
      settings.productNameMaxLength
    );
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

function applyVariationUpdatesV2Core_() {
  getSettingsValues_();
  var workRows = getWorkSheetRows_(APP_CONFIG.SHEETS.WORK_VARIATION, APP_CONFIG.VARIATION_HEADERS.length);
  var selectionRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [
    COL.SELECTION_PRODUCT_CODE,
    COL.SELECTION_NAME,
    COL.SELECTION_SKU_CODE,
    COL.SELECTION_NORMAL_PRICE,
    COL.SELECTION_DISPLAY_PRICE
  ], 'IR selection CSV');
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
    targetRow[COL.SELECTION_NORMAL_PRICE - 1] = calculateDiscountedPriceV2_(displayPrice, Number(row[1]));
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

function restoreSingleProductsV2Core_() {
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
    row[COL.ITEMSUB_NAME - 1] = stripSalePrefixV2_(normalizeString_(row[COL.ITEMSUB_NAME - 1]));
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

function showFilteredExportDialog_(kind) {
  ensureBaseSheets_();
  var payload = buildFilteredExportPayload_(kind);
  var template = HtmlService.createTemplateFromFile('DownloadDialog');
  template.fileName = payload.fileName;
  template.base64 = payload.base64;
  template.description = payload.description;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(230),
    'CSVダウンロード'
  );
}

function buildFilteredExportPayload_(kind) {
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
  var filteredValues = filterExportValuesByWorkTargets_(kind, values);
  if (filteredValues.length <= 1) {
    throw new Error('出力対象がありません。作業シートで「反映する」をオンにした商品を確認してください。');
  }

  var fileName = config.prefix + '_作業分_' + Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss') + '.csv';
  var blob = Utilities.newBlob(buildCsvText_(filteredValues), 'Shift_JIS', fileName);
  return {
    fileName: fileName,
    base64: Utilities.base64Encode(blob.getBytes()),
    description: config.sheetName + ' のうち、作業シートで選択した商品だけを Shift_JIS / CRLF で出力します。'
  };
}

function filterExportValuesByWorkTargets_(kind, values) {
  var targets = collectSelectedWorkTargetsForExport_();
  if (kind === 'item') {
    return filterExportRowsByProductCodes_(values, COL.ITEM_PRODUCT_CODE, targets.itemProductCodes);
  }
  if (kind === 'itemsub') {
    return filterExportRowsBySourceRowNumbers_(values, targets.singleSourceRowNumbers);
  }
  if (kind === 'selection') {
    return filterExportRowsBySourceRowNumbers_(values, targets.variationSourceRowNumbers);
  }
  throw new Error('不明な出力種別です。');
}

function collectSelectedWorkTargetsForExport_() {
  var singleTargets = getSelectedSingleWorkTargetsV2_();
  var variationTargets = getSelectedVariationWorkTargetsV2_();
  var itemProductCodes = {};

  singleTargets.productCodes.forEach(function (productCode) {
    itemProductCodes[productCode] = true;
  });
  variationTargets.productCodes.forEach(function (productCode) {
    itemProductCodes[productCode] = true;
  });

  return {
    itemProductCodes: itemProductCodes,
    singleSourceRowNumbers: singleTargets.sourceRowNumbers,
    variationSourceRowNumbers: variationTargets.sourceRowNumbers
  };
}

function getSelectedSingleWorkTargetsV2_() {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.WORK_SINGLE);
  var productCodes = [];
  var productCodeSet = {};
  var sourceRowNumbers = {};
  if (sheet.getLastRow() <= 1) {
    return { productCodes: productCodes, sourceRowNumbers: sourceRowNumbers };
  }
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, APP_CONFIG.SINGLE_HEADERS.length).getValues();
  values.forEach(function (row) {
    var productCode = normalizeString_(row[2]);
    var sourceRowNumber = Number(row[3]);
    if (row[0] !== true || !isValidDiscountInteger_(row[1]) || !productCode || !sourceRowNumber) {
      return;
    }
    if (!productCodeSet[productCode]) {
      productCodeSet[productCode] = true;
      productCodes.push(productCode);
    }
    sourceRowNumbers[sourceRowNumber] = true;
  });
  return { productCodes: productCodes, sourceRowNumbers: sourceRowNumbers };
}

function getSelectedVariationWorkTargetsV2_() {
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.WORK_VARIATION);
  var productCodes = [];
  var productCodeSet = {};
  var sourceRowNumbers = {};
  if (sheet.getLastRow() <= 1) {
    return { productCodes: productCodes, sourceRowNumbers: sourceRowNumbers };
  }
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, APP_CONFIG.VARIATION_HEADERS.length).getValues();
  values.forEach(function (row) {
    var productCode = normalizeString_(row[2]);
    var sourceRowNumber = Number(row[5]);
    if (row[0] !== true || !isValidDiscountInteger_(row[1]) || !productCode || !sourceRowNumber) {
      return;
    }
    if (!productCodeSet[productCode]) {
      productCodeSet[productCode] = true;
      productCodes.push(productCode);
    }
    sourceRowNumbers[sourceRowNumber] = true;
  });
  return { productCodes: productCodes, sourceRowNumbers: sourceRowNumbers };
}

function filterExportRowsByProductCodes_(values, productCodeColumn, productCodeSet) {
  var filtered = [values[0]];
  if (!Object.keys(productCodeSet).length) {
    return filtered;
  }
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var productCode = normalizeString_(row[productCodeColumn - 1]);
    if (productCodeSet[productCode]) {
      filtered.push(row);
    }
  }
  return filtered;
}

function filterExportRowsBySourceRowNumbers_(values, sourceRowNumbers) {
  var filtered = [values[0]];
  if (!Object.keys(sourceRowNumbers).length) {
    return filtered;
  }
  for (var i = 1; i < values.length; i++) {
    if (sourceRowNumbers[i + 1]) {
      filtered.push(values[i]);
    }
  }
  return filtered;
}
