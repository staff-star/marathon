function clearManagedFlagsForRestoredProductsV3_(itemValues, productCodeSet, flagPrefix) {
  var clearedCount = 0;
  if (!productCodeSet) {
    return clearedCount;
  }
  for (var i = 1; i < itemValues.length; i++) {
    var row = itemValues[i];
    var productCode = normalizeString_(row[COL.ITEM_PRODUCT_CODE - 1]);
    if (!productCodeSet[productCode]) {
      continue;
    }
    if (startsManagedFlag_(row[COL.ITEM_FLAG - 1], flagPrefix)) {
      row[COL.ITEM_FLAG - 1] = '';
      clearedCount++;
    }
  }
  return clearedCount;
}

function restoreSingleProductsV3Core_() {
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
  var itemValues = itemRows.values;
  var itemsubValues = itemsubRows.values;
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
  var restoredProductCodes = {};

  targetCodes.forEach(function (productCode) {
    var itemsubEntry = itemsubIndex.map[productCode];
    if (!itemsubEntry || itemsubIndex.duplicates[productCode]) {
      skippedCount++;
      errorCount++;
      return;
    }
    var row = itemsubValues[itemsubEntry.rowNumber - 1];
    row[COL.ITEMSUB_NORMAL_PRICE - 1] = row[COL.ITEMSUB_DISPLAY_PRICE - 1];
    row[COL.ITEMSUB_NAME - 1] = stripSalePrefixV2_(row[COL.ITEMSUB_NAME - 1]);
    row[COL.ITEMSUB_START_AT - 1] = restoreStart;
    row[COL.ITEMSUB_END_AT - 1] = restoreEnd;
    row[COL.ITEMSUB_DOUBLE_PRICE_TEXT - 1] = settings.doublePriceText;
    restoredProductCodes[productCode] = true;
    restoredCount++;
  });

  var clearedFlagCount = clearManagedFlagsForRestoredProductsV3_(itemValues, restoredProductCodes, settings.flagPrefix);
  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, itemsubValues);
  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, itemValues);
  return {
    targetCount: targetCodes.length,
    updatedCount: 0,
    restoredCount: restoredCount,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message:
      '単品を復旧しました。\n' +
      '復旧: ' + restoredCount + '件\n' +
      '付箋解除: ' + clearedFlagCount + '件\n' +
      'スキップ: ' + skippedCount + '件\n' +
      'エラー: ' + errorCount + '件'
  };
}

function restoreVariationProductsV3Core_() {
  var settings = ensureRestoreSettingsComplete_();
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var selectionRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, [
    COL.SELECTION_PRODUCT_CODE,
    COL.SELECTION_NAME,
    COL.SELECTION_SKU_CODE,
    COL.SELECTION_NORMAL_PRICE,
    COL.SELECTION_DISPLAY_PRICE
  ], 'IR selection CSV');
  var itemsubRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, [
    COL.ITEMSUB_NAME
  ], 'IR itemsub CSV');
  var itemsubColumns = getItemsubSalePeriodColumnsV2_(itemsubRows.header);
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var selectionIndex = indexUniqueRows_(selectionRows, COL.SELECTION_SKU_CODE);
  var itemsubRowsByProductCode = indexRowsByColumnValueV2_(itemsubRows, itemsubColumns.productCode);
  var itemValues = itemRows.values;
  var selectionValues = selectionRows.values;
  var itemsubValues = itemsubRows.values;
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

  var restoreStart = formatDateTimeForCsv_(settings.restoreStartDate, settings.restoreStartTime);
  var restoreEnd = formatDateTimeForCsv_(settings.restoreEndDate, settings.restoreEndTime);
  var restoredSelectionCount = 0;
  var restoredNameCount = 0;
  var restoredSalePeriodCount = 0;
  var skippedCount = itemIndex.duplicateCount;
  var errorCount = itemIndex.duplicateCount + selectionIndex.duplicateCount;
  var restoredSelectionProductCodes = {};
  var clearedFlagProductCodes = {};

  for (var i = 1; i < selectionValues.length; i++) {
    var row = selectionValues[i];
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
    restoredSelectionCount++;
    restoredSelectionProductCodes[productCode] = true;
  }

  Object.keys(targetProductCodes).forEach(function (productCode) {
    if (!restoredSelectionProductCodes[productCode]) {
      skippedCount++;
      errorCount++;
      return;
    }
    var itemsubEntries = itemsubRowsByProductCode[productCode];
    if (!itemsubEntries || !itemsubEntries.length) {
      skippedCount++;
      errorCount++;
      return;
    }
    itemsubEntries.forEach(function (entry) {
      var itemsubRow = itemsubValues[entry.rowNumber - 1];
      itemsubRow[COL.ITEMSUB_NAME - 1] = stripSalePrefixV2_(itemsubRow[COL.ITEMSUB_NAME - 1]);
      itemsubRow[itemsubColumns.startAt - 1] = restoreStart;
      itemsubRow[itemsubColumns.endAt - 1] = restoreEnd;
      restoredNameCount++;
      restoredSalePeriodCount++;
    });
    clearedFlagProductCodes[productCode] = true;
  });

  var clearedFlagCount = clearManagedFlagsForRestoredProductsV3_(itemValues, clearedFlagProductCodes, settings.flagPrefix);
  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_SELECTION, selectionValues);
  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEMSUB, itemsubValues);
  writeBackImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, itemValues);
  return {
    targetCount: Object.keys(targetProductCodes).length,
    updatedCount: 0,
    restoredCount: restoredSelectionCount + restoredNameCount,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message:
      'バリエーションを復旧しました。\n' +
      '価格復旧: ' + restoredSelectionCount + '件\n' +
      '親商品名復旧: ' + restoredNameCount + '件\n' +
      '販売期間復旧: ' + restoredSalePeriodCount + '件\n' +
      '付箋解除: ' + clearedFlagCount + '件\n' +
      'スキップ: ' + skippedCount + '件\n' +
      'エラー: ' + errorCount + '件'
  };
}
