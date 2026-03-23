function updateItemFlagsV2Core_() {
  var itemRows = getImportedValues_(APP_CONFIG.SHEETS.IMPORT_ITEM, [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG], 'IR item CSV');
  var targets = collectSelectedWorkTargetsForExport_();
  var productCodeSet = targets.itemProductCodes || {};
  var targetCount = Object.keys(productCodeSet).length;
  if (!targetCount) {
    throw new Error('作業シートで反映対象を選択してください。');
  }

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

    if (productCodeSet[productCode]) {
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
    targetCount: targetCount,
    updatedCount: updatedCount + clearedCount,
    restoredCount: 0,
    skippedCount: itemIndex.duplicateCount,
    errorCount: itemIndex.duplicateCount,
    message:
      'IR item の付箋を更新しました。\n' +
      '対象商品: ' + targetCount + '件\n' +
      '上書き: ' + updatedCount + '件\n' +
      'クリア: ' + clearedCount + '件\n' +
      '重複エラー: ' + itemIndex.duplicateCount + '件'
  };
}
