var VARIATION_PARENT_LINK_OPTION_KEY_COLUMN = columnToIndex_('G');

function generateVariationWorkSheetByProductCodeCore_() {
  var rmsRows = getImportedValues_(
    APP_CONFIG.SHEETS.IMPORT_RMS,
    [COL.RMS_PRODUCT_CODE, COL.RMS_SKU_CODE, COL.RMS_RESULT, COL.RMS_DISPLAY_PRICE, COL.RMS_SALE_PRICE],
    'RMS価格表示結果CSV'
  );
  var itemRows = getImportedValues_(
    APP_CONFIG.SHEETS.IMPORT_ITEM,
    [COL.ITEM_PRODUCT_CODE, COL.ITEM_STOCK_TYPE, COL.ITEM_FLAG],
    'IR item CSV'
  );
  var selectionRows = getImportedValues_(
    APP_CONFIG.SHEETS.IMPORT_SELECTION,
    [
      COL.SELECTION_PRODUCT_CODE,
      COL.SELECTION_NAME,
      VARIATION_PARENT_LINK_OPTION_KEY_COLUMN,
      COL.SELECTION_NORMAL_PRICE,
      COL.SELECTION_DISPLAY_PRICE
    ],
    'IR selection CSV'
  );

  var displayed = collectDisplayedRms_(rmsRows);
  var itemIndex = indexUniqueRows_(itemRows, COL.ITEM_PRODUCT_CODE);
  var records = [];
  var errorCount = itemIndex.duplicateCount;
  var skippedCount = 0;
  var values = selectionRows.values;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var productCode = normalizeString_(row[COL.SELECTION_PRODUCT_CODE - 1]);
    if (!productCode || !displayed.productCodeSet[productCode]) {
      continue;
    }

    var itemEntry = itemIndex.map[productCode];
    if (!itemEntry || itemIndex.duplicates[productCode]) {
      skippedCount++;
      continue;
    }
    if (normalizeString_(itemEntry.values[COL.ITEM_STOCK_TYPE - 1]) !== '2') {
      continue;
    }

    var displayPrice = toNumber_(row[COL.SELECTION_DISPLAY_PRICE - 1]);
    if (displayPrice === null) {
      skippedCount++;
      errorCount++;
      continue;
    }

    records.push([
      false,
      '',
      productCode,
      normalizeString_(row[COL.SELECTION_NAME - 1]),
      buildVariationSheetSku_(productCode, row[VARIATION_PARENT_LINK_OPTION_KEY_COLUMN - 1]),
      i + 1,
      displayPrice,
      toSheetValue_(row[COL.SELECTION_NORMAL_PRICE - 1]),
      '',
      ''
    ]);
  }

  writeVariationWorkSheet_(records);
  var sheet = getSheetOrThrow_(APP_CONFIG.SHEETS.WORK_VARIATION);
  sheet.getRange(1, 5).setValue('バリエーションSKU');
  styleHeaderRow_(sheet.getRange(1, 1, 1, APP_CONFIG.VARIATION_HEADERS.length));

  return {
    targetCount: records.length,
    updatedCount: records.length,
    restoredCount: 0,
    skippedCount: skippedCount,
    errorCount: errorCount,
    message: 'バリエーション_作業シートを作成しました。\n対象: ' + records.length + '件\nスキップ: ' + skippedCount + '件\nエラー: ' + errorCount + '件'
  };
}

function buildVariationSheetSku_(productCode, optionKey) {
  return normalizeString_(productCode) + normalizeString_(optionKey);
}
