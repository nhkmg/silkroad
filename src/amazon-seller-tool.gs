/**
 * 設定：商品データが入っているメインシートの名前
 */
var DATA_SHEET_NAME = '在庫・販売管理';

/**
 * 共通設定：スプレッドシートの列インデックス（0から始まる番号）
 */
var COL_INDEX = {
  sku: 1, // B列：自社管理番号（SKU）
  productId: 2, // C列：バーコード番号
  idType: 3, // D列：バーコード種別
  listingType: 4, // E列：出品タイプ
  asin: 5, // F列：ASIN
  brandName: 6, // G列：ブランド名
  makerName: 7, // H列：メーカー名
  partNumber: 8, // I列：型番
  pureProductName: 9, // I列：商品名
  condition: 10, // K列：コンディション名
  note: 12, // M列：コンディション説明
  status: 13, // N列：ステータス (例: '020.出品準備中')
  price: 14, // O列：販売価格
  fbaDelivery: 15, // P列：FBA納品配送方法
  channel: 16, // Q列：販路
  batteriesRequired: 17, // R列：電池の有無
  supplierDeclaredDgHzRegulation: 18, // S列：危険物規制の種類
};

/**
 * 🔘 ボタン1：Amazon相乗り出品用TSV出力（5行ヘッダー対応版）
 */
function exportAmazonMeTooCatalogTSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('AmazonTemplate_相乗り');
  if (!templateSheet) {
    Browser.msgBox('エラー: 「AmazonTemplate_相乗り」シートが見つかりません。');
    return;
  }

  var lastRow = templateSheet.getLastRow();
  var lastCol = templateSheet.getLastColumn();
  var templateData = templateSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var tsvLines = [];

  // 1. 英語ヘッダーの行を正確に特定（新規用ロジックを移植）
  var englishHeaderRowIdx = -1;
  var headerRowsCount = 0;

  for (var r = 0; r < lastRow; r++) {
    var rowStr = templateData[r].join('\t');
    // 英語キーが含まれる行をヘッダー行として特定
    if (
      rowStr.indexOf('contribution_sku') !== -1 ||
      rowStr.indexOf('::record_action') !== -1
    ) {
      englishHeaderRowIdx = r;
      headerRowsCount = r + 2;
      break;
    }
  }

  if (englishHeaderRowIdx === -1) {
    englishHeaderRowIdx = 5;
    headerRowsCount = 7;
  }

  var englishHeaders = templateData[englishHeaderRowIdx];
  // 2. テンプレートの全ヘッダー行をTSVにコピー
  for (var r = 0; r < headerRowsCount; r++) {
    var headerRow = [];
    for (var c = 0; c < lastCol; c++) {
      headerRow.push(
        templateData[r][c] !== undefined ? templateData[r][c] : '',
      );
    }
    tsvLines.push(headerRow.join('\t'));
  }

  // 3. データ抽出とマッピング
  var dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  var data = dataSheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    if (
      data[i][COL_INDEX.listingType] === '相乗り' &&
      data[i][COL_INDEX.status] === '020.出品準備中' &&
      data[i][COL_INDEX.fbaDelivery] !== ''
    ) {
      var sku = data[i][COL_INDEX.sku];
      var productId = data[i][COL_INDEX.productId];
      var idType = data[i][COL_INDEX.idType];
      var asin = data[i][COL_INDEX.asin];
      var price = data[i][COL_INDEX.price];
      var condition = data[i][COL_INDEX.condition];
      var note = data[i][COL_INDEX.note];
      var batteriesRequired = data[i][COL_INDEX.batteriesRequired];
      var supplierDeclaredDgHzRegulation =
        data[i][COL_INDEX.supplierDeclaredDgHzRegulation];
      var conditionType = getAmazonConditionType(data[i][COL_INDEX.condition]);

      var rowData = Array(lastCol).fill('');

      // 新規出品と同じマッピングロジックで全列走査
      for (var colIdx = 0; colIdx < lastCol; colIdx++) {
        var cellHeader = englishHeaders[colIdx];
        if (!cellHeader) continue;

        var headerKey = cellHeader.toString().trim().toLowerCase();

        // 💡 抽出データ・コア情報のマッピング
        if (headerKey.indexOf('contribution_sku') === 0) {
          rowData[colIdx] = sku;
          continue;
        } else if (headerKey.indexOf('::record_action') === 0) {
          rowData[colIdx] = '作成または編集';
          continue;
        } else if (
          headerKey.indexOf('externally_assigned_product_identifier') === 0
        ) {
          if (headerKey.endsWith('.type')) {
            rowData[colIdx] = idType;
          } else if (headerKey.endsWith('.value')) {
            rowData[colIdx] = productId;
          }
          continue;
        } else if (headerKey.indexOf('purchasable_offer') === 0) {
          if (
            headerKey.indexOf('audience=all') !== -1 &&
            headerKey.indexOf('our_price') !== -1 &&
            headerKey.indexOf('schedule') !== -1 &&
            headerKey.endsWith('value_with_tax')
          ) {
            // 商品の販売価格
            rowData[colIdx] = price;
          } else if (
            headerKey.indexOf('audience=all') !== -1 &&
            headerKey.indexOf('minimum_seller_allowed_price') !== -1 &&
            headerKey.indexOf('schedule') !== -1 &&
            headerKey.endsWith('value_with_tax')
          ) {
            // 販売者許可最低価格
            rowData[colIdx] = price;
          } else if (
            headerKey.indexOf('audience=all') !== -1 &&
            headerKey.indexOf('maximum_seller_allowed_price') !== -1 &&
            headerKey.indexOf('schedule') !== -1 &&
            headerKey.endsWith('value_with_tax')
          ) {
            // 販売者許可最大価格
            rowData[colIdx] = price * 1.2;
          }
          continue;
        } else if (headerKey.indexOf('batteries_required') === 0) {
          if (headerKey.endsWith('value')) {
            rowData[colIdx] = batteriesRequired;
          }
          continue;
        } else if (
          headerKey.indexOf('supplier_declared_dg_hz_regulation') === 0
        ) {
          if (headerKey.endsWith('value')) {
            rowData[colIdx] = supplierDeclaredDgHzRegulation;
          }
          continue;
        } else if (headerKey.indexOf('condition_type') === 0) {
          rowData[colIdx] = conditionType;
          continue;
        } else if (headerKey.indexOf('condition_note') === 0) {
          rowData[colIdx] = note
            ? note.toString().replace(/[\r\n\t]/g, ' ')
            : '';
          continue;
        } else if (
          headerKey.indexOf('fulfillment_availability') === 0 &&
          headerKey.indexOf('quantity') !== -1
        ) {
          rowData[colIdx] = '1';
          continue;
        } else if (
          headerKey.indexOf('fulfillment_availability') === 0 &&
          headerKey.indexOf('fulfillment_channel_code') !== -1
        ) {
          rowData[colIdx] = 'AMAZON_JP';
          continue;
        }
      }
      tsvLines.push(rowData.join('\t'));
      count++;
    }
  }

  if (count > 0) {
    executeOriginalStyleDownload(
      count,
      tsvLines.join('\r\n') + '\r\n',
      'amazon_me_too_catalog_upload.tsv',
    );
  } else {
    Browser.msgBox('対象データがありませんでした。');
  }
}

/**
 * 🔘 ボタン2：Amazon新規出品用TSV出力
 */
function exportAmazonNewCatalogTSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var templateSheet = ss.getSheetByName('AmazonTemplate_新規_おもちゃ');
  if (!templateSheet) {
    Browser.msgBox(
      'エラー: 「AmazonTemplate_新規_おもちゃ」タブが見つかりません。',
    );
    return;
  }

  var lastRow = templateSheet.getLastRow();
  var lastCol = templateSheet.getLastColumn();
  var templateData = templateSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var tsvLines = [];

  // 英語ヘッダーの行を正確に特定
  var englishHeaderRowIdx = -1;
  var headerRowsCount = 0;

  for (var r = 0; r < lastRow; r++) {
    var rowStr = templateData[r].join('\t');
    if (
      rowStr.indexOf('contribution_sku') !== -1 ||
      rowStr.indexOf('feed_product_type') !== -1
    ) {
      englishHeaderRowIdx = r;
      headerRowsCount = r + 2;
      break;
    }
  }

  if (englishHeaderRowIdx === -1) {
    englishHeaderRowIdx = 5;
    headerRowsCount = 7;
  }

  var englishHeaders = templateData[englishHeaderRowIdx];

  // テンプレートのヘッダー行をそのままTSVにコピー
  for (var r = 0; r < headerRowsCount; r++) {
    var headerRow = [];
    for (var c = 0; c < lastCol; c++) {
      var cellVal = templateData[r][c] !== undefined ? templateData[r][c] : '';
      headerRow.push(cellVal.toString().replace(/[\r\n\t]/g, ' '));
    }
    tsvLines.push(headerRow.join('\t'));
  }

  // 2. 「在庫・販売管理」シートからデータを抽出
  var sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    Browser.msgBox(
      'エラー: シート「' + DATA_SHEET_NAME + '」が見つかりません。',
    );
    return;
  }

  var data = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var listingType = data[i][COL_INDEX.listingType];
    var status = data[i][COL_INDEX.status];
    var fbaDelivery = data[i][COL_INDEX.fbaDelivery];

    if (
      listingType === '新規' &&
      status === '020.出品準備中' &&
      fbaDelivery !== ''
    ) {
      var sku = data[i][COL_INDEX.sku];
      var productId = data[i][COL_INDEX.productId];
      var idType = data[i][COL_INDEX.idType];
      var pureProductName = data[i][COL_INDEX.pureProductName];
      var brandName = data[i][COL_INDEX.brandName];
      var makerName = data[i][COL_INDEX.makerName];
      var price = data[i][COL_INDEX.price];
      var condition = data[i][COL_INDEX.condition];
      var partNumber = data[i][COL_INDEX.partNumber];

      var idTypeClean = 'JAN';
      if (idType) {
        var t = idType.toString().toUpperCase();
        if (t.indexOf('UPC') !== -1) idTypeClean = 'UPC';
        else if (t.indexOf('EAN') !== -1) idTypeClean = 'EAN';
      }

      var conditionType = getAmazonConditionType(data[i][COL_INDEX.condition]);

      var amazonFullProductName = makerName + ' ' + pureProductName;
      var rowData = Array(lastCol).fill('');

      // 各列のキーを正確に走査し、エラー項目に適切なデフォルト値や抽出データをマッピング
      for (var colIdx = 0; colIdx < lastCol; colIdx++) {
        var cellHeader = englishHeaders[colIdx];
        if (!cellHeader) continue;

        var headerKey = cellHeader.toString().trim().toLowerCase();

        // 💡 抽出データ・コア情報のマッピング
        if (headerKey.indexOf('contribution_sku') === 0) {
          rowData[colIdx] = sku;
          continue;
        } else if (headerKey.indexOf('item_name') === 0) {
          rowData[colIdx] = amazonFullProductName;
          continue;
        } else if (headerKey.indexOf('brand') === 0) {
          // ブランド名
          rowData[colIdx] = brandName;
          continue;
        } else if (headerKey.indexOf('manufacturer[') === 0) {
          // メーカー名
          rowData[colIdx] = makerName;
          continue;
        } else if (headerKey.indexOf('amzn1.volt.ca') === 0) {
          if (headerKey.indexOf('product_id_type') !== -1) {
            rowData[colIdx] = idTypeClean;
          } else if (headerKey.indexOf('product_id_value') !== -1) {
            rowData[colIdx] = productId;
          }
          continue;
        } else if (headerKey.indexOf('standard_price') === 0) {
          rowData[colIdx] = price;
          continue;
        } else if (headerKey.indexOf('condition_type') === 0) {
          rowData[colIdx] = conditionType;
          continue;
        } else if (headerKey.indexOf('product_type') === 0) {
          rowData[colIdx] = 'TOYS_AND_GAMES';
          continue;
        } else if (headerKey.indexOf('target_audience_keyword') === 0) {
          rowData[colIdx] = '男女両用';
          continue;
        } else if (headerKey.indexOf('product_description') === 0) {
          rowData[colIdx] = amazonFullProductName;
          continue;
        } else if (headerKey.indexOf('bullet_point') === 0) {
          rowData[colIdx] =
            'この商品は人気モデルです。詳細な仕様はメーカー公式サイトをご確認ください。';
          continue;
        } else if (headerKey.indexOf('material') === 0) {
          rowData[colIdx] = 'プラスチック';
          continue;
        } else if (headerKey.indexOf('model_number') === 0) {
          // 品番・型番
          rowData[colIdx] = partNumber;
          continue;
        } else if (headerKey.indexOf('part_number') === 0) {
          // メーカー型番
          rowData[colIdx] = partNumber;
          continue;
        } else if (headerKey.indexOf('distribution_designation') === 0) {
          rowData[colIdx] = '正規品';
          continue;
        } else if (headerKey.indexOf('is_exclusive_product') === 0) {
          rowData[colIdx] = 'いいえ';
          continue;
        } else if (headerKey.indexOf('manufacturer_minimum_age') === 0) {
          rowData[colIdx] = '3';
          continue;
        } else if (headerKey.indexOf('manufacturer_maximum_age') === 0) {
          rowData[colIdx] = '99';
          continue;
        } else if (headerKey.indexOf('included_components') === 0) {
          rowData[colIdx] = '本体、付属品一式';
          continue;
        } else if (headerKey.indexOf('list_price') === 0) {
          rowData[colIdx] = price;
          continue;
        } else if (headerKey.indexOf('fulfillment_availability') === 0) {
          if (headerKey.indexOf('fulfillment_channel_code') !== -1) {
            // フルフィルメントチャネルコード
            rowData[colIdx] = 'AMAZON_JP';
          } else if (headerKey.indexOf('quantity') !== -1) {
            // 在庫数
            rowData[colIdx] = '0';
          } else if (headerKey.indexOf('is_inventory_available') !== -1) {
            // 常時在庫
            rowData[colIdx] = '無';
          }
          continue;
        } else if (headerKey.indexOf('is_assembly_required') === 0) {
          rowData[colIdx] = 'いいえ';
          continue;
        } else if (headerKey.indexOf('item_dimensions') === 0) {
          if (headerKey.endsWith('#1.width.value') === 0) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.width.unit')) {
            rowData[colIdx] = 'ミリメートル';
          } else if (headerKey.endsWith('#1.height.value')) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.height.unit')) {
            rowData[colIdx] = 'ミリメートル';
          } else if (headerKey.endsWith('#1.length.value')) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.length.unit')) {
            rowData[colIdx] = 'ミリメートル';
          } else if (headerKey.endsWith('#1.width.value')) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.width.unit')) {
            rowData[colIdx] = 'ミリメートル';
          }
          continue;
        } else if (headerKey.indexOf('item_package_dimensions') === 0) {
          if (headerKey.endsWith('#1.width.value') === 0) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.width.unit')) {
            rowData[colIdx] = 'ミリメートル';
          } else if (headerKey.endsWith('#1.height.value')) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.height.unit')) {
            rowData[colIdx] = 'ミリメートル';
          } else if (headerKey.endsWith('#1.length.value')) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.length.unit')) {
            rowData[colIdx] = 'ミリメートル';
          } else if (headerKey.endsWith('#1.width.value')) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.width.unit')) {
            rowData[colIdx] = 'ミリメートル';
          }
          continue;
        } else if (headerKey.indexOf('item_package_weight') === 0) {
          if (headerKey.endsWith('#1.value')) {
            rowData[colIdx] = '50';
          } else if (headerKey.endsWith('#1.unit')) {
            rowData[colIdx] = 'グラム';
          }
          continue;
        } else if (headerKey.indexOf('safety_warning') === 0) {
          if (headerKey.endsWith('#1.value')) {
            rowData[colIdx] = '小さな部品を含みます';
          } else if (headerKey.endsWith('#2.value')) {
            rowData[colIdx] = '窒息の危険があります';
          }
          continue;
        } else if (headerKey.indexOf('number_of_boxes') === 0) {
          rowData[colIdx] = '1';
          continue;
        } else if (headerKey.indexOf('country_of_origin') === 0) {
          rowData[colIdx] = '日本';
          continue;
        } else if (headerKey.indexOf('batteries_required') === 0) {
          rowData[colIdx] = 'いいえ';
          continue;
        } else if (
          headerKey.indexOf('supplier_declared_dg_hz_regulation') === 0
        ) {
          rowData[colIdx] = '該当なし';
          continue;
        } else if (
          headerKey.indexOf('quantity') === 0 &&
          headerKey.indexOf('limit') === -1 &&
          headerKey.indexOf('max') === -1 &&
          headerKey.indexOf('minimum') === -1
        ) {
          rowData[colIdx] = '0';
          continue;
        }
      }

      tsvLines.push(rowData.join('\t'));
      count++;
    }
  }

  if (count === 0) {
    Browser.msgBox(
      '対象データが0件です。「' +
        DATA_SHEET_NAME +
        '」シートの条件を確認してください。',
    );
    return;
  }

  var finalTsvContent = tsvLines.join('\r\n') + '\r\n';
  executeOriginalStyleDownload(
    count,
    finalTsvContent,
    'amazon_new_catalog_upload.tsv',
  );
}

/**
 * 自動ダウンロード用HTML書き出し処理
 */
function executeOriginalStyleDownload(count, tsvContent, fileName) {
  var bom = '\uFEFF';
  var blob = Utilities.newBlob(
    '',
    'text/tab-separated-values; charset=utf-8',
    fileName,
  );
  blob.setDataFromString(bom + tsvContent, 'utf-8');

  var file = DriveApp.createFile(blob);
  var downloadUrl =
    'https://docs.google.com/uc?export=download&id=' + file.getId();

  var htmlString =
    '<p>TSVファイルを作成しました。自動保存されます。</p>' +
    '<a href="' +
    downloadUrl +
    '" target="_blank" style="font-size:16px; font-weight:bold; color:#0066cc;">➔ 保存されない場合はここをクリック</a>' +
    '<script>' +
    '  window.onload = function() { ' +
    '    location.href = "' +
    downloadUrl +
    '"; ' +
    '    setTimeout(function() { google.script.host.close(); }, 3000); ' +
    '  }' +
    '</script>';

  var htmlOutput = HtmlService.createHtmlOutput(htmlString)
    .setWidth(400)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Amazon用TSV出力完了');
}

/**
 * Amazonの仕様に基づき、スプレッドシートの入力値を正式なコンディション値に変換する
 */
function getAmazonConditionType(conditionText) {
  if (!conditionText) return '新品'; // デフォルト
  var c = conditionText.toString();

  // 中古系
  if (c.indexOf('中古') !== -1) {
    if (c.indexOf('ほぼ新品') !== -1) return '中古-ほぼ新品';
    if (c.indexOf('非常に良い') !== -1) return '中古-非常に良い';
    if (c.indexOf('良い') !== -1) return '中古-良い';
    if (c.indexOf('許容可能') !== -1) return '中古-許容可能';
    return '中古-許容可能'; // マッチしなければデフォルトで許容可能
  }

  // コレクター商品系
  if (c.indexOf('コレクター') !== -1) {
    if (c.indexOf('新品同様') !== -1) return 'コレクター商品-新品同様';
    if (c.indexOf('非常に良い') !== 'コレクター商品-非常に良い')
      return 'コレクター商品-非常に良い';
    if (c.indexOf('良い') !== -1) return 'コレクター商品-良い';
    if (c.indexOf('許容可能') !== -1) return 'コレクター商品-許容可能';
  }

  // 特殊系
  if (c.indexOf('開封済み未使用') !== -1) return '新品 - 開封済み未使用';
  if (c.indexOf('新品-OEM') !== -1) return '新品-OEM';
  if (c.indexOf('再生品') !== -1) return '再生品';
  if (c.indexOf('クラブ') !== -1) return 'クラブ';

  return '新品'; // どれにも該当しない場合は「新品」
}
