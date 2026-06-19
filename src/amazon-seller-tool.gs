function exportFilteredDataToTSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // シートの全データを取得
  var range = sheet.getDataRange();
  var data = range.getValues();

  var tsvLines = [];

  // --- 1. ヘッダー行の処理（改行除去） ---
  var originalHeader = data[0];
  var cleanHeader = originalHeader.map(function (cell) {
    // セルが文字列の場合、改行コード（\n や \r）をすべて半角スペース（または空文字）に置き換える
    if (typeof cell === 'string') {
      return cell.replace(/[\r\n]+/g, ' ').trim();
    }
    return cell;
  });
  tsvLines.push(cleanHeader.join('\t'));

  // --- 2. データ行の処理（フィルターを考慮） ---
  // 2行目（インデックス1）からループ
  for (var i = 1; i < data.length; i++) {
    // ★ここがポイント：現在行がフィルター等で非表示（非表示の行＝isRowHiddenByFilterがtrue）ならスキップ
    if (sheet.isRowHiddenByFilter(i + 1)) {
      continue;
    }

    var rowData = data[i];

    // データ内の改行やタブがTSVを崩さないように念のためクリーニング
    var cleanRowData = rowData.map(function (cell) {
      if (typeof cell === 'string') {
        // データ内の改line（もしあれば）もスペースに置換、タブもスペースに置換
        return cell.replace(/[\r\n\t]+/g, ' ').trim();
      }
      return cell;
    });

    // タブ区切りで結合して行を追加
    tsvLines.push(cleanRowData.join('\t'));
  }

  // 全行を改行で結合してTSVの中身を完成させる
  var tsvContent = tsvLines.join('\n');

  // --- 3. 自動ダウンロード処理 ---
  var blob = Utilities.newBlob(
    tsvContent,
    'text/tab-separated-values',
    'management_data_all.tsv',
  );
  var base64Data = Utilities.base64Encode(blob.getBytes());

  var htmlOutput = HtmlService.createHtmlOutput(
    '<script>window.onload = function() { ' +
      'var a = document.createElement("a"); ' +
      'a.href = "data:text/tab-separated-values;base64,' +
      base64Data +
      '"; ' +
      'a.download = "management_data_all.tsv"; ' +
      'document.body.appendChild(a); ' +
      'a.click(); ' +
      'google.script.host.close(); ' +
      '}</script>',
  )
    .setWidth(50)
    .setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'AI分析用TSVを出力中...');
}

function exportAmazonSellerTSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // 1. 提示いただいた最新テンプレートの1行目の項目名を完全再現（タブ区切り・末尾に改行）
  var tsvContent =
    'SKU\t出品情報アクション\t外部製品コードの種類\t外部製品ID\tマーチャント推奨ASIN\tUNSPSCコード\tナショナルストックナンバー\tこの商品はAmazon.co.jp限定商品ですか？\tオファーをスキップ\t商品の状態\t提供条件に関する注記\t商品タックスコード\t予約商品の販売開始日\t最大注文可能個数\tギフトメッセージ付きで提供可能です\tギフト包装は利用できますか\tメイン画像の場所\tその他の画像の場所\tその他の画像の場所\tその他の画像の場所\tその他の画像の場所\tその他の画像の場所\t除外する支払い方法の指定\t付属品\tバッテリー寿命パーセンテージ\t化粧品機能\t機能\t機能\t機能\t機能\t動作状態\tパッケージソースタイプ\tフルフィルメントチャネルコード (JP)\t在庫数 (JP)\t処理時間 (JP)\t再入荷日 (JP)\t常時在庫 (JP)\t商品の販売価格 JPY (Amazonで販売, JP)\t自動価格設定ルール (Amazonで販売, JP)\t販売者許可最低価格 (Amazonで販売, JP)\t販売者許可最大価格 (Amazonで販売, JP)\tセール価格 JPY (Amazonで販売, JP)\tセール開始日 (Amazonで販売, JP)\tセール終了日 (Amazonで販売, JP)\t発売日 (Amazonで販売, JP)\t販売停止日 (Amazonで販売, JP)\t商品の販売価格 JPY (Amazonビジネス（B2B）, JP)\t販売者許可最低価格 (Amazonビジネス（B2B）, JP)\t販売者許可最大価格 (Amazonビジネス（B2B）, JP)\t発売日 (Amazonビジネス（B2B）, JP)\t販売停止日 (Amazonビジネス（B2B）, JP)\tまとめ買い価格タイプ (Amazonビジネス（B2B）, JP)\t数量しきい値 (最小, Amazonビジネス（B2B）, JP)\tまとめ買い価格 (定価/割引率, Amazonビジネス（B2B）, JP)\t数量しきい値 (最小, Amazonビジネス（B2B）, JP)\tまとめ買い価格 (定価/割引率, Amazonビジネス（B2B）, JP)\t数量しきい値 (最小, Amazonビジネス（B2B）, JP)\tまとめ買い価格 (定価/割引率, Amazonビジネス（B2B）, JP)\t数量しきい値 (最小, Amazonビジネス（B2B）, JP)\tまとめ買い価格 (定価/割引率, Amazonビジネス（B2B）, JP)\t数量しきい値 (最小, Amazonビジネス（B2B）, JP)\tまとめ買い価格 (定価/割引率, Amazonビジネス（B2B）, JP)\t配送テンプレート (JP)\t商品の長さ\t商品本体サイズ：奥行きの単位\t商品本体サイズ：幅\t商品本体サイズ：幅の単位\t商品本体サイズ：高さ\t商品本体サイズ：高さの単位\t商品パッケージの長さ\t商品パッケージサイズ：奥行きの単位\t商品パッケージサイズ：幅\t商品パッケージサイズ：幅の単位\t商品パッケージサイズ：高さ\t商品パッケージサイズ：height_unit\t商品パッケージ重量\t商品パッケージ重量の単位\t原産国\t電池/バッテリーが必要な商品ですか？\tこの商品に電池/バッテリーは含まれていますか？\t電池の種類\tリストされていないバッテリーセルの構成\t電池の重量\t電池の重量の単位\t電池の規格・形状\t電池の数\t電池の規格・形状\t電池の数\t電池の規格・形状\t電池の数\t電池の規格・形状\t電池の数\t電池の規格・形状\t電池の数\tリチウム金属電池のセルの数\tリチウムイオン電池のセルの数\tリチウム電池の電力量(Watt hours/ワット時定格量)\tリチウム電池の電力量の単位\t電池はどのように収納されていますか？\tリチウムバッテリーの重量の単位\tリチウムバッテリーの重量の単位\t商品に適用される危険物規制の種類\t商品に適用される危険物規制の種類\t商品に適用される危険物規制の種類\t商品に適用される危険物規制の種類\t商品に適用される危険物規制の種類\tSDSもしくは商品のパッケージに表示されている危険ラベルのクラス(GHS)\tSDSもしくは商品のパッケージに表示されている危険ラベルのクラス(GHS)\tSDSもしくは商品のパッケージに表示されている危険ラベルのクラス(GHS)\tSDSもしくは商品のパッケージに表示されている危険ラベルのクラス(GHS)\tSDSもしくは商品のパッケージに表示されている危険ラベルのクラス(GHS)\t危険物の物性を示すIDの種類\t危険物の物性を示すID\t安全データシート(SDSまたはMSDS)のURL\t商品の重量\t商品の重量単位\t購入者年齢制限の対象商品ですか\tグローバル出荷\tGHS化学物質Hコード\tGHS化学物質Hコード\tGHS化学物質Hコード\tGHS化学物質Hコード\tGHS化学物質Hコード\n';

  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var status = data[i][9]; // J列（10番目）：ステータス
    var fbaDelivery = data[i][20]; // U列（21番目）：FBA納品配送方法

    if (status === '020.出品準備中' && fbaDelivery !== '') {
      var sku = data[i][1]; // B列：自社管理番号 ➔ 「SKU」
      var prodId = data[i][2]; // C列：バーコード番号 ➔ 「外部製品ID」
      var idType = data[i][3]; // D列：バーコード種別 ➔ 「外部製品コードの種類」
      var conditionText = data[i][5]; // F列：コンディション名 ➔ 「商品の状態」
      var note = data[i][7]; // H列：コンディション説明 ➔ 「提供条件に関する注記」
      var price = data[i][17]; // R列：販売価格 ➔ 「商品の販売価格 JPY...」

      // Amazonの標準形式に値を変換
      var idTypeClean =
        idType === 'JAN' || idType === 'UPC' ? idType.toLowerCase() : 'upc';
      var condClean =
        conditionText.indexOf('新品') !== -1 ? 'New' : 'UsedLikeNew'; // 中古は一律「ほぼ新品」として扱いますが適宜調整してください

      // 2. テンプレートの列順にキッチリ合わせてデータを配置。空の列は \t でスキップ。
      var row =
        sku +
        '\t' + // SKU (1列目)
        'Update' +
        '\t' + // 出品情報アクション (2列目)
        idTypeClean +
        '\t' + // 外部製品コードの種類 (3列目)
        prodId +
        '\t' + // 外部製品ID (4列目)
        '\t\t\t\t\t' + // 5〜9列目（空）
        condClean +
        '\t' + // 商品の状態 (10列目)
        note +
        '\t' + // 提供条件に関する注記 (11列目)
        '\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t' + // 12〜31列目（空）
        'AMAZON_JP' +
        '\t' + // フルフィルメントチャネルコード (32列目: FBA発送)
        '' +
        '\t' + // 在庫数 (33列目: FBAなので空)
        '\t\t\t' + // 34〜36列目（空）
        price +
        '\t' + // 商品の販売価格 JPY (37列目)
        '\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t'; // 38列目以降（空）

      tsvContent += row + '\n';
      count++;
    }
  }

  if (count === 0) {
    Browser.msgBox(
      '「020.出品準備中」かつ「FBA納品配送方法」が入力されている商品は見つかりませんでした。',
    );
    return;
  }

  // Shift_JISで保存してGoogleドライブ経由でダウンロード
  var blob = Utilities.newBlob(
    '',
    'text/tab-separated-values; charset=Shift_JIS',
    'amazon_upload.tsv',
  );
  blob.setDataFromString(tsvContent, 'Shift_JIS');

  var file = DriveApp.createFile(blob);
  var downloadUrl =
    'https://docs.google.com/uc?export=download&id=' + file.getId();

  var htmlOutput = HtmlService.createHtmlOutput(
    '<p>最新テンプレート形式（Shift-JIS）のTSVファイルを作成しました。</p>' +
      '<a href="' +
      downloadUrl +
      '" target="_blank" style="font-size:16px; font-weight:bold; color:#0066cc;">➔ ファイルをダウンロードする</a>' +
      '<script>window.onload = function() { location.href = "' +
      downloadUrl +
      '"; setTimeout(function() { google.script.host.close(); }, 3000); }</script>',
  )
    .setWidth(400)
    .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Amazon用TSV出力完了');
}
