/**
 * 検索ボタン押下時処理
 * 対象スプレッドシートを検索する。
 * @return {boolean} 検索処理を実行した場合はtrue、実行していない場合はfalse
 */
function searchSpreadSheet() {

  // 実行確認
  if (!confirmSearch()) {
    return false;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
  // 対象スプレッドシート一覧をクリア
  sheet.getRange(ROW_INDEX_SS_LIST_TOP, COL_INDEX_SS_URL, sheet.getLastRow(), COLS_SS_LIST).clearContent();

  var conditionDate = "";
  // 検索用時間を取得
  var searchTime = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(NAME_RANGE_SEARCH_TIME).getValue();
  if (Number.isInteger(searchTime)) {
    var date = Moment.moment();
    var modifiedDate = date.subtract(searchTime, 'hours').format("YYYY-MM-DDTHH:mm:ss+09:00");
    conditionDate = " and modifiedDate >= '" + modifiedDate + "'";
  }

  var searchParam = "mimeType = '" + MimeType.GOOGLE_SHEETS + "' and fullText contains '新規出品' and fullText contains '" + NEW_EXHIBIT_MARK + "'" + conditionDate;
  // ※「fullText」は全角記号「◎」「○」を検索できない

  // ドライブのスプレッドシートを検索
  var targetFiles = DriveApp.searchFiles(searchParam);
  var values = [];
  while (targetFiles.hasNext()) {
    var targetFile = targetFiles.next();
    var name = targetFile.getName();
    // ファイル名に特定の文字が含まれる場合
    if (name.indexOf(SS_NM_TARGET_MARK) !== -1) {
      values.push([targetFile.getUrl(), targetFile.getName()]);
    }
  }
  if (values.length > 0) {
    sheet.getRange(ROW_INDEX_SS_LIST_TOP, COL_INDEX_SS_URL, values.length, values[0].length).setValues(values);
  }

  return true;
}

/**
 * CSV出力ボタン押下時処理
 * CSVファイルを出力してダウンロードする。
 */
function downloadFile() {

  try {
    // dialog.html をもとにHTMLファイルを生成
    // evaluate() は dialog.html 内の GAS を実行するため（ <?= => の箇所）
    var html = HtmlService.createTemplateFromFile("dialog").evaluate().setWidth(1).setHeight(1);
    // 上記HTMLファイルをダイアログ出力
    SpreadsheetApp.getUi().showModalDialog(html, "ダウンロード中");
  } catch(e) {
    Browser.msgBox("エラーが発生しました。\\n" + e.stack);
  }
}

/**
 * 検索＆CSV出力ボタン押下時処理
 * 対象スプレッドシートを検索、CSVファイルを出力してダウンロードする。
 */
function searchAndDownloadFile() {

  if (searchSpreadSheet()) {
    downloadFile();
  }
}

/**
 * 検索処理の確認メッセージを表示する。
 * @return {boolean} 検索処理を実行する場合はtrue、実行しない場合はfalse
 */
function confirmSearch() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var listValues = sheet.getRange(ROW_INDEX_SS_LIST_TOP, COL_INDEX_SS_URL, sheet.getLastRow(), COLS_SS_LIST).getValues();

  var isUnfinished = listValues.some(function(listRow, index, array) {
    // 出品完了していないシートが存在する
    if (listRow[0].indexOf("http") === 0 && listRow[2] !== SS_LIST_FIN_MARK) {
      return true;
    }
    return false;
  });

  if (isUnfinished) {
    var ret = Browser.msgBox("出品完了していないシートがあります。\\n対象スプレッドシートがクリアされますが検索を実行しますか？", Browser.Buttons.YES_NO);
    if (ret !== "yes") {
      return false;
    }
  }

  return true;
}

/**
 * 商品情報をCSVデータとして取得する。JSから使用。
 * @return {string} CSVデータ
 */
function getCsvData() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var urlList = sheet.getRange(ROW_INDEX_SS_LIST_TOP, COL_INDEX_SS_URL, sheet.getLastRow(), 1).getValues();

  // CSVリスト
  var csvList = [CSV_HEADER.join(",")];
  for (var i = 0; i < urlList.length; i++) {
    // URLがhttpで始まらない場合はスキップ
    if (urlList[i][0].indexOf("http") !== 0) {
      continue;
    }
    var targetSpreadSheet = SpreadsheetApp.openByUrl(urlList[i][0]);
    targetSpreadSheet.getSheets().forEach(function(targetSheet) {
      if (targetSheet.getName() !== "カテゴリ一覧") {
        var tmpCsvList = getCsvList(targetSheet);
        if (tmpCsvList != null) {
          csvList = csvList.concat(tmpCsvList);
        }
      }
    });
  }
  // 配列を文字列に変換
  var csvString = csvList.join("\r\n");
  return csvString;
}

/**
 * CSVファイル名を取得する。JSから使用。
 * @return {string} CSVファイル名
 */
function getCsvFileName() {

  return "auctown_csv_" + Moment.moment().format("YYYYMMDDHHmmss") + ".csv";
}

/**
 * 1シートの商品情報をCSVデータとして取得する。
 * @param {Sheet} sheet スプレッドシートのシート
 * @return {array} CSVデータ
 */
function getCsvList(sheet) {

  var csvList = [];
  var sheetValues = sheet.getRange(ROW_INDEX_ITEM_HEADER, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  // CSVファイル作成ルールに従い、改行、半角カンマ、ダブルクォーテーション、シングルクォーテーションは含まれていない前提とする
  for (var i = 0; i < sheetValues.length; i++) {
    // 出力対象外の場合はスキップ
    if (getSheetValue(sheetValues, i, "新規出品") !== NEW_EXHIBIT_MARK) {
      continue;
    }
    var csv = CSV_HEADER.map(function(header) {
      // 固定値
      if (header in CSV_CONST_VAL) {
        return CSV_CONST_VAL[header];
      }
      // 可変値
      if (header === "カテゴリ") {
        return getSheetValue(sheetValues, i, "カテゴリ").replace(new RegExp(CATEGORY_SEPARATOR + ".*"), "");
      }
      if (header === "タイトル") {
        return getSheetValue(sheetValues, i, "商品名");
      }
      if (header === "説明") {
        return getExplanation(sheet, sheetValues, i);
      }
      if (header === "開始価格" || header === "即決価格") {
        return getSheetValue(sheetValues, i, "販売価格");
      }
      if (header.search(/画像\d{1,2}$/) === 0) {
        var itemCd = getItemCd(sheet, sheetValues, i);
        // 商品コード_画像番号.jpg
        var fileNm = itemCd + "_" + header.replace("画像", "") + ".jpg";
        return fileNm;
      }
      if (header === "送料負担") {
        if (getSheetValue(sheetValues, i, "送料負担種類").trim() === "") {
          return "出品者";
        }
        return "落札者";
      }
      if (header === "商品の状態") {
        return getItemCondition(sheetValues, i);
      }
      if (header === "返品の可否") {
        if (getSheetValue(sheetValues, i, "送料負担種類").trim() === "") {
          return "返品可";
        }
        return "返品不可";
      }
      if (header === "送料固定") {
        if (getSheetValue(sheetValues, i, "送料負担種類").trim() === "") {
          return "";
        }
        return "はい";
      }
      if (header === "配送方法1全国一律価格" || header === "北海道料金1" || header === "沖縄料金1" || header === "離島料金1") {
        var postageCostType = getSheetValue(sheetValues, i, "送料負担種類").trim();
        if (postageCostType === "") {
          return "";
        }
        return getPostage(postageCostType, header);
      }

      // 設定なし
      return "";
    });
    csvList.push(csv.join(","));
  }

  return csvList;
}

/**
 * スプレッドシートデータの指定行から指定した列名の値を取得する。
 * @param {array} sheetValues 商品スプレッドシートのデータ（二次元配列）
 * @param {number} rowIndex 行インデックス（配列の添え字）
 * @param {string} colName 列名
 * @return {string} 指定した列名の値。見つからない場合は空文字
 */
function getSheetValue(sheetValues, rowIndex, colName) {

  if (sheetValues.length > 1) {
    var colIndex = sheetValues[0].indexOf(colName);
    if (colIndex > -1) {
      return sheetValues[rowIndex][colIndex];
    }
  }
  return "";
}

/**
 * 商品コードを取得する。
 * @param {Sheet} sheet 商品スプレッドシート
 * @param {array} sheetValues 商品スプレッドシートのデータ（二次元配列）
 * @param {number} rowIndex 行インデックス（配列の添え字）
 * @return {string} 商品の状態
 */
function getItemCd(sheet, sheetValues, rowIndex) {

  // テンプレート
  var template = sheet.getRange(CELL_ITEM_TEMPLATE).getValue();
  var itemCd = template.replace("****", getSheetValue(sheetValues, rowIndex, "商品コード"));

  return itemCd;
}

/**
 * 商品の状態を取得する。
 * @param {array} sheetValues 商品スプレッドシートのデータ（二次元配列）
 * @param {number} rowIndex 行インデックス（配列の添え字）
 * @return {string} 商品の状態
 */
function getItemCondition(sheetValues, rowIndex) {

  var condition = "";
  switch (getSheetValue(sheetValues, rowIndex, "コンディション")) {
    case "SS":
      condition = "未使用";
      break;
    case "S":
      condition = "未使用に近い";
      break;
    case "A":
      condition = "目立った傷や汚れなし";
      break;
    case "B":
      condition = "やや傷や汚れあり";
      break;
    case "C":
      condition = "傷や汚れあり";
      break;
    case "D":
      condition = "全体的に状態が悪い";
      break;
  }

  return condition;
}

/**
 * 送料を取得する。
 * @param {string} postageCostType 送料負担種類
 * @param {string} header CSVのヘッダー名
 * @return {string} 送料
 */
function getPostage(postageCostType, header) {

  var postageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("送料表");
  var values = postageSheet.getDataRange().getDisplayValues();
  // 行
  var rowIndex;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === postageCostType) {
      rowIndex = i;
      break;
    }
  }
  // 列
  var colIndex;
  for (var i = 0; i < values[0].length; i++) {
    // ヘッダー名には場所が含まれる
    if (header.indexOf(values[0][i]) !== -1) {
      colIndex = i;
      break;
    }
  }

  return values[rowIndex][colIndex];
}

/**
 * 説明を取得する。
 * @param {array} sheetValues 商品スプレッドシートのデータ（二次元配列）
 * @param {number} rowIndex 行インデックス（配列の添え字）
 * @return {string} 説明
 */
function getExplanation(sheet, sheetValues, rowIndex) {

  var htmlSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  var row = htmlSheet.getRange("2:2").getValues()[0];
  // テンプレート
  var template = row[0].replace(/\r\n|\n/g, "");
  // 置換項目名
  var replaceItemList = row[1].split(/\r\n|\n/);
  // 寸法項目名
  var sizeItemList = row[2].split(/\r\n|\n/);

  var html = template;
  replaceItemList.forEach(function(replaceItem) {
    var replaceString = "";
    switch (replaceItem) {
      case "寸法":
        sizeItemList.forEach(function(sizeItem) {
          var sizeValue = getSheetValue(sheetValues, rowIndex, sizeItem);
          if (sizeValue !== "") {
            replaceString += sizeItem + "：" + getSheetValue(sheetValues, rowIndex, sizeItem) + " cm<br>";
          }
        });
        break;
      case "商品の状態":
        replaceString = getItemCondition(sheetValues, rowIndex);
        break;
      case "商品コード":
        replaceString = getItemCd(sheet, sheetValues, rowIndex);
        break;      
      default:
        replaceString = getSheetValue(sheetValues, rowIndex, replaceItem);
        break;
    }
    html = html.replace(new RegExp("{%" + replaceItem + "%}", 'g'), replaceString);
  });

  return html;
}