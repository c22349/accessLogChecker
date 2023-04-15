const folderId = "GoogleドライブのフォルダIDを記入"

// 最新のCSVファイルを取得
function findLatestCSV(files) {
  let latestDate = 0;
  let latestFile = null;

  while (files.hasNext()) {
    let file = files.next();
    let fileDate = file.getDateCreated();

    if (file.getMimeType() === MimeType.CSV && new Date(fileDate).getTime() > latestDate) {
      latestDate = new Date(fileDate).getTime();
      latestFile = file;
    }
  }

  return latestFile;
}

// バリデーションの情報取得
function getValidationCriteria(validationSheet) {
  const headerRow = validationSheet.getRange(1, 1, 1, validationSheet.getLastColumn()).getValues()[0];

  const getColumnValues = (columnName) => {
    const columnIndex = headerRow.findIndex(cell => cell === columnName) + 1;

    if (columnIndex === 0) {
      throw new Error(`${columnName}列が見つかりませんでした。`);
    }

    const criteriaRange = validationSheet.getRange(2, columnIndex, 35);
    const criteria = criteriaRange.getValues().flat();
    return criteria.filter(Boolean); // 空のセルを除外
  };

  const actionCriteria = getColumnValues("Action");
  const computerNameCriteria = getColumnValues("Computer name");
  const usersCriteria = getColumnValues("Users");

  return {
    actionCriteria,
    excludeComputerNames: computerNameCriteria,
    excludeUsers: usersCriteria
  };
}

// CSVファイルを読み込む
function importCSVtoSheet(latestFile) {
  let csv;

  try {
    const data = latestFile.getBlob().getDataAsString();
    csv = Utilities.parseCsv(data);
  } catch (error) {
    Browser.msgBox("CSVファイルの読み込みに失敗しました。")
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const validationSheet = ss.getSheetByName('validation');
  const validationSheetIndex = validationSheet.getIndex();

  // CSVファイルのヘッダ行を取得
  const headerRow = csv[0];

  const newSheetName = getUniqueSheetName();
  const newSheet = ss.insertSheet(newSheetName, validationSheetIndex);

  const criteria = getValidationCriteria(validationSheet);
  const actionCriteria = criteria.actionCriteria;
  const excludeComputerNames = criteria.excludeComputerNames;
  const excludeUsers = criteria.excludeUsers;

  // 条件に一致する行をフィルタリング
    const filteredCSV = csv.filter((row, index) => {
      // ヘッダ行をスキップ
      if (index === 0) return false;

      const action = row[headerRow.findIndex(cell => cell === "Action")];
      const computerName = row[headerRow.findIndex(cell => cell === "Computer name")];
      const user = row[headerRow.findIndex(cell => cell === "Users")];

      // バリデーション情報で選別
      return actionCriteria.includes(action) &&
            !excludeComputerNames.includes(computerName) &&
            !excludeUsers.includes(user);
    });

  // ヘッダ行を新しいシートに追加
  newSheet.appendRow(headerRow);

  // フィルターをかける前の行数
  const totalRowsBeforeFilter = csv.length - 3; // ヘッダーと追加行を除外

  for (let i = 0; i < filteredCSV.length; i++) {
    newSheet.appendRow(filteredCSV[i]);
  }

  const totalRowsAfterFilter = filteredCSV.length;

  // 4行追加し、各結果を表示する
  newSheet.insertRows(1, 4);
  newSheet.getRange("A1").setValue("CSVファイル: " + latestFile.getName());
  newSheet.getRange("A2").setValue("確認したログ数: " + totalRowsBeforeFilter);

  if (totalRowsAfterFilter === 0) {
    newSheet.getRange("A3").setValue("チェック結果: 問題ありません。");
  } else {
    newSheet.getRange("A3").setValue("チェック結果: 不正なアクセスの可能性があるログは " + totalRowsAfterFilter + "件です。");
  }

  Logger.log("取り込んだCSVファイル: " + latestFile.getName());
  Logger.log("確認したログ数: " + totalRowsBeforeFilter);
  Logger.log("該当ログ数: " + totalRowsAfterFilter);
}

// 新規シートの名前
function getUniqueSheetName() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const baseName = Utilities.formatDate(today, 'Asia/Tokyo', "yyMMdd");
  let sheetName = baseName;
  let index = 2;

  while (spreadsheet.getSheetByName(sheetName)) {
    sheetName = baseName + '-' + index;
    index++;
  }
  return sheetName;
}

// メインの処理
function main() {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  const latestFile = findLatestCSV(files);

  if (latestFile) {
    importCSVtoSheet(latestFile);
  } else {
    Logger.log("CSVファイルをフォルダ内で見つけれませんでした。");
    Browser.msgBox("CSVファイルをフォルダ内で見つけれませんでした。")
  }
}

// メニューバーに表示
function onOpen(){
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('>>アクセスログチェッカー')
  menu.addItem('アクセスログをチェック', 'main')
  menu.addToUi()
}
