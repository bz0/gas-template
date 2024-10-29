import { CONSTANTS, ERROR_MESSAGE } from "./constants";

type args = {
  message: string, 
  color: string | null,
  checkLineNum: number
}

export let existsCheckerSheet:GoogleAppsScript.Spreadsheet.Sheet | null = null;
export function sheetExistChecker () {
  try {
    const spreadSheet = SpreadsheetApp.getActive();
    Logger.log(spreadSheet);

    existsCheckerSheet = spreadSheet.getSheetByName(CONSTANTS.SHEET_NAME);
    if (!existsCheckerSheet) {
      return false;
    }

    const lastRowNum:number = existsCheckerSheet.getLastRow() - 1;
    if (lastRowNum === -1) {
      // データが存在しないとき
      Logger.log(ERROR_MESSAGE.SHEET_NAME_LIST_EMPTY_MESSAGE);
      return false;
    }

    Logger.log("lastRowNum:" + lastRowNum);

    // 存在チェックする列全て取得
    // ※空行の場合は空の配列がセットされる
    let sheetNameList: any = existsCheckerSheet.getRange(
      CONSTANTS.START_COLUMN_NUM, 
      CONSTANTS.CHECK_ROW_NUM, 
      lastRowNum
    ).getValues();
    Logger.log(sheetNameList);

    for (let i=0; i<sheetNameList.length; i++){
      let isSheet = null;
      if (sheetNameList[i] !== "") {
        // シート名が空でない時のみ
        isSheet = spreadSheet.getSheetByName(sheetNameList[i]);
      }

      const checkLineNum = CONSTANTS.START_COLUMN_NUM + i;
      const args:args = {
        "message"      : "",
        "color"        : null,
        "checkLineNum" : checkLineNum
      };

      if (!isSheet || (isSheet && isSheet.isSheetHidden())) {
        Logger.log('存在しないシート：' + sheetNameList[i]);
        args.color   = CONSTANTS.WARNING_COLOR;
        args.message = CONSTANTS.WARNING_MESSAGE;
      }

      // 結果をシートに書き込み
      setResult(args);
    }
  } catch (e: any) {
      console.warn(e.stack);
      console.warn("エラー発生");
  }
}

/**
 * シート存在有無でメッセージ出力・背景色付け
 */
export function setResult (args: args) {
    // メッセージ出力
    setMessage(args);
    // 背景色設定
    setBackgroundColor(args);
}


/**
 * メッセージ出力
 */
export function setMessage (args: args) {
  if (existsCheckerSheet) {
    const output_column = existsCheckerSheet.getRange(args.checkLineNum, CONSTANTS.RESULT_ROW_NUM);
    output_column.setValue(args.message);
  }
}

/**
 * 背景色設定
 */
export function setBackgroundColor (args: args) {
  if (existsCheckerSheet) {
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet?hl=ja#getRange(Integer,Integer,Integer,Integer)
    const output_background_color_range = existsCheckerSheet.getRange(args.checkLineNum, 1, 1, CONSTANTS.RESULT_ROW_NUM);
    output_background_color_range.setBackground(args.color);
  }
}
