import { CONSTANTS, ERROR_MESSAGE } from "./constants";

type Args = {
  message: string, 
  color: string | null,
  checkLineNum: number
}

type ErrorResult = {
  errorMessage?: string
}

type existingSheetNameList = string[];

type Result = {
  sheetNameList?: string[][],
  existingSheetNameList?: existingSheetNameList
}

type sheetNameList = string[][];

export let existsCheckerSheet:GoogleAppsScript.Spreadsheet.Sheet | null = null;
export let spreadSheet:GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;

export function sheetExistChecker () {
  let result: ErrorResult | Result | null = null;
  let sheetNameList: string[][];

  try {
    // スプレッドシートオブジェクトを取得
    spreadSheet = SpreadsheetApp.getActive();
    if (!spreadSheet) {
      throw new Error(ERROR_MESSAGE.SHEET_FETCH_FAILURE_MESSAGE);
    }
    sheetNameList = getSheetNameList(spreadSheet);
    result = getExistingSheetNameList(sheetNameList);
  } catch (e: unknown) {
    if (e instanceof Error) {
      console.warn(e.stack);
      console.warn("エラー発生");
      result = { errorMessage : e.message };
    }
  }

  return result;
}

/**
 * 存在するシート名リストを取得する
 * @param sheetNameList 
 * @returns 
 */
export function getExistingSheetNameList (sheetNameList: sheetNameList): Result {
  let existingSheetNameList:existingSheetNameList = [];
  if (!spreadSheet) {
    throw new Error(ERROR_MESSAGE.SHEET_FETCH_FAILURE_MESSAGE);
  }

  for (const [i, [sheetName]] of sheetNameList.entries()) {
    let isSheet = sheetName ? spreadSheet.getSheetByName(sheetName) : null;
  
    const checkLineNum = CONSTANTS.START_COLUMN_NUM + i;
    const args: Args = {
      message: "",
      color: isSheet && !isSheet.isSheetHidden() ? null : CONSTANTS.WARNING_COLOR,
      checkLineNum
    };
    
    if (!isSheet || (isSheet && isSheet.isSheetHidden())) {
      args.message = CONSTANTS.WARNING_MESSAGE;
    } else {
      existingSheetNameList.push(sheetName);
    }
    
    setResult(args);
  }

  console.warn(sheetNameList);

  return {
    sheetNameList: sheetNameList,
    existingSheetNameList: existingSheetNameList
  }
}

/**
 * シート名リストを取得する
 * @param spreadSheet スプレッドシートオブジェクト
 * @returns sheetNameList シート名リスト
 */
export function getSheetNameList (spreadSheet:GoogleAppsScript.Spreadsheet.Spreadsheet):sheetNameList {
  let sheetNameList: string[][];

  if (!spreadSheet) {
    throw new Error(ERROR_MESSAGE.SHEET_FETCH_FAILURE_MESSAGE);
  } 

  // シートオブジェクト取得
  existsCheckerSheet = spreadSheet.getSheetByName(CONSTANTS.SHEET_NAME);
  if (!existsCheckerSheet) {
    throw new Error(ERROR_MESSAGE.SHEET_FETCH_FAILURE_MESSAGE);
  }

  // シート行数取得
  const lastRowNum:number = existsCheckerSheet.getLastRow() - 1;
  if (lastRowNum === -1) {
    // データが存在しないとき
    throw new Error(ERROR_MESSAGE.SHEET_NAME_LIST_EMPTY_MESSAGE);
  }

  // 存在チェックする列全て取得
  // ※空行の場合は空の配列がセットされる
  sheetNameList = existsCheckerSheet.getRange(
    CONSTANTS.START_COLUMN_NUM, 
    CONSTANTS.CHECK_ROW_NUM, 
    lastRowNum
  ).getValues();
  Logger.log(sheetNameList);

  return sheetNameList;
}

/**
 * シート存在有無でメッセージ出力・背景色付け
 */
export function setResult (args: Args) {
    // メッセージ出力
    setMessage(args);
    // 背景色設定
    setBackgroundColor(args);
}


/**
 * メッセージ出力
 */
export function setMessage (args: Args) {
  if (existsCheckerSheet) {
    const output_column = existsCheckerSheet.getRange(args.checkLineNum, CONSTANTS.RESULT_ROW_NUM);
    output_column.setValue(args.message);
  }
}

/**
 * 背景色設定
 */
export function setBackgroundColor (args: Args) {
  if (existsCheckerSheet) {
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet?hl=ja#getRange(Integer,Integer,Integer,Integer)
    const output_background_color_range = existsCheckerSheet.getRange(args.checkLineNum, 1, 1, CONSTANTS.RESULT_ROW_NUM);
    output_background_color_range.setBackground(args.color);
  }
}
