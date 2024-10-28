"use strict";

// シート存在チェック実行
function SheetExistCheckerExec () {
  SheetExistChecker.exec();
}

type Define = {
    readonly SHEET_NAME : string,
    readonly CHECK_ROW_NUM : number, // 存在チェックする列
    readonly START_COLUMN_NUM : number, // 存在チェック開始行
    readonly RESULT_ROW_NUM : number, // 存在チェックの結果を色付けする行数
    readonly WARNING_COLOR : string, // ワーニング色指定
    readonly WARNING_MESSAGE : string, // シートがないときの警告文
    readonly SHEET_NAME_LIST_EMPTY_MESSAGE : string,
}

const SheetExistChecker:{
    define: Define;
    existsCheckerSheet: GoogleAppsScript.Spreadsheet.Sheet|null;
    exec: () => void;
} = {
  define: {
    SHEET_NAME : "チェック", // 存在チェックするシート名
    CHECK_ROW_NUM : 1, // 存在チェックする列
    START_COLUMN_NUM : 2, // 存在チェック開始行
    RESULT_ROW_NUM : 3, // 存在チェックの結果を色付けする行数
    WARNING_COLOR : "#FFB2B2", // ワーニング色指定
    WARNING_MESSAGE : "シートなし", // シートがないときの警告文
    SHEET_NAME_LIST_EMPTY_MESSAGE : "シート名が1つも入力されてない為処理終了",
  },
  existsCheckerSheet : null, // シートオブジェクト
  /**
   * シート存在有無チェック
   */
  exec: function () {
    try {
      const spreadSheet = SpreadsheetApp.getActive();
      Logger.log(spreadSheet);

      this.existsCheckerSheet = spreadSheet.getSheetByName(this.define.SHEET_NAME);
      if (!this.existsCheckerSheet) {
        return false;
      }

      const lastRowNum: number = this.existsCheckerSheet.getLastRow() - 1;
      if (lastRowNum === -1) {
        // データが存在しないとき
        Logger.log(this.define.SHEET_NAME_LIST_EMPTY_MESSAGE);
        return false;
      }

      Logger.log("lastRowNum:" + lastRowNum);

      // 存在チェックする列全て取得
      // ※空行の場合は空の配列がセットされる
      let sheetNameList: string[][] = this.existsCheckerSheet.getRange(
        this.define.START_COLUMN_NUM, 
        this.define.CHECK_ROW_NUM, 
        lastRowNum
      ).getValues();
      Logger.log(sheetNameList);

      for (let i=0; i<sheetNameList.length; i++){
        let isSheet = null;
        if (sheetNameList[i] !== "") {
          // シート名が空でない時のみ
          isSheet = spreadSheet.getSheetByName(sheetNameList[i]);
        }
        /** @type {number} */
        const checkLineNum = this.define.START_COLUMN_NUM + i;
        /** @type {{ message: string, color: string | null, checkLineNum: number }} */
        const args = {
          "message"      : "",
          "color"        : null,
          "checkLineNum" : checkLineNum
        };

        if (!isSheet || (isSheet && isSheet.isSheetHidden())) {
          Logger.log('存在しないシート：' + sheetNameList[i]);
          args.color   = this.define.WARNING_COLOR;
          args.message = this.define.WARNING_MESSAGE;
        }

        // 結果をシートに書き込み
        this.setResult(args);
      }
    } catch (e) {
        console.warn(e.stack);
        console.warn("エラー発生");
    }
  },
  /**
   * シート存在有無でメッセージ出力・背景色付け
   * @param {{ message: string, color: string | null, checkLineNum: number }} args - シート出力する結果情報
   */
  // シート存在有無でメッセージ出力・背景色付け
  setResult : function(args) { // todo:複数の引数が分かりづらい
      // メッセージ出力
      this.setMessage(args);
      // 背景色設定
      this.setBackgroundColor(args);
  },
  /**
   * メッセージ出力
   * @param {{ message: string, color: string | null, checkLineNum: number }} args - シート出力する結果情報
   */
  setMessage : function (args) {
      const output_column = this.existsCheckerSheet.getRange(args.checkLineNum, this.define.RESULT_ROW_NUM);
      output_column.setValue(args.message);
  },

  /**
   * 背景色設定
   * @param {{ message: string, color: string | null, checkLineNum: number }} args - シート出力する結果情報
   */
  setBackgroundColor : function (args) {
      // https://developers.google.com/apps-script/reference/spreadsheet/sheet?hl=ja#getRange(Integer,Integer,Integer,Integer)
      const output_background_color_range = this.existsCheckerSheet.getRange(args.checkLineNum, 1, 1, this.define.RESULT_ROW_NUM);
      output_background_color_range.setBackground(args.color);
  }
}
