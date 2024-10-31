// sheetChecker.test.ts
import { sheetExistChecker, existsCheckerSheet, setResult } from "./sheetExistChecker";
import { CONSTANTS, ERROR_MESSAGE } from "./constants";

// グローバルオブジェクトのモック
const mockSpreadsheet = {
  getSheetByName: jest.fn(),
  getLastRow: jest.fn().mockReturnValue(2),
  getRange: jest.fn().mockReturnThis(),
  getValues: jest.fn().mockReturnValue([["Sheet1"], ["Sheet2"]]),
};

global.SpreadsheetApp = {
  getActive: jest.fn().mockReturnValue(mockSpreadsheet),
} as any;

global.Logger = {
  log: jest.fn(),
} as any;

describe("sheetExistChecker", () => {
  beforeEach(() => {
    jest.clearAllMocks(); // 全てのモックのクリア
  });

  describe.only("異常処理", () => {
    it("シートが存在しない場合、エラーメッセージを返すこと", () => {
        mockSpreadsheet.getSheetByName.mockReturnValueOnce(null); // シートが存在しない場合
        const result = sheetExistChecker();
        expect(result?.errorMessage).toBe(ERROR_MESSAGE.SHEET_FETCH_FAILURE_MESSAGE);
    });
    
    it("シートが存在し行が空の場合、エラーメッセージを返すこと", () => {
        mockSpreadsheet.getSheetByName.mockReturnValueOnce(mockSpreadsheet);
        mockSpreadsheet.getLastRow.mockReturnValue(0); // データがない場合
        const result = sheetExistChecker();
        expect(result?.errorMessage).toBe(ERROR_MESSAGE.SHEET_NAME_LIST_EMPTY_MESSAGE);
    });
  })

  it("シートが存在し、データがある場合、シートチェックが行われること", () => {
    mockSpreadsheet.getSheetByName.mockReturnValueOnce(mockSpreadsheet); // シートが存在する場合
    mockSpreadsheet.getValues.mockReturnValue([["Sheet1"], ["Sheet2"]]); // シート名が2つ存在

    sheetExistChecker();

    expect(Logger.log).toHaveBeenCalledWith("lastRowNum:1");
    expect(mockSpreadsheet.getRange).toHaveBeenCalledWith(
      CONSTANTS.START_COLUMN_NUM,
      CONSTANTS.CHECK_ROW_NUM,
      1
    );
  });

  it("存在しないシートの場合、警告メッセージを設定すること", () => {
    mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
      if (name === "Sheet1") return null; // Sheet1は存在しない
      return mockSpreadsheet; // 他のシートは存在する
    });

    sheetExistChecker();

    expect(Logger.log).toHaveBeenCalledWith("存在しないシート：Sheet1");
    expect(setResult).toHaveBeenCalledWith({
      message: CONSTANTS.WARNING_MESSAGE,
      color: CONSTANTS.WARNING_COLOR,
      checkLineNum: CONSTANTS.START_COLUMN_NUM,
    });
  });
});
