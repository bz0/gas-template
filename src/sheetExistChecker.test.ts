// sheetChecker.test.ts
import { sheetExistChecker, existsCheckerSheet, setResult } from "./sheetExistChecker";
import { CONSTANTS, ERROR_MESSAGE } from "./constants";

// グローバルオブジェクトのモック
const mockSpreadsheet = {
  getSheetByName: jest.fn().mockReturnThis(),
  getLastRow: jest.fn().mockReturnValue(2),
  getRange: jest.fn().mockReturnThis(),
  getValues: jest.fn().mockReturnValue([["Sheet1"], ["Sheet2"]]),
  setValue: jest.fn(),
  setBackground: jest.fn(),
  isSheetHidden: jest.fn().mockReturnValue(false),
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

  describe("異常処理", () => {
    it("シートが存在しない場合、エラーメッセージを返すこと", () => {
        mockSpreadsheet.getSheetByName.mockReturnValueOnce(null); // シートが存在しない場合
        const result = sheetExistChecker();
        expect(result).toStrictEqual({errorMessage: ERROR_MESSAGE.SHEET_FETCH_FAILURE_MESSAGE});
    });
    
    it("シートが存在し行が空の場合、エラーメッセージを返すこと", () => {
        mockSpreadsheet.getSheetByName.mockReturnValueOnce(mockSpreadsheet);
        mockSpreadsheet.getLastRow.mockReturnValue(0); // データがない場合
        const result = sheetExistChecker();
        expect(result).toStrictEqual({errorMessage: ERROR_MESSAGE.SHEET_NAME_LIST_EMPTY_MESSAGE});
    });
  })

  describe("正常処理", () => {
    it("シートが存在し、データがある場合、シートチェックが行われること", () => {
        mockSpreadsheet.getSheetByName.mockReturnValueOnce(mockSpreadsheet); // シートが存在する場合
        mockSpreadsheet.getValues.mockReturnValue([["Sheet1"], ["Sheet2"]]); // シート名が2つ存在

        const result = sheetExistChecker();
        expect(result).toStrictEqual({sheetNameList: [["Sheet1"], ["Sheet2"]], acceptSheetNameList: ["Sheet1", "Sheet2"]});
    });
    
    it("存在しないシートの場合、警告メッセージを設定すること", () => {
        mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
            if (name === "Sheet1") return null; // Sheet1は存在しない
            return mockSpreadsheet; // 他のシートは存在する
        });

        sheetExistChecker();
        const result = sheetExistChecker();
        expect(result).toStrictEqual({sheetNameList: [["Sheet1"], ["Sheet2"]], acceptSheetNameList: ["Sheet2"]});
    });
  })
});
