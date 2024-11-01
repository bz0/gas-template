// sheetChecker.test.ts
import { sheetExistChecker, getSheetNameList, getExistingSheetNameList, getSpreadSheet } from "./sheetExistChecker";
import { CONSTANTS, ERROR_MESSAGE } from "./constants";

/**
 * - 前提：GAS独自のオブジェクトやメソッドはモックオブジェクトで代用
 * - 目的：
 *  - シートの存在チェックが意図通りに動いていることを確認する
 *  - シート取得失敗した等の異常が起きたときの挙動を確認する
 */

// グローバルオブジェクトのモック
const mockSpreadsheet = {
  getSheetByName: jest.fn().mockReturnThis(),
  getLastRow: jest.fn(),
  getRange: jest.fn().mockReturnThis(),
  getValues: jest.fn(),
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

describe("sheetExistChecker関数", () => {
    beforeEach(() => {
        jest.clearAllMocks(); // 全てのモックのクリア
    });

    describe("正常処理", () => {
        it("シート名が3件存在する場合、シート名を3件返すこと", () => {
            mockSpreadsheet.getSheetByName.mockReturnValueOnce(mockSpreadsheet);
            mockSpreadsheet.getLastRow.mockReturnValue(3);
            mockSpreadsheet.getValues.mockReturnValue([["Sheet1"], ["Sheet2"], ["Sheet3"]]);
            const result = sheetExistChecker();
            expect(result).toStrictEqual({"existingSheetNameList": ["Sheet1", "Sheet2", "Sheet3"], "sheetNameList": [["Sheet1"], ["Sheet2"], ["Sheet3"]]});
        });
    })

    describe("異常処理", () => {
        it("シートが存在しない場合、エラーメッセージを返すこと", () => {
            global.SpreadsheetApp.getActive = jest.fn().mockReturnValue(null);
            const result = sheetExistChecker();
            expect(result).toStrictEqual({ errorMessage : "シートの取得に失敗しました" });
        });
    })

    afterEach(() => {
        // テスト終了後に初期値に戻す
        global.SpreadsheetApp.getActive = jest.fn().mockReturnValue(mockSpreadsheet);
    });
})

describe("getSpreadSheet関数", () => {
    beforeEach(() => {
        jest.clearAllMocks(); // 全てのモックのクリア
    });

    describe("異常処理", () => {
        it("シートが存在しない場合、エラーメッセージを返すこと", () => {
            function test () {
                global.SpreadsheetApp.getActive = jest.fn().mockReturnValue(null);
                getSpreadSheet();
            }

            expect(test).toThrow(new Error('シートの取得に失敗しました'));
            global.SpreadsheetApp.getActive = jest.fn().mockReturnValue(mockSpreadsheet); // テスト終了後に初期値に戻す
        });
    })
})

describe("getSheetNameList関数", () => {
    beforeEach(() => {
        jest.clearAllMocks(); // 全てのモックのクリア
        getSpreadSheet();
    });

    describe("正常処理", () => {
        it("シート名取得し、シート名配列を返すこと", () => {
            mockSpreadsheet.getSheetByName.mockReturnValueOnce(mockSpreadsheet);
            mockSpreadsheet.getLastRow.mockReturnValue(3);
            mockSpreadsheet.getValues.mockReturnValue([["Sheet1"], ["Sheet2"], ["Sheet3"]]);
            const spreadSheet = getSpreadSheet();
            const result = getSheetNameList(spreadSheet);
            expect(result).toStrictEqual([["Sheet1"], ["Sheet2"], ["Sheet3"]]);
        });
    })

    describe("異常処理", () => {
        it("シート取得に失敗した場合、エラーメッセージを返すこと", () => {
            function test() {
                mockSpreadsheet.getSheetByName.mockReturnValueOnce(null); // シートが存在しない場合
                const spreadSheet = getSpreadSheet();
                getSheetNameList(spreadSheet);
            }
            expect(test).toThrow(new Error('シートの取得に失敗しました'));
        });
        
        it("シートが存在し行が空の場合、エラーメッセージを返すこと", () => {
            function test() {
                mockSpreadsheet.getSheetByName.mockReturnValueOnce(mockSpreadsheet);
                mockSpreadsheet.getLastRow.mockReturnValue(0); // データがない場合
                const spreadSheet = getSpreadSheet();
                getSheetNameList(spreadSheet);
            }
            expect(test).toThrow(new Error('シート名が入力されてない為処理終了'));
        });
    })
})

describe("getExistingSheetNameList関数", () => {
    beforeEach(() => {
        jest.clearAllMocks(); // 全てのモックのクリア
    });

    describe("正常処理", () => {
        it("Sheet2・Sheet3のみ存在する場合、存在するシート名リストにSheet2・Sheet3を返すこと", () => {
            mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
                if (name === "Sheet1") return null; // Sheet1は存在しない
                return mockSpreadsheet; // 他のシートは存在する
            });

            const spreadSheet = getSpreadSheet();
            const result = getExistingSheetNameList(spreadSheet, [["Sheet1"], ["Sheet2"], ["Sheet3"]]);
            console.warn(result);
            expect(result).toStrictEqual({"existingSheetNameList": ["Sheet2", "Sheet3"], "sheetNameList": [["Sheet1"], ["Sheet2"], ["Sheet3"]]});
        });
    })

    describe("異常処理", () => {
        it("シートが存在しない場合、エラーメッセージを返すこと", () => {
            function test () {
                getExistingSheetNameList(null, [["Sheet1"], ["Sheet2"], ["Sheet3"]]);
            }

            expect(test).toThrow(new Error('シートの取得に失敗しました'));
        });
    })
})