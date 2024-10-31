export const CONSTANTS = {
    SHEET_NAME : "チェック", // 存在チェックするシート名
    CHECK_ROW_NUM : 1, // 存在チェックする列
    START_COLUMN_NUM : 2, // 存在チェック開始行
    RESULT_ROW_NUM : 3, // 存在チェックの結果を色付けする行数
    WARNING_COLOR : "#FFB2B2", // ワーニング色指定
    WARNING_MESSAGE : "シートなし" // シートがないときの警告文
} as const satisfies Record<string, string | number>

export const ERROR_MESSAGE = {
    SHEET_NAME_LIST_EMPTY_MESSAGE : "シート名が入力されてない為処理終了",
    SHEET_NAME_FETCH_FAILURE_MESSAGE : "チェック対象のシート名が取得できませんでした",
    SHEET_FETCH_FAILURE_MESSAGE : "シートの取得に失敗しました"
} as const satisfies Record<string, string>