/**
 * マスタシートへ渡す条件となるモード指定
 * @type {string}
 */
const TARGET_MODE = "REQUIRED";

// テーブル定義の項目名判定に利用
const FIELD_NAME = "フィールド名";
const FIELD_TYPE = "種類";
const FIELD_MODE = "モード";
const FIELD_DESCRIPTION = "説明";

// 作成日時のセット(テーブル定義には含まれないため、個別にセット)
const CREATE_DATE = [
    'DATE',
    '作成日時',
    'REQUIRED'
];

// 更新日時のセット(テーブル定義には含まれないため、個別にセット)
const UPDATE_DATE = [
    'DATE',
    '更新日時',
    'REQUIRED'
]