/**
 * Whats this?
 * BigQueryからデータをスプレッドシート(以下、参照先スプレッドシート)のうち、必須項目に対してデータ入力規則と入力規則をセットします
 *
 * !!!caution!!!
 * 各シートのA1セルには、参照先スプレッドシートのリンクを埋め込んでください!!
 * 各シートのヘッダはA列：フィールド名、B列：種類、C列：モード、D列：説明の順にセットしてください!!
 * 参照先スプレッドシートにセットする入力書式もしくはデータ入力規則のカスタム関数を変更する場合は、settingオブジェクトの値を書き換えてください!!
 * モード指定を変更する場合は、TARGET_MODEの変数名を書き換えてください。(複数のモードには未対応です)
 * 参照先スプレッドシートの一番左のシートにBigQueryから連携するデータをセットしてください。
 */

// グローバル変数

/**
 * masterテーブルの情報を管理するオブジェクト
 * @type {{}}
 */
let tableObj = {};

/**
 * テーブルデータのうち、必須項目(REQUIRED)のフィールドを格納する配列
 * マスタシートへ入力規則と書式設定を反映する時に使います
 * @type {*[]}
 * Settings：種類(type)[0], 説明(description)[1], モード(mode)[2]
 */
let requiredArray = [];

/**
 * ハイパーリンクのセル指定に利用する変数
 * @type {string}
 */
let cellNotation = "";

function Main() {

    // テーブル定義のシートデータを連想配列として格納し、フォーマットが整合結果を返します
    let tableCategoryFlag = _tableCategory();

    // テーブル情報を取得できていれば、後続処理を開始します
    if (tableCategoryFlag) {

        for (let key in tableObj) {

            // 全テーブルデータのうち、REQUIREDモードのフィールドだけを抽出し直します
            _setRequiredData(key);

            // 参照先スプレッドシートの一番左のシートを参照します
            let targetSheet = _getReferenceSheet();

            // 参照先スプレッドシートの最終行、最終列をそれぞれ取得します
            let targetRange = _getSheetRange(targetSheet);
            // 参照先スプレッドシートの最終行
            let lastRow = targetRange[0];
            // 参照先スプレッドシートの最終列
            let lastColumn = targetRange[1];

            // 参照先スプレッドシートの必須項目列に対して、入力書式と入力規則をそれぞれセットします
            _setCustomFunction(lastRow, lastColumn, targetSheet);
            // 参照シートのカスタム関数セット完了したら、配列を空にします
            requiredArray = [];
        }
    } else {
        console.log("処理が失敗しました！カテゴリ名が誤っている可能性があります。");
    }
}

/**
 * @returns {boolean}
 * @private
 * テーブル定義のカテゴリデータを連想配列として格納します
 * カテゴリ名が誤っていた場合、セットせずに処理を終了するようにfalseを返します
 */
function _tableCategory() {

    let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    // スプレッドシートのシートを全て取得して、ループ。シート分処理を繰り返します
    for (let count = 0; count < sheets.length; count++) {
        tableObj["sheetNum" + count] = {}; // シートNo(キー)をセット
        tableObj["sheetNum" + count]["sheetName"] = sheets[count].getName(); // シート名をセット
        let lastRow = sheets[count].getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow(); // 最終行を指定

        // マスタシートA1にセットされているURLを取得します
        let targetLink = sheets[count].getRange(1, 1).getRichTextValues();
        tableObj["sheetNum" + count]["sheetLink"] = targetLink[0][0].getLinkUrl(); // シートURLをセット

        // 各データをキー配列として保持できるようにします
        tableObj["sheetNum" + count]["cellData"] = {
            field: [], // フィールド名をセット
            type: [], // 種類をセット
            mode: [], // モードをセット
            description: [] // 説明をセット
        };

        let cellData = sheets[count].getRange(1, 1, lastRow, 5).getValues(); // A列〜D列の最終行を指定

        var flag = (cellData, count) => {
            // シートの最終行までループさせて、各項目を配列に追加します(ポリシータグ、スキーマは使わないので除外)
            for (let sheetRow = 0; sheetRow < cellData.length; sheetRow++) {

                // カテゴリ名がデフォルト定義と異なっていた場合、処理を終了するためfalseを返します
                if ((sheetRow == 0) && ((cellData[sheetRow][0] !== FIELD_NAME) || (cellData[sheetRow][1] !== FIELD_TYPE) || (cellData[sheetRow][2] !== FIELD_MODE) || (cellData[sheetRow][4] !== FIELD_DESCRIPTION))) {
                    return false;
                }
                // もし他の項目を追加する必要があるときは、引数に追加でセットします
                tableObj["sheetNum" + count]["cellData"]["field"].push(cellData[sheetRow][0]); // フィールド名
                tableObj["sheetNum" + count]["cellData"]["type"].push(cellData[sheetRow][1]); // 種類
                tableObj["sheetNum" + count]["cellData"]["mode"].push(cellData[sheetRow][2]); // モード
                tableObj["sheetNum" + count]["cellData"]["description"].push(cellData[sheetRow][4]); // 説明
            }
            // 最終行まで処理完了したらtrueを返し、次シートを処理します
            return true;
        }
        // カテゴリ名が異なっていた場合、falseを返し処理を終了します
        if (!flag(cellData,count)) {
            return false;
        }
    }
    // 全て処理完了したら、trueを返します
    return true;
}

/**
 * テーブルデータのうち、必須項目(REQUIRED)のフィールドを配列へ格納します
 * @param key
 * @private
 */
function _setRequiredData(key) {
    // スプレッドシートのシートIDを取得します
    requiredArray.push(tableObj[key]["sheetLink"].split("/")[5]);
    // C列のモードが"REQUIRED"となっているフィールドのみ、配列へセットし直します
    for (let reqCount = 0; reqCount < tableObj[key]["cellData"]["field"].length; reqCount++) {
        // もし他のモードを追加する必要があるときは、以下に"or"で条件を追加します
        if ((tableObj[key]["cellData"]["mode"][reqCount] === TARGET_MODE) && (tableObj[key]["cellData"]["type"][reqCount] === "INTEGER" || tableObj[key]["cellData"]["type"][reqCount] === "STRING" || tableObj[key]["cellData"]["type"][reqCount] === "DATE")) {
            // もし他の項目を追加する必要があるときは、引数に追加でセットします
            requiredArray.push([tableObj[key]["cellData"]["type"][reqCount], tableObj[key]["cellData"]["description"][reqCount], tableObj[key]["cellData"]["mode"][reqCount]]);
        }
    }
    requiredArray.push(CREATE_DATE);
    requiredArray.push(UPDATE_DATE);
    // console.log(requiredArray);
}

/**
 * 参照先スプレッドシートの一番左のシートを参照します
 * @returns {*}
 * @private
 */
function _getReferenceSheet() {
    let targetSpreadsheet = SpreadsheetApp.openById(requiredArray[0]);
    let sheet = targetSpreadsheet.getSheets()[0];
    return sheet;
}

/**
 * 参照先スプレッドシートの最終行、最終列をそれぞれ取得します
 * @returns {*[]}
 * @private
 */
function _getSheetRange(targetSheet) {
    var row = targetSheet.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    var column = targetSheet.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
    return [row, column];
}

/**
 * 必須項目列に対して、入力書式と入力規則をセットします
 * @param lastRow
 * @param lastColumn
 * @param targetSheet
 * @private
 */
function _setCustomFunction(lastRow, lastColumn, targetSheet) {
    // 参照先スプレッドシートの最終行まで1行ずつ処理します
    for (let targetRow = 1; targetRow < lastColumn; targetRow++) {

        // 参照先スプレッドシートA列の項目名を取得します。この値は、requiredArray配列のdescriptionとの比較に使います
        let headerData = targetSheet.getRange(1, targetRow ,1, 1).getValue();
        // グローバルの変数に代入し、該当列に対して書式設定とデータ入力をセットできるようにします
        cellNotation = targetSheet.getRange(1, targetRow ,1, 1).getA1Notation().replace("1","");
        // console.log(cellNotation);

        /**
         シートにセットする入力書式と入力関数のカスタム関数をセットしておきます
         @type {{}}
         */
        let setting = {};
        setting.rule = {
            // formatオブジェクトに、条件付き書式設定のカスタム数式をセットします
            // キーはデータの型を指定します
            format: {
                "INTEGER": '=OR(isBlank($' + cellNotation + '2),ISNUMBER($' + cellNotation + '2)=FALSE,ISTEXT($' + cellNotation + '2)=TRUE,ISDATE($' + cellNotation + '2)=TRUE)',
                "STRING": '=OR(isBlank($' + cellNotation + '2),ISNUMBER($' + cellNotation + '2)=TRUE,ISTEXT($' + cellNotation + '2)=FALSE,ISDATE($' + cellNotation + '2)=TRUE)',
                "DATE": '=OR(isBlank($' + cellNotation + '2),ISNUMBER($' + cellNotation + '2)=TRUE,ISTEXT($' + cellNotation + '2)=FALSE,ISDATE($' + cellNotation + '2)=FALSE,REGEXMATCH($' + cellNotation + '2,"/")=TRUE)'
            }
        }

        // REQUIRED配列をループして、必要な比較ができるようにします
        for (let field = 1; field < requiredArray.length; field++) {
            // 参照先スプレッドシートの項目列と、descriptionが一致している場合、書式設定とデータ入力をセットします
            if (headerData === requiredArray[field][1]) {
                _setFormatRule(lastRow, field, targetSheet, setting, headerData);

                for (let cell = 2; cell < lastRow + 1; cell++) {
                    setting.rule = {
                        // inpuオブジェクトに、データ入力規則のカスタム数式をセットします
                        // キーはデータの型を指定します
                        input: {
                            "INTEGER": '=AND(ISNUMBER($' + cellNotation + cell + ')=TRUE,ISTEXT($' + cellNotation + cell + ')=FALSE,ISDATE($' + cellNotation + cell + ')=FALSE)',
                            "STRING": '=AND(ISNUMBER($' + cellNotation + cell + ')=FALSE,ISTEXT($' + cellNotation + cell + ')=TRUE,ISDATE($' + cellNotation + cell + ')=FALSE)',
                            "DATE": '=AND(ISNUMBER($' + cellNotation + cell + ')=FALSE,ISTEXT($' + cellNotation + cell + ')=TRUE,ISDATE($' + cellNotation + cell + ')=TRUE,REGEXMATCH($' + cellNotation + cell + ',"-")=TRUE)'
                        }
                    }
                    _setInputRule(cell, field, targetSheet, setting, headerData);
                }
            }
        }
    }
}

/**
 * 参照先スプレッドシートに対して、条件付き書式設定をセットします
 * @param lastRow
 * @param key
 * @param targetSheet
 * @param setting
 * @param header
 * @private
 */
function _setFormatRule(lastRow, key, targetSheet, setting, header) {
    let formatRule = SpreadsheetApp.newConditionalFormatRule()
        // 警告色セット
        .setBackground('#FF0000')
        // セットするカスタム数式
        .whenFormulaSatisfied(setting.rule.format[requiredArray[key][0]])
        // 適用範囲
        .setRanges([targetSheet.getRange(cellNotation + '2:' + cellNotation + lastRow + '')])
        .build();
    var rules = targetSheet.getConditionalFormatRules();
    rules.push(formatRule);
    targetSheet.setConditionalFormatRules(rules);
}

/**
 * 参照先スプレッドシートに対して、データの入力規則をセットします
 * @param cell
 * @param key
 * @param targetSheet
 * @param setting
 * @param header
 * @private
 */
function _setInputRule(cell, key, targetSheet, setting, header) {
    targetCell = cell;
    let inputRule = SpreadsheetApp.newDataValidation()
        // セットするカスタム数式
        .requireFormulaSatisfied(setting.rule.input[requiredArray[key][0]])
        // 無効なデータとしてセット
        .setAllowInvalid(false)
        .build();
    targetSheet.getRange(cellNotation + cell).setDataValidation(inputRule);
}