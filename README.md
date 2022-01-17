# プログラム概要

BigQueryのテーブルスキーマを元に、REQUIREDモードのフィールドを取得してソース元のスプレッドシートにあるフィールド名と一致した列に対してデータの入力規則と入力書式をセットします
フォーマットがあっていれば、複数の外部データシートに対して入力書式のセットが可能です。

# 事前準備

A.BigQueryのデータベースに登録されている各テーブルスキーマの情報を以下のような形でスプレッドシートにまとめておきます。
<img width="1358" alt="スクリーンショット 2022-01-16 22 30 10" src="https://user-images.githubusercontent.com/43813301/149662019-1e4ce732-0620-42bd-a0e3-ca11c78ba080.png">
※このスプレッドシートは以降、テーブルシートとします。

B.実際にBigQueryへ読み込む外部データのスプレッドシートを項目毎に並べます。この時項目名はテーブルスキーマの説明と一致していることが必要なことに注意してください。
<img width="2467" alt="スクリーンショット 2022-01-16 22 28 38" src="https://user-images.githubusercontent.com/43813301/149661967-49275979-2730-467e-9b4b-f0550b45ace0.png">
※このスプレッドシートは以降、外部データシートとします。

#使い方

1.テーブルシートのGoogleAppScriptを開いて、main.jsとconfig.jsをコピーします。
2.config.jsを開いて、テーブル定義の項目名を調整します。
3.main.jsを開いて、セルの入力書式にセットするカスタム数式を調整します。 setting.ruleオブジェクトのformatが該当します。
4.3と同様にデータの入力書式にセットするカスタム数式を調整します。setting.ruleオブジェクトのinputが該当します。
5.関数Main()を実行します。

#フォルダ構成

・README.md：このファイル
・config.js：設定ファイル
・main.js:メインロジック。外部データシートに対してセルの入力書式とデータ入力規則をセットする

#注意点
* 各シートのA1セルには、外部データシートのリンクを埋め込んでください!!
* 各シートのヘッダはA列：フィールド名、B列：種類、C列：モード、D列：説明の順にセットしてください!!
* 外部データしーとにセットする入力書式もしくはデータ入力規則のカスタム関数を変更する場合は、setting.ruleオブジェクトの値を書き換えてください!!
* モード指定を変更する場合は、TARGET_MODEの変数名を書き換えてください。(現時点では複数のモードには未対応です)
* 外部データシートは一番左のシートに対してBigQueryから連携するデータをセットしてください。(一番左のシート)