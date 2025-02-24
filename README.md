# Money Forward ME自動データ入力プログラム

## 概要
- このプログラムは、Money Forward MEに対して自動的にデータを入力するためのものです。
- Excelファイルからデータを読み込み、ウェブスクレイピングを使用してデータを入力します。
- プログラムは、Python言語で実装されています。


## 入力
- 設定ファイル: Money Forward MEへのログイン情報やExcelファイルのパスなどの設定が記述されています。
- Excelファイル: Money Forward MEに入力するデータが含まれているExcelファイルが必要です。

## 出力
- ログファイル: プログラムの実行ログが記録されます。ログには、処理の開始、終了、エラー情報などが含まれます。

## 処理詳細
1. Excelファイルからデータの読み込み
   - 指定されたExcelファイルのテーブルからデータを読み取ります。
   - Excelファイル内の特定のテーブルからデータを抽出します。
   
2. Money Forward MEへのログイン
   - Money Forward MEに自動的にログインします。
   - ユーザー名とパスワードは、設定ファイルから読み込まれます。
   
3. データの入力
   - Excelファイルから読み込んだデータをMoney Forward MEの入出金登録フォームに自動的に入力します。
   - 各行のデータには、日付、大項目、中項目、内容、金額が含まれます。

## 想定実行環境
- Windows OSのPC
- Python 3.7以降
- 次のPythonライブラリー
  - openpyxl
  - pandas
  - selenium
- Microsoft Edge WebDriver

## 検証環境
- OS：Microsoft Windows 11 Home
- CPU：Intel64 Family 6 Model 154 Stepping 3 GenuineIntel ~2100 Mhz
- メモリー：16 GB
- ネットワーク：1 Gbps
- SSDの残容量：233 GB
- Python 3.11.4
- 次のPythonライブラリー
  - openpyxl==3.1.2
  - pandas==2.2.2
  - selenium==4.21.0
- Microsoft Edge WebDriver バージョン 125.0.2535.67

## ライセンス
このプログラムは、MITライセンスの下で提供されます。
