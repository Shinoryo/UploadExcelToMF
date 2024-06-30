# インストールガイド

## Python環境のセットアップ
- Python 3.7以降のインストールが必要です。公式ウェブサイト( https://www.python.org/ )からPython 3.7以降をダウンロードし、インストールしてください。

## 必要なライブラリのインストール
- コマンドラインまたはターミナルを開き、以下のコマンドを実行して必要なPythonライブラリをインストールしてください。
  ```
  $ pip install openpyxl pandas selenium
  ```

## WebDriverの準備
- Microsoft EdgeのWebDriverが必要です。公式ウェブサイト( https://developer.microsoft.com/microsoft-edge/tools/webdriver/ )から使用しているMicrosoft Edgeのバージョンに対応したものをダウンロードしてください。
- WebDriverは、`msedgedriver.exe`というファイル名で、upload_excel_to_mf.pyと同じフォルダーに配置してください。

## 設定ファイルの編集:
- プログラム内で使用される設定ファイル(config.ini)を編集します。
  | 項目名           | 説明                                     |
  |------------------|------------------------------------------|
  | `input_file`     | Excelファイルのパス                       |
  | `table_name`     | Excelファイル内のテーブル名               |
  | `user`           | Money Forward MEのユーザー名(メールアドレス) |
  | `password`       | Money Forward MEのパスワード               |
  | `signin_url`     | Money Forward MEのサインインURL            |
  | `input_url`      | データ入力画面のURL                       |
  | `wallet_xpath`   | 入出金元のXPATH                           |
- 設定ファイルは、`config.ini`というファイル名で、UploadExcelToMF.pyと同じフォルダーに配置してください。

## Excelファイルの編集:
- プログラム内で使用されるExcelファイル(input_table.xlsx)を編集します。
  | 項目名             | 説明                                             |
  |--------------------|--------------------------------------------------|
  | `Date`             | 日付(yyyy/mm/ddの形式)                           |
  | `Large Category`   | 大項目名                                          |
  | `Middle Category`  | 中項目名                                          |
  | `Content`          | 内容                                              |
  | `Amount`           | 金額(正の整数の場合は収入に、0以下の整数の場合は支出) |
- Excelファイルは、設定ファイルで指定したパスに配置してください。

## プログラムの実行
- コマンドラインまたはターミナルで、プログラムの保存されているフォルダーに移動します。
- 以下のコマンドを実行して、プログラムを起動します。
  ```
  $ python upload_excel_to_mf.py
  ```
- プログラムが開始され、処理が自動的に実行されます。
  - 処理の途中で例外が発生した場合、その場でWebDriverをシャットダウンして処理を終了します。
  - logフォルダ内に作成されたログファイルを参照して、処理の詳細を確認できます。
    - ログファイルは、`log_yyyymmdd.log`というファイル名で出力されます。
    - ログには、処理の開始、終了、エラー情報などが含まれます。
