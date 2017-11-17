# tku_syllabus_check
Campusmateシラバス・アップロード用Excelデータの自動点検ツール

# Usage
Campusmateシラバス管理システムより点検対象のシラバスデータ（Excel形式）をダウンロードしておきます。
事前に指定した検証項目を実装し、検証結果をWord形式として書き出す。

# Instructions
0. Python3系の仮想環境を設定し、`pip install -r requirements.txt`で依存関係をインストールします。
1. Excelデータを確認しながら`positions.py`において項目の位置を設定します。
2. `instructions.py`の中に、点検項目に必要なメッセージ（項目を通過できなかった場合のメッセージ）、テスト関数を記入しておく。helper関数が必要な場合、`helpers.py`に入れておきます。最後に、`checks.append([message, test_name])`でテストを点検リストにつけます。
3. 端末からExcelデータの保存したフォルダに移動し、`python /path/to/syllabus_check.py`を起こします。（もしくは引数にパス名を指定する⇒ `python /path/to/syllabus_check.py /path/to/xls_files`）。Excelデータフォルダのなかにresultsフォルダーが生成され、点検結果の報告書が生成されます。

