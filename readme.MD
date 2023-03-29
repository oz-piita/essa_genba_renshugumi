# RMsystem version 4.3

エッサの練習メニューを、組合せ最適化で計算する内輪向けのアプリです。

練習組みが大変な公演ならば、時間短縮効果が期待できます。

ちなみにRMはアルゴリズムを開発した方のイニシャルです。

# Features

バージョン4.3はPC向けGUIであり、実務上はWindowsOSでのexeファイルを使用することを想定しています。

入力をGoogleスプレッドシートからダウンロードしたExcelファイルから行います。

# Requirement

必要なPythonライブラリは以下の通りです。

* tkinter 8.6
* pandas 1.2.4
* pulp    2.4

# Usage(test)

Requirementのうち、pulpは標準パッケージに含まれませんので、先にインストールしてください。

実行方法は以下です。
```bash
git clone https://github.com/oz-piita/essa_genba_renshugumi
```
アプリケーションを起動します。

```bash
python essa_genba_renshugumi/main.py
```

起動したら、ファイルの場所にダウンロードしたExcelを指定し、
日付IDとして"d1"～"d7"を入力した状態で実行ボタンを押下します。

すると右の結果欄に計算結果が表示されます。

# How to convert it into a .exe file with pyinstaller

配布を容易にするために、pyintallerを使用してexe化する場合は、以下のコマンドを一行で実行してください。

```bash
pyinstaller --noconfirm --onefile --console --add-data "C:/Users/.../Lib/site-packages/pulp;pulp/" "C:/Users/.../main.py"
```
2つのpahtは、前者が"pythonフォルダ内のpulpフォルダへのパス;実行ファイル保存先のパス"、"後者がこのアプリケーションのパス"です。

なお、Pythonフォルダ内のpulpフォルダへのパスは、以下のpythonコードを実行することでも確認できます。
```pyhon
import pulp
print(pulp.__path__[0])
```

こちらを参考にしています。

https://stackoverflow.com/questions/60926887/py-project-works-but-pyinstaller-gives-error-related-to-pulp

# Note

核になっている組合せ最適化処理はplugins/my_calcファイルにモジュール化されています。

最適化を試験したいときは、calcのオブジェクトに適合するようデータを加工して食わせてください。

データ加工はmy_inputを参考にしてください。

# Author

最適化アルゴリズム
* ryo(11th)

入力及びUI作成
* taipi(13th)

スプレッドシート構成
* mamo(12th)
* taipi
