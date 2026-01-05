# OfflineList DAT Converter

OfflineList の DAT 形式 XML ファイルから `<game>` 部分を抽出して  
**CSV に変換するツール**、および  
変換した **CSV を OfflineList 形式の XML（DAT）に戻すツール**です。

* OfflineList DAT → CSV（Shift-JIS / UTF-8 自動判別）
* CSV → OfflineList DAT（configuration + games を自動生成）
* GUI 操作で簡単に変換可能
* ドラッグ＆ドロップ起動対応
* config.json による設定管理

本ツールの Python コードは、作者が要件を提示し、  
**Microsoft Copilot によって自動生成されたものです**。

---

## 📦 必須ライブラリ（.py のまま使用する場合）

### ✔ 外部ライブラリ  
bash
```
pip install beautifulsoup4
```
### ✔ 標準ライブラリ（インストール不要）

* csv
* json
* os
* sys
* datetime
* xml.etree.ElementTree
* xml.dom.minidom
* tkinter（Windows の公式 Python には標準で含まれています）

## 🚀 実行方法（Python）
generator.py と config.json を同じフォルダに置き、以下を実行します。  
bash
```
python generator.py
```
GUI が起動し、CSV/XML の変換が行えます。

## 🔄 変換機能の詳細
### ✔ XML（OfflineList DAT） → CSV
Shift-JIS / UTF-8 を自動判別して読み込み

<game> タグ以下の情報をすべて CSV に展開

日本語タイトル・コメントも正しく抽出

### ✔ CSV → XML（OfflineList DAT）
<configuration> 部分を自動生成

<games> 部分を CSV から構築

OfflineList でそのまま読み込める DAT を生成

## 🌐 config.json の base\_url について
base\_url は、OfflineList の DAT 更新機能で使用される
DAT 配布サーバーのベース URL（共通部分） を指定する項目です。

生成される XML の以下の項目に使用されます：

* DATのバージョン情報  
  <base\_url>\\game.txt  
* DAT本体  
  <base\_url>\\game.zip  
* 画像フォルダ  
  <base\_url>\\game\\  

例：
DAT を https://example.com/offlinelist/ に置く場合  
json
```
"base\_url": "https://example.com/offlinelist/"
```
空欄のままでも動作しますが、
OfflineList の「オンライン更新機能」を使う場合は設定が必要です。

---

## 🧪 exe 化（PyInstaller）
Python をインストールしていない環境でも使えるように
exe を作成することができます。

1. PyInstaller のインストール  
   bash
   ```
   pip install pyinstaller
   ```
3. exe の作成（config.json を外部ファイルのまま使う場合）  
   bash
   ```
   pyinstaller --noconsole generator.py
```
生成物は以下に出力されます：

コード
```
dist/
└ generator/
    ├ generator.exe
    ├ \_internal/（Python ランタイム）
    └ その他 DLL
```
generator.exe と同じフォルダに config.json  を置いて使用します。
---

## 📄 ライセンス
このプロジェクトは MIT License のもとで公開されています。
詳細は LICENSE を参照してください。

## 🙏 開発について
本ツールの Python コードは、作者が要件を提示し、
Microsoft Copilot によって自動生成されたものです。












