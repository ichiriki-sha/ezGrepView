# 🔍 ezGrepView

サクラエディタの grep 検索結果を  
📊 **Excel マクロ（VBA）で読み込み・整形・可視化** するツールです。

大量の検索結果を Excel 上でスッキリ整理し、  
コメント行  、一致文字列 を色分け表示しながら、  
**ショートカット一発でエディタへジャンプ** できます。

---

## ✨ 特徴

- サクラエディタの grep 出力結果を Excel に取り込み
- grep ヘッダから検索条件を自動解析
- 検索結果を表形式で一覧表示
- 一致キーワードを **赤色＋太字** でハイライト
- コメント行を **緑色** で強調表示
- バイナリ行・文字化け行を自動判定
- `Ctrl + J` で **サクラエディタを起動し該当位置へジャンプ**
- 結果専用の Excel ファイルを自動生成

---

## 🖥 動作環境

- Windows 10 / 11
- Microsoft Excel（マクロ有効 / `.xlsm`）
- PowerShell 5.1 以上
- サクラエディタ  v2.4.1以上

---

## 🚀インストール & セットアップ手順

本プロジェクトは **Excel VBA を GitHub で安全に管理するための構成**を採用しています。  
VBA ソースはすべてテキスト（`.bas / .cls / .dcm`）として管理し、  
Excel ファイル（`.xlsm`）は **PowerShell により自動生成**します。

## 1. リポジトリの取得

```powershell
git clone https://github.com/ichiriki-sha/ezGrepView.git
cd ezGrepView
```

## 2. ディレクトリ構成

```powershell
ezGrepView/
├─ src/                        # Git管理の正
│   ├─ modBusinessCommon.bas
│   ├─ modBusinessMain.bas
│   …
│   └─ ThisWorkbook.dcm
├─ excel/
│   └─ ezGrepView.xlsm        # exportで自動生成（VBAなし）
├─ dev/
│   └─ ezGrepView.xlsm        # 開発用
├─ tools/
│   ├─ export_vba.ps1         # dev/Excel → src/ + excel/Excel
│   └─ import_vba.ps1         # src/ + excel/Excel → excel/Excel 
├─ .gitignore
└─ README.md
```

## 3. Excel の初期設定（重要）

1. Excel を起動

2. ファイル → オプション

3. セキュリティ センター → セキュリティ センターの設定

4. マクロの設定

- 「警告を表示してすべてのマクロを無効にする」

5. VBA プロジェクト モデルへのアクセス

- ✅ 有効にする

  ※ これが無効だと PowerShell から VBA を操作できません

## 4. Excel の生成（src → excel）

```powershell
cd tools
.\import_vba.ps1
```

- `excel/ezGrepView.xlsm` にソースがインポートされます

## 5. 開発フロー

1. `excel/ezGrepView.xlsm` を `dev` にコピーする。

```powershell
cd ..\dev
Copy-Item -Path ..\excel\ezGrepView.xlsm -Destination .
```

2. VBA を編集

- dev/ezGrepView.xlsm を開く

- VBA エディタで修正

- Excel を保存

3. ソースをエクスポート（dev → src）

```powershell
.\export_vba.ps1
```

- src/*.bas / *.cls / *.dcm が更新されます

- 同時に VBAを含まない配布用 Excel が excel/ に生成されます

---

## 🛠 使い方

### ① サクラエディタで grep 実行

サクラエディタで検索を実行し、  **結果をファイルとして保存** します。

### ② Main シートで設定

以下を入力・設定します。

- grep 結果ファイルのパス
- 文字コード（UTF-8 / Shift_JIS）
- ハイライト有無
- コメント／バイナリ／文字化け行のマーク文字

### ③ 取り込み実行
「取り込み」ボタンを押すと…

- grep ヘッダ解析
- 結果専用 Excel ファイル生成
- Result シートへ一覧出力

が自動で行われます。

### ④ Result シート操作

- 行選択でソース全文を表示
- 一致箇所・コメントを自動ハイライト
- `Ctrl + J` でサクラエディタを起動＆該当行へジャンプ

---

## 📑 シート構成

| シート名 | 内容 |
|--------|------|
| Main | 入力画面（grep結果・設定） |
| Result | 検索結果一覧・ソース表示 |
| Comment | 拡張子別コメント定義 |

---

## 📜 ライセンス

**MIT License**
