# Office2PDF v5.x（Office一括PDF変換・加工ツール）

![2](https://github.com/user-attachments/assets/8914e645-a319-4291-a62c-9eb61e0c6039)



**事務作業の「面倒」をまとめて自動化する、実務特化型のOffice一括PDF変換・加工ツールです。**  
Officeファイルや画像をPDF化するだけでなく、透かし・ページ番号・分割/結合・命名ルールまで数クリックでまとめて処理できます。

> **完全オフライン変換**：Microsoft Office または WPS Office（COM）を利用します。
> Office文書には対応するCOMコンポーネントが必要です。PDF・画像の処理にはOffice/WPSは不要です。

変換・キャンセルボタンと進捗は下部に固定され、設定はスクロール、ログは折りたたみ可能です。エンジンは「自動（Office優先）」「Microsoft Office」「WPS Office」から選択でき、設定・プリセットに保存されます。自動モードはOfficeコンポーネントの起動失敗時のみWPSを試します。WPSの対応はインストール済みCOM機能に依存し、`.wps/.et/.dps`形式は追加していません。

---

## ⚠️ 保守・メンテナンスについて

**【重要】作者はプロの開発者ではありません。**  
本ツールは個人の業務効率化目的で作成しており、複雑な不具合対応や高度な保守が難しい場合があります。

- **エンジニアの方へ：** 改善案、バグ修正、機能追加などのPull Requestを歓迎します
- **ユーザーの方へ：** 本ソフトは **現状有姿（as-is）** で提供されます

---

## 🌟 主な機能（v5.x）

### 一括PDF変換
- 対応形式：
  - **Word**：`.doc`, `.docx`
  - **Excel**：`.xls`, `.xlsx`, `.xlsm`
  - **PowerPoint**：`.ppt`, `.pptx`
  - **画像**：`.jpg`, `.jpeg`, `.png`
  - **PDF**：再保存/正規化＋加工

### 分割・結合が柔軟
- **全結合**：すべてを1つのPDFへ結合
- **ページごと分割**：Word / PowerPoint / PDF / Excel
- **シートごと分割**：Excel

### 範囲指定（ページ指定 / シート指定）
- リスト項目を**ダブルクリック**して指定できます
  - Word / PowerPoint / PDF：ページ範囲  
    例）`1-3,5,8-` / `2` / `-3` / `1-`
  - Excel：対象シート（複数選択）

### 透かし・ページ番号・プレビュー
- 透かし最大 **2段**（位置/透明度/フォント/サイズ/色）
- ページ番号：書式（例 `- {n} / {total} -`）で挿入  
  - ※（現行ビルド）ページ番号は **黒・10.5pt固定**
- **プレビュー**：選択ファイルの **1ページ目だけ** に透かし/番号を当てて確認できます

### 命名ルール・プリセット
- 命名テンプレートにタグを使って自動命名（後述）
- 設定プリセットの保存/読込（JSON）

### セキュリティ・最適化
- PDFパスワードの一括設定
- メタデータ削除
- PDF軽量化（pypdfの対応状況により有効）

### 使い勝手
- ドラッグ＆ドロップ登録
- 並び替え、削除、リスト全消去
- 出力先：元フォルダ or 指定フォルダ
- 完了後：PDFを開く / フォルダを開く / リストクリア

---

## 🏷️ 利用可能なタグ一覧

タグは **命名ルール** と **透かし文字** に使えます。処理時に自動置換されます。

- `{name}`：元のファイル名（拡張子なし）
- `{sheet}`：Excelのシート名（Excel以外は空文字）
- `{parent}`：親フォルダ名
- `{seq}`：全体の通し番号
- `{fseq}`：入力ファイルごとの連番
- `{pseq}`：ページ連番（ページ分割/ページカウント時）
- `{total}`：処理する全ファイル数
- `{ptotal}`：そのファイルの総ページ数
- `{username}`：Windowsログインユーザー名
- `{rand}`：4桁ランダム数字（例 `0427`）

### 日付・時刻タグ
- `{date:yyyy-mm-dd}` → `2026-02-06`
- `{date:yyyy/mm/dd HH:MM}` → `2026/02/06 16:54`

利用可能プレースホルダ：
- `yyyy` 年 / `mm` 月 / `dd` 日
- `HH` 時 / `MM` 分 / `SS` 秒

---

## 🚀 使い方

### 動作環境
- Windows 10 / 11（64bit）
- **Microsoft Office または COM 対応の WPS Office の必要なコンポーネントがインストールされていること**

### 基本手順
1. アプリを起動
2. ファイルをドラッグ＆ドロップで登録
3. （必要なら）ファイルをダブルクリックしてページ範囲/シートを指定
4. 透かし/ページ番号、分割/結合、命名ルール、出力先を設定
5. **PDF変換開始** をクリック

---

## 📦 導入方法

1. GitHubの **Releases** から最新版 `Office2PDF.zip` をダウンロード
2. 解凍して `Office2PDF.exe` を実行（インストール不要）

> まだReleasesが無い場合：GitHubの「Releases」→「Draft a new release」で作成できます。

---

## 🛠️ ソースコードから実行・ビルド

### 1) 実行環境
- Python **3.9以上**（Windows）
- Microsoft Office または COM 対応 WPS Office インストール済み

### 2) 依存ライブラリのインストール
```bash
pip install pywin32 pypdf Pillow reportlab tkinterdnd2
````

### 3) 起動

```bash
python Office2PDF.py
```

### 4) exe化（PyInstaller）

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 --name Office2PDF Office2PDF.py
```

ビルド後、`dist/Office2PDF.exe` が生成されます。

> 補足：PowerPointは環境によって非表示起動が不安定な場合があります。  
> 変換が失敗する場合は「Visible=False」を強制しない方針が安全です。

***

## ☕ カンパのお願い

もし本ソフトがお役に立ちましたら、開発継続の支援としてカンパをいただけると励みになります。

*   Amazonギフトカード（Eメールタイプ）  
    <https://www.amazon.co.jp/Amazon-eGift-Card-Flower-Bouquet-Animated/dp/B004N3APDM/>
*   受取人メール：`the.desert.of.the.real@gmail.com`

***

## ⚖️ ライセンス

MIT License — ./LICENSE
