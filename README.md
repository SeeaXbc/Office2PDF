# Office to PDF Converter

Word/ExcelファイルをPDFに一括変換するツールです。

## 必要な環境

- Windows 7以降
- Microsoft Office（Word/Excel）がインストールされていること
- PowerShell（Windows標準搭載）

## ファイル構成

- Office2PDF.bat  : 実行用バッチファイル
- Office2PDF.ps1  : PowerShellスクリプト本体
- README.txt      : このファイル

## 使用方法

### 方法1: ドラッグ&ドロップ（推奨）
1. 変換したいWord/Excelファイルを選択
2. 選択したファイルを「Office2PDF.bat」にドラッグ&ドロップ
3. 変換が自動的に開始されます

### 方法2: 送る（Send to）メニューに追加
1. Windows + R キーを押して「shell:sendto」を入力
2. 開いたフォルダに「Office2PDF.bat」のショートカットを配置
3. ファイルを右クリック → 送る → Office2PDF.bat

## 対応ファイル形式

- Word: .doc, .docx
- Excel: .xls, .xlsx

## 変換仕様

- **Word**: 全ページをPDFに変換
- **Excel**: 全シートを1つのPDFに変換
- 印刷範囲、ページ設定は元ファイルの設定を使用
- PDFは元ファイルのフォルダ内に作成される「PDF」フォルダに保存されます
- ファイル名は元のファイル名.pdf となります
- 同名のPDFファイルが存在する場合は上書きされます

### フォルダ構造の例
```
C:\Documents\
├── 報告書.docx
├── 売上表.xlsx
└── PDF\              ← 自動作成
    ├── 報告書.pdf
    └── 売上表.pdf
```

## 実行時の確認

変換実行前に以下の内容が表示され、確認を求められます：
- 変換対象ファイルの一覧
- PDFの保存先フォルダ
- 既存ファイルの上書き警告
- Y/Nで実行確認

## 注意事項

- 大量のファイルを一度に変換する場合は時間がかかります
- 変換中はWord/Excelが自動的に起動しますが、操作しないでください
- PDFフォルダは自動的に作成されます
- 既存のPDFファイルは確認後に上書きされます
- パスワード保護されたファイルは変換できません
- Nを選択した場合、変換はキャンセルされます

## トラブルシューティング

### 「実行できません」エラーが出る場合
PowerShellの実行ポリシーが原因の可能性があります。
管理者権限でPowerShellを開き、以下を実行してください：
```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Office関連のエラーが出る場合
- Officeが正しくインストールされているか確認
- Officeのライセンスが有効か確認
- 他のプログラムでファイルを開いていないか確認
