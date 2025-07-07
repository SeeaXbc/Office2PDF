@echo off
REM Office2PDF.bat - Word/ExcelファイルをPDFに変換

REM ウィンドウタイトルを設定
title Office to PDF Converter

REM 文字コードをUTF-8に設定（日本語表示用）
chcp 65001 > nul

REM バッチファイルのディレクトリに移動
cd /d "%~dp0"

REM PowerShellスクリプトの存在確認
if not exist "Office2PDF.ps1" (
    echo エラー: Office2PDF.ps1 が見つかりません。
    echo バッチファイルと同じフォルダに配置してください。
    pause
    exit /b 1
)

REM ファイルが指定されていない場合の処理
if "%~1"=="" (
    echo.
    echo ========================================
    echo    Office to PDF Converter
    echo ========================================
    echo.
    echo 使用方法:
    echo   1. 変換したいWord/Excelファイルを選択
    echo   2. このバッチファイルにドラッグ＆ドロップ
    echo.
    echo 対応形式:
    echo   - Word: .doc, .docx
    echo   - Excel: .xls, .xlsx
    echo.
    echo 保存先:
    echo   元ファイルのフォルダ内の「PDF」フォルダに保存されます
    echo.
    pause
    exit /b 0
)

REM PowerShellスクリプトを実行
REM -ExecutionPolicy Bypass: スクリプト実行ポリシーを一時的に回避
REM -NoProfile: プロファイルを読み込まない（高速化）
REM -File: スクリプトファイルを指定
REM %*: すべての引数（ドロップされたファイル）を渡す

echo 変換対象ファイルを確認しています...
echo.

powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0Office2PDF.ps1" %*

REM 実行結果を確認できるように一時停止
echo.
echo 処理が完了しました。
pause