# Office2PDF.ps1
# Word/ExcelファイルをPDFに変換するPowerShellスクリプト

param(
    [Parameter(Mandatory=$false, ValueFromRemainingArguments=$true)]
    [string[]]$Files = @()
)

# 引数のデバッグ情報（問題がある場合はコメントを外して確認）
# Write-Host "受け取った引数数: $($Files.Count)" -ForegroundColor Yellow
# foreach ($f in $Files) {
#     Write-Host "引数: $f" -ForegroundColor Yellow
# }

# 進捗表示用の関数
function Write-Progress-Message {
    param(
        [string]$Message,
        [string]$Status = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    switch ($Status) {
        "Success" { 
            Write-Host "[$timestamp] " -NoNewline
            Write-Host "? $Message" -ForegroundColor Green 
        }
        "Error" { 
            Write-Host "[$timestamp] " -NoNewline
            Write-Host "? $Message" -ForegroundColor Red 
        }
        "Warning" { 
            Write-Host "[$timestamp] " -NoNewline
            Write-Host "? $Message" -ForegroundColor Yellow 
        }
        default { 
            Write-Host "[$timestamp] ? $Message" -ForegroundColor Cyan 
        }
    }
}

# メイン処理
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   Office to PDF Converter" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ファイルが指定されていない場合
if ($Files.Count -eq 0) {
    Write-Progress-Message "変換するファイルが指定されていません。" "Error"
    Write-Progress-Message "使用方法: ファイルをバッチファイルにドラッグ&ドロップしてください。" "Info"
    exit 1
}

# 変換対象ファイルの確認
$validFiles = @()
$invalidFiles = @()
$fileInfoList = @()

foreach ($file in $Files) {
    if (Test-Path $file) {
        if ($file -match '\.(doc|docx|xls|xlsx)$') {
            $fullPath = (Resolve-Path $file).Path
            $directory = [System.IO.Path]::GetDirectoryName($fullPath)
            $fileName = [System.IO.Path]::GetFileName($fullPath)
            $pdfFolder = Join-Path $directory "PDF"
            $pdfFileName = [System.IO.Path]::ChangeExtension($fileName, ".pdf")
            $pdfPath = Join-Path $pdfFolder $pdfFileName
            
            $validFiles += $file
            $fileInfoList += @{
                SourcePath = $fullPath
                SourceFileName = $fileName
                PdfFolder = $pdfFolder
                PdfPath = $pdfPath
                PdfFileName = $pdfFileName
                Directory = $directory
            }
        } else {
            $invalidFiles += $file
        }
    } else {
        Write-Progress-Message "ファイルが見つかりません: $file" "Error"
    }
}

if ($invalidFiles.Count -gt 0) {
    Write-Progress-Message "以下のファイルは対応していない形式です:" "Warning"
    foreach ($file in $invalidFiles) {
        Write-Host "  - $([System.IO.Path]::GetFileName($file))" -ForegroundColor Yellow
    }
    Write-Host ""
}

if ($validFiles.Count -eq 0) {
    Write-Progress-Message "変換可能なファイルがありません。" "Error"
    exit 1
}

# 変換対象と保存先の表示
Write-Host "変換対象ファイル:" -ForegroundColor Cyan
Write-Host ""

# フォルダごとにグループ化して表示
$groupedFiles = $fileInfoList | Group-Object -Property Directory

foreach ($group in $groupedFiles) {
    Write-Host "フォルダ: $($group.Name)" -ForegroundColor White
    Write-Host "保存先 : $($group.Group[0].PdfFolder)" -ForegroundColor Gray
    Write-Host ""
    
    foreach ($fileInfo in $group.Group) {
        Write-Host "  ・$($fileInfo.SourceFileName)" -ForegroundColor White
        Write-Host "    → $($fileInfo.PdfFileName)" -ForegroundColor Gray
        
        # 既存ファイルのチェック
        if (Test-Path $fileInfo.PdfPath) {
            Write-Host "    (既存のファイルを上書きします)" -ForegroundColor Yellow
        }
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor DarkGray
Write-Host "合計: $($validFiles.Count) ファイル" -ForegroundColor Cyan
Write-Host ""

# 確認プロンプト
Write-Host "上記のファイルをPDFに変換しますか？" -ForegroundColor Yellow
$confirmation = Read-Host "実行する場合は 'Y' を、キャンセルする場合は 'N' を入力してください (Y/N)"

if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
    Write-Progress-Message "処理をキャンセルしました。" "Warning"
    exit 0
}

Write-Host ""
Write-Progress-Message "変換を開始します..." "Info"
Write-Host ""

# Office COMオブジェクト
$word = $null
$excel = $null
$successCount = 0
$errorCount = 0

try {
    # PDFフォルダの作成
    $createdFolders = @()
    foreach ($fileInfo in $fileInfoList) {
        if (-not (Test-Path $fileInfo.PdfFolder)) {
            New-Item -ItemType Directory -Path $fileInfo.PdfFolder -Force | Out-Null
            if ($fileInfo.PdfFolder -notin $createdFolders) {
                $createdFolders += $fileInfo.PdfFolder
                Write-Progress-Message "フォルダを作成しました: $($fileInfo.PdfFolder)" "Success"
            }
        }
    }
    
    if ($createdFolders.Count -gt 0) {
        Write-Host ""
    }
    
    # 各ファイルを処理
    for ($i = 0; $i -lt $fileInfoList.Count; $i++) {
        $fileInfo = $fileInfoList[$i]
        $file = $validFiles[$i]
        
        # 進捗表示
        $progress = [math]::Round((($i + 1) / $fileInfoList.Count) * 100)
        Write-Progress -Activity "PDF変換中" -Status "処理中: $($fileInfo.SourceFileName)" -PercentComplete $progress
        
        try {
            Write-Progress-Message "変換開始: $($fileInfo.SourceFileName)" "Info"
            
            # Word文書の場合
            if ($file -match '\.(doc|docx)$') {
                if (-not $word) {
                    Write-Progress-Message "Microsoft Wordを起動しています..." "Info"
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0  # wdAlertsNone
                }
                
                # PDFフォルダの存在を再確認
                if (-not (Test-Path $fileInfo.PdfFolder)) {
                    New-Item -ItemType Directory -Path $fileInfo.PdfFolder -Force | Out-Null
                    Start-Sleep -Milliseconds 500  # フォルダ作成待機
                }
                
                $doc = $word.Documents.Open($fileInfo.SourcePath, $false, $true)  # ReadOnly = true
                $doc.SaveAs2($fileInfo.PdfPath, 17)  # 17 = wdFormatPDF
                $doc.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                
                Write-Progress-Message "変換完了: $($fileInfo.PdfFileName)" "Success"
                $successCount++
            }
            
            # Excel文書の場合
            elseif ($file -match '\.(xls|xlsx)$') {
                if (-not $excel) {
                    Write-Progress-Message "Microsoft Excelを起動しています..." "Info"
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                }
                
                # PDFフォルダの存在を再確認
                if (-not (Test-Path $fileInfo.PdfFolder)) {
                    New-Item -ItemType Directory -Path $fileInfo.PdfFolder -Force | Out-Null
                    Start-Sleep -Milliseconds 500  # フォルダ作成待機
                }
                
                # Excelファイルを開く
                $workbook = $excel.Workbooks.Open($fileInfo.SourcePath, $false, $true)
                
                # 全シートを1つのPDFに出力（Workbookから直接エクスポート）
                $workbook.ExportAsFixedFormat(
                    0,                      # Type: 0 = xlTypePDF
                    $fileInfo.PdfPath       # Filename
                )
                
                # ワークブックを閉じる
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                
                Write-Progress-Message "変換完了: $($fileInfo.PdfFileName)" "Success"
                $successCount++
            }
            
        }
        catch {
            $errorCount++
            Write-Progress-Message "変換失敗: $($fileInfo.SourceFileName)" "Error"
            Write-Progress-Message "エラー詳細: $($_.Exception.Message)" "Error"
            
            # Excelの場合は追加の情報を表示
            if ($file -match '\.(xls|xlsx)$') {
                Write-Progress-Message "Excelファイルの変換でエラーが発生しました。" "Error"
                Write-Progress-Message "印刷設定やシートの内容を確認してください。" "Info"
            }
        }
    }
    
    Write-Progress -Activity "PDF変換中" -Completed
    
}
catch {
    Write-Progress-Message "予期しないエラーが発生しました" "Error"
    Write-Progress-Message "エラー詳細: $($_.Exception.Message)" "Error"
}
finally {
    # クリーンアップ
    if ($word) {
        Write-Progress-Message "Wordを終了しています..." "Info"
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($excel) {
        Write-Progress-Message "Excelを終了しています..." "Info"
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# 結果サマリー
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   変換結果" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  成功: $successCount ファイル" -ForegroundColor Green
Write-Host "  失敗: $errorCount ファイル" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Gray" })
Write-Host ""

if ($successCount -gt 0) {
    Write-Progress-Message "PDFファイルは各フォルダ内の'PDF'フォルダに保存されました。" "Success"
}