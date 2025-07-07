# Office2PDF.ps1
# Word/ExcelファイルをPDFに変換するPowerShellスクリプト

param(
    [Parameter(Mandatory=$false)]
    [string[]]$Files = @()
)

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

foreach ($file in $Files) {
    if (Test-Path $file) {
        if ($file -match '\.(doc|docx|xls|xlsx)$') {
            $validFiles += $file
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

Write-Progress-Message "変換対象: $($validFiles.Count) ファイル" "Info"
Write-Host ""

# Office COMオブジェクト
$word = $null
$excel = $null
$successCount = 0
$errorCount = 0

try {
    # 各ファイルを処理
    for ($i = 0; $i -lt $validFiles.Count; $i++) {
        $file = $validFiles[$i]
        $fileName = [System.IO.Path]::GetFileName($file)
        
        # 進捗表示
        $progress = [math]::Round(($i / $validFiles.Count) * 100)
        Write-Progress -Activity "PDF変換中" -Status "処理中: $fileName" -PercentComplete $progress
        
        try {
            $filePath = (Resolve-Path $file).Path
            $pdfPath = [System.IO.Path]::ChangeExtension($filePath, ".pdf")
            
            Write-Progress-Message "変換開始: $fileName" "Info"
            
            # Word文書の場合
            if ($file -match '\.(doc|docx)$') {
                if (-not $word) {
                    Write-Progress-Message "Microsoft Wordを起動しています..." "Info"
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0  # wdAlertsNone
                }
                
                $doc = $word.Documents.Open($filePath, $false, $true)  # ReadOnly = true
                $doc.SaveAs2($pdfPath, 17)  # 17 = wdFormatPDF
                $doc.Close($false)
                
                Write-Progress-Message "変換完了: $([System.IO.Path]::GetFileName($pdfPath))" "Success"
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
                
                $workbook = $excel.Workbooks.Open($filePath, $false, $true)  # ReadOnly = true
                
                # 全シートを1つのPDFに出力
                $workbook.ExportAsFixedFormat(
                    0,          # Type: 0 = xlTypePDF
                    $pdfPath,   # Filename
                    0,          # Quality: 0 = xlQualityStandard
                    $true,      # IncludeDocProperties
                    $false,     # IgnorePrintAreas - falseで印刷範囲を使用
                    $null,      # From
                    $null       # To
                )
                
                $workbook.Close($false)
                
                Write-Progress-Message "変換完了: $([System.IO.Path]::GetFileName($pdfPath))" "Success"
                $successCount++
            }
            
        }
        catch {
            $errorCount++
            Write-Progress-Message "変換失敗: $fileName" "Error"
            Write-Progress-Message "エラー詳細: $($_.Exception.Message)" "Error"
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
    Write-Progress-Message "PDFファイルは元のファイルと同じフォルダに保存されました。" "Success"
}