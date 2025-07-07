# Office2PDF.ps1
# Word/Excel�t�@�C����PDF�ɕϊ�����PowerShell�X�N���v�g

param(
    [Parameter(Mandatory=$false)]
    [string[]]$Files = @()
)

# �i���\���p�̊֐�
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

# ���C������
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   Office to PDF Converter" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# �t�@�C�����w�肳��Ă��Ȃ��ꍇ
if ($Files.Count -eq 0) {
    Write-Progress-Message "�ϊ�����t�@�C�����w�肳��Ă��܂���B" "Error"
    Write-Progress-Message "�g�p���@: �t�@�C�����o�b�`�t�@�C���Ƀh���b�O&�h���b�v���Ă��������B" "Info"
    exit 1
}

# �ϊ��Ώۃt�@�C���̊m�F
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
        Write-Progress-Message "�t�@�C����������܂���: $file" "Error"
    }
}

if ($invalidFiles.Count -gt 0) {
    Write-Progress-Message "�ȉ��̃t�@�C���͑Ή����Ă��Ȃ��`���ł�:" "Warning"
    foreach ($file in $invalidFiles) {
        Write-Host "  - $([System.IO.Path]::GetFileName($file))" -ForegroundColor Yellow
    }
    Write-Host ""
}

if ($validFiles.Count -eq 0) {
    Write-Progress-Message "�ϊ��\�ȃt�@�C��������܂���B" "Error"
    exit 1
}

Write-Progress-Message "�ϊ��Ώ�: $($validFiles.Count) �t�@�C��" "Info"
Write-Host ""

# Office COM�I�u�W�F�N�g
$word = $null
$excel = $null
$successCount = 0
$errorCount = 0

try {
    # �e�t�@�C��������
    for ($i = 0; $i -lt $validFiles.Count; $i++) {
        $file = $validFiles[$i]
        $fileName = [System.IO.Path]::GetFileName($file)
        
        # �i���\��
        $progress = [math]::Round(($i / $validFiles.Count) * 100)
        Write-Progress -Activity "PDF�ϊ���" -Status "������: $fileName" -PercentComplete $progress
        
        try {
            $filePath = (Resolve-Path $file).Path
            $pdfPath = [System.IO.Path]::ChangeExtension($filePath, ".pdf")
            
            Write-Progress-Message "�ϊ��J�n: $fileName" "Info"
            
            # Word�����̏ꍇ
            if ($file -match '\.(doc|docx)$') {
                if (-not $word) {
                    Write-Progress-Message "Microsoft Word���N�����Ă��܂�..." "Info"
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0  # wdAlertsNone
                }
                
                $doc = $word.Documents.Open($filePath, $false, $true)  # ReadOnly = true
                $doc.SaveAs2($pdfPath, 17)  # 17 = wdFormatPDF
                $doc.Close($false)
                
                Write-Progress-Message "�ϊ�����: $([System.IO.Path]::GetFileName($pdfPath))" "Success"
                $successCount++
            }
            
            # Excel�����̏ꍇ
            elseif ($file -match '\.(xls|xlsx)$') {
                if (-not $excel) {
                    Write-Progress-Message "Microsoft Excel���N�����Ă��܂�..." "Info"
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                }
                
                $workbook = $excel.Workbooks.Open($filePath, $false, $true)  # ReadOnly = true
                
                # �S�V�[�g��1��PDF�ɏo��
                $workbook.ExportAsFixedFormat(
                    0,          # Type: 0 = xlTypePDF
                    $pdfPath,   # Filename
                    0,          # Quality: 0 = xlQualityStandard
                    $true,      # IncludeDocProperties
                    $false,     # IgnorePrintAreas - false�ň���͈͂��g�p
                    $null,      # From
                    $null       # To
                )
                
                $workbook.Close($false)
                
                Write-Progress-Message "�ϊ�����: $([System.IO.Path]::GetFileName($pdfPath))" "Success"
                $successCount++
            }
            
        }
        catch {
            $errorCount++
            Write-Progress-Message "�ϊ����s: $fileName" "Error"
            Write-Progress-Message "�G���[�ڍ�: $($_.Exception.Message)" "Error"
        }
    }
    
    Write-Progress -Activity "PDF�ϊ���" -Completed
    
}
catch {
    Write-Progress-Message "�\�����Ȃ��G���[���������܂���" "Error"
    Write-Progress-Message "�G���[�ڍ�: $($_.Exception.Message)" "Error"
}
finally {
    # �N���[���A�b�v
    if ($word) {
        Write-Progress-Message "Word���I�����Ă��܂�..." "Info"
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($excel) {
        Write-Progress-Message "Excel���I�����Ă��܂�..." "Info"
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ���ʃT�}���[
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   �ϊ�����" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  ����: $successCount �t�@�C��" -ForegroundColor Green
Write-Host "  ���s: $errorCount �t�@�C��" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Gray" })
Write-Host ""

if ($successCount -gt 0) {
    Write-Progress-Message "PDF�t�@�C���͌��̃t�@�C���Ɠ����t�H���_�ɕۑ�����܂����B" "Success"
}