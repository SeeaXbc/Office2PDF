# Office2PDF.ps1
# Word/Excel�t�@�C����PDF�ɕϊ�����PowerShell�X�N���v�g

param(
    [Parameter(Mandatory=$false, ValueFromRemainingArguments=$true)]
    [string[]]$Files = @()
)

# �����̃f�o�b�O���i��肪����ꍇ�̓R�����g���O���Ċm�F�j
# Write-Host "�󂯎����������: $($Files.Count)" -ForegroundColor Yellow
# foreach ($f in $Files) {
#     Write-Host "����: $f" -ForegroundColor Yellow
# }

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

# �ϊ��Ώۂƕۑ���̕\��
Write-Host "�ϊ��Ώۃt�@�C��:" -ForegroundColor Cyan
Write-Host ""

# �t�H���_���ƂɃO���[�v�����ĕ\��
$groupedFiles = $fileInfoList | Group-Object -Property Directory

foreach ($group in $groupedFiles) {
    Write-Host "�t�H���_: $($group.Name)" -ForegroundColor White
    Write-Host "�ۑ��� : $($group.Group[0].PdfFolder)" -ForegroundColor Gray
    Write-Host ""
    
    foreach ($fileInfo in $group.Group) {
        Write-Host "  �E$($fileInfo.SourceFileName)" -ForegroundColor White
        Write-Host "    �� $($fileInfo.PdfFileName)" -ForegroundColor Gray
        
        # �����t�@�C���̃`�F�b�N
        if (Test-Path $fileInfo.PdfPath) {
            Write-Host "    (�����̃t�@�C�����㏑�����܂�)" -ForegroundColor Yellow
        }
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor DarkGray
Write-Host "���v: $($validFiles.Count) �t�@�C��" -ForegroundColor Cyan
Write-Host ""

# �m�F�v�����v�g
Write-Host "��L�̃t�@�C����PDF�ɕϊ����܂����H" -ForegroundColor Yellow
$confirmation = Read-Host "���s����ꍇ�� 'Y' ���A�L�����Z������ꍇ�� 'N' ����͂��Ă������� (Y/N)"

if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
    Write-Progress-Message "�������L�����Z�����܂����B" "Warning"
    exit 0
}

Write-Host ""
Write-Progress-Message "�ϊ����J�n���܂�..." "Info"
Write-Host ""

# Office COM�I�u�W�F�N�g
$word = $null
$excel = $null
$successCount = 0
$errorCount = 0

try {
    # PDF�t�H���_�̍쐬
    $createdFolders = @()
    foreach ($fileInfo in $fileInfoList) {
        if (-not (Test-Path $fileInfo.PdfFolder)) {
            New-Item -ItemType Directory -Path $fileInfo.PdfFolder -Force | Out-Null
            if ($fileInfo.PdfFolder -notin $createdFolders) {
                $createdFolders += $fileInfo.PdfFolder
                Write-Progress-Message "�t�H���_���쐬���܂���: $($fileInfo.PdfFolder)" "Success"
            }
        }
    }
    
    if ($createdFolders.Count -gt 0) {
        Write-Host ""
    }
    
    # �e�t�@�C��������
    for ($i = 0; $i -lt $fileInfoList.Count; $i++) {
        $fileInfo = $fileInfoList[$i]
        $file = $validFiles[$i]
        
        # �i���\��
        $progress = [math]::Round((($i + 1) / $fileInfoList.Count) * 100)
        Write-Progress -Activity "PDF�ϊ���" -Status "������: $($fileInfo.SourceFileName)" -PercentComplete $progress
        
        try {
            Write-Progress-Message "�ϊ��J�n: $($fileInfo.SourceFileName)" "Info"
            
            # Word�����̏ꍇ
            if ($file -match '\.(doc|docx)$') {
                if (-not $word) {
                    Write-Progress-Message "Microsoft Word���N�����Ă��܂�..." "Info"
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0  # wdAlertsNone
                }
                
                # PDF�t�H���_�̑��݂��Ċm�F
                if (-not (Test-Path $fileInfo.PdfFolder)) {
                    New-Item -ItemType Directory -Path $fileInfo.PdfFolder -Force | Out-Null
                    Start-Sleep -Milliseconds 500  # �t�H���_�쐬�ҋ@
                }
                
                $doc = $word.Documents.Open($fileInfo.SourcePath, $false, $true)  # ReadOnly = true
                $doc.SaveAs2($fileInfo.PdfPath, 17)  # 17 = wdFormatPDF
                $doc.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                
                Write-Progress-Message "�ϊ�����: $($fileInfo.PdfFileName)" "Success"
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
                
                # PDF�t�H���_�̑��݂��Ċm�F
                if (-not (Test-Path $fileInfo.PdfFolder)) {
                    New-Item -ItemType Directory -Path $fileInfo.PdfFolder -Force | Out-Null
                    Start-Sleep -Milliseconds 500  # �t�H���_�쐬�ҋ@
                }
                
                # Excel�t�@�C�����J��
                $workbook = $excel.Workbooks.Open($fileInfo.SourcePath, $false, $true)
                
                # �S�V�[�g��1��PDF�ɏo�́iWorkbook���璼�ڃG�N�X�|�[�g�j
                $workbook.ExportAsFixedFormat(
                    0,                      # Type: 0 = xlTypePDF
                    $fileInfo.PdfPath       # Filename
                )
                
                # ���[�N�u�b�N�����
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                
                Write-Progress-Message "�ϊ�����: $($fileInfo.PdfFileName)" "Success"
                $successCount++
            }
            
        }
        catch {
            $errorCount++
            Write-Progress-Message "�ϊ����s: $($fileInfo.SourceFileName)" "Error"
            Write-Progress-Message "�G���[�ڍ�: $($_.Exception.Message)" "Error"
            
            # Excel�̏ꍇ�͒ǉ��̏���\��
            if ($file -match '\.(xls|xlsx)$') {
                Write-Progress-Message "Excel�t�@�C���̕ϊ��ŃG���[���������܂����B" "Error"
                Write-Progress-Message "����ݒ��V�[�g�̓��e���m�F���Ă��������B" "Info"
            }
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
    Write-Progress-Message "PDF�t�@�C���͊e�t�H���_����'PDF'�t�H���_�ɕۑ�����܂����B" "Success"
}