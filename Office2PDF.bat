@echo off
REM Office2PDF.bat - Word/Excel�t�@�C����PDF�ɕϊ�

REM �E�B���h�E�^�C�g����ݒ�
title Office to PDF Converter

REM �����R�[�h��UTF-8�ɐݒ�i���{��\���p�j
chcp 65001 > nul

REM �o�b�`�t�@�C���̃f�B���N�g���Ɉړ�
cd /d "%~dp0"

REM PowerShell�X�N���v�g�̑��݊m�F
if not exist "Office2PDF.ps1" (
    echo �G���[: Office2PDF.ps1 ��������܂���B
    echo �o�b�`�t�@�C���Ɠ����t�H���_�ɔz�u���Ă��������B
    pause
    exit /b 1
)

REM �t�@�C�����w�肳��Ă��Ȃ��ꍇ�̏���
if "%~1"=="" (
    echo.
    echo ========================================
    echo    Office to PDF Converter
    echo ========================================
    echo.
    echo �g�p���@:
    echo   1. �ϊ�������Word/Excel�t�@�C����I��
    echo   2. ���̃o�b�`�t�@�C���Ƀh���b�O���h���b�v
    echo.
    echo �Ή��`��:
    echo   - Word: .doc, .docx
    echo   - Excel: .xls, .xlsx
    echo.
    echo �ۑ���:
    echo   ���t�@�C���̃t�H���_���́uPDF�v�t�H���_�ɕۑ�����܂�
    echo.
    pause
    exit /b 0
)

REM PowerShell�X�N���v�g�����s
REM -ExecutionPolicy Bypass: �X�N���v�g���s�|���V�[���ꎞ�I�ɉ��
REM -NoProfile: �v���t�@�C����ǂݍ��܂Ȃ��i�������j
REM -File: �X�N���v�g�t�@�C�����w��
REM %*: ���ׂĂ̈����i�h���b�v���ꂽ�t�@�C���j��n��

echo �ϊ��Ώۃt�@�C�����m�F���Ă��܂�...
echo.

powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0Office2PDF.ps1" %*

REM ���s���ʂ��m�F�ł���悤�Ɉꎞ��~
echo.
echo �������������܂����B
pause