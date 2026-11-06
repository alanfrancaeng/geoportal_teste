@echo off
title BATTRE - GEOportal & BI-Obras
color 0B
chcp 65001 >nul
cls

echo ======================================================================
echo       BATTRE - GEOPORTAL & PAINEL EXECUTIVO PMO DE OBRAS
echo ======================================================================
echo.

:: Garante que o diretorio de trabalho e sempre a pasta deste .bat
cd /d "%~dp0"

echo  [1/3] Pasta do projeto: %~dp0
echo.

:: ─── Detectar Python ────────────────────────────────────────────────────
set "PY_EXE="

:: Tenta QGIS 3.40 (apps\Python312)
if exist "C:\Program Files\QGIS 3.40.12\apps\Python312\python.exe" (
    set "PY_EXE=C:\Program Files\QGIS 3.40.12\apps\Python312\python.exe"
    goto :found_py
)
:: Tenta QGIS 3.40 (bin)
if exist "C:\Program Files\QGIS 3.40.12\bin\python.exe" (
    set "PY_EXE=C:\Program Files\QGIS 3.40.12\bin\python.exe"
    goto :found_py
)
:: Tenta QGIS 3.38
if exist "C:\Program Files\QGIS 3.38\apps\Python312\python.exe" (
    set "PY_EXE=C:\Program Files\QGIS 3.38\apps\Python312\python.exe"
    goto :found_py
)
:: Tenta Python padrao Windows
if exist "C:\Program Files\Python312\python.exe" (
    set "PY_EXE=C:\Program Files\Python312\python.exe"
    goto :found_py
)
if exist "C:\Program Files\Python311\python.exe" (
    set "PY_EXE=C:\Program Files\Python311\python.exe"
    goto :found_py
)
if exist "C:\Python312\python.exe" (
    set "PY_EXE=C:\Python312\python.exe"
    goto :found_py
)
:: Tenta encontrar python no PATH do sistema
for /f "tokens=*" %%i in ('where python 2^>nul') do (
    set "PY_EXE=%%i"
    goto :found_py
)

:found_py
:: ─── Executar servidor Python ────────────────────────────────────────────
if defined PY_EXE (
    if exist "%PY_EXE%" (
        echo  [OK] Python localizado: %PY_EXE%
        echo  [2/3] Sincronizando dados e iniciando servidor local...
        echo.
        echo  ----------------------------------------------------------------
        echo   SISTEMA PRONTO! O navegador abrira automaticamente.
        echo   Mantenha esta janela ABERTA enquanto usar o GEOportal.
        echo   Para encerrar: feche esta janela ou pressione Ctrl+C
        echo  ----------------------------------------------------------------
        echo.
        "%PY_EXE%" "%~dp0server.py"
        goto :fim
    )
)

:: ─── Fallback: sem Python — abre via servidor HTTP do PowerShell ──────────
echo  [AVISO] Python nao encontrado. Iniciando via PowerShell...
echo.

:: Mata qualquer servidor anterior na porta 8080
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":8080 "') do (
    taskkill /F /PID %%a >nul 2>&1
)

:: Inicia servidor HTTP simples com PowerShell e abre o navegador
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$port = 8080; $dir = '%~dp0'; " ^
    "Add-Type -AssemblyName System.Net; " ^
    "$listener = New-Object System.Net.HttpListener; " ^
    "$listener.Prefixes.Add(\"http://localhost:$port/\"); " ^
    "$listener.Start(); " ^
    "Start-Process \"http://localhost:$port/index.html\"; " ^
    "Write-Host ' [OK] Servidor iniciado em http://localhost:' + $port; " ^
    "Write-Host ' Mantenha esta janela ABERTA. Feche para encerrar.'; " ^
    "while ($listener.IsListening) { " ^
    "  $ctx = $listener.GetContext(); " ^
    "  $req = $ctx.Request; $resp = $ctx.Response; " ^
    "  $path = $req.Url.LocalPath.TrimStart('/'); " ^
    "  if ($path -eq '' -or $path -eq '/') { $path = 'index.html' }; " ^
    "  $file = Join-Path $dir $path; " ^
    "  if (Test-Path $file -PathType Leaf) { " ^
    "    $bytes = [System.IO.File]::ReadAllBytes($file); " ^
    "    $resp.OutputStream.Write($bytes, 0, $bytes.Length); " ^
    "  } else { $resp.StatusCode = 404 }; " ^
    "  $resp.OutputStream.Close(); " ^
    "}"

:fim
echo.
echo  GEOportal encerrado.
pause
