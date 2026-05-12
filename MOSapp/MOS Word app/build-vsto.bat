@echo off
setlocal enabledelayedexpansion

echo === VSTOアドイン ビルドスクリプト ===

set CONFIGURATION=Debug
if "%1"=="Release" set CONFIGURATION=Release

echo.
echo NuGetパッケージを復元しています...
nuget restore "MOS Word app.sln"
if errorlevel 1 (
    echo NuGetパッケージの復元に失敗しました。続行します...
)

echo.
echo ビルドを実行しています...
msbuild "MOSWordVSTOAddIn\MOSWordVSTOAddIn.csproj" /p:Configuration=%CONFIGURATION% /p:Platform="Any CPU" /v:minimal /nologo

if errorlevel 1 (
    echo ビルドに失敗しました。
    exit /b 1
)

echo.
echo ビルド出力を確認しています...
set OUTPUT_DIR=MOSWordVSTOAddIn\bin\%CONFIGURATION%
set ALL_FILES_EXIST=1

if exist "%OUTPUT_DIR%\MOSWordVSTOAddIn.dll" (
    echo   [OK] MOSWordVSTOAddIn.dll
) else (
    echo   [NG] MOSWordVSTOAddIn.dll ^(見つかりません^)
    set ALL_FILES_EXIST=0
)

if exist "%OUTPUT_DIR%\MOSWordVSTOAddIn.dll.manifest" (
    echo   [OK] MOSWordVSTOAddIn.dll.manifest
) else (
    echo   [NG] MOSWordVSTOAddIn.dll.manifest ^(見つかりません^)
    set ALL_FILES_EXIST=0
)

if exist "%OUTPUT_DIR%\MOSWordVSTOAddIn.vsto" (
    echo   [OK] MOSWordVSTOAddIn.vsto
) else (
    echo   [NG] MOSWordVSTOAddIn.vsto ^(見つかりません^)
    set ALL_FILES_EXIST=0
)

if %ALL_FILES_EXIST%==1 (
    echo.
    echo ビルドが正常に完了しました！
    echo 出力ディレクトリ: %OUTPUT_DIR%
) else (
    echo.
    echo 警告: 一部のファイルが見つかりませんでした。
)

pause

