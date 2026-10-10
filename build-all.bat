@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1

rem ==========================================================================
rem  build-all.bat
rem  The one entry point: builds all seven workbooks, converts each .docx to
rem  .pdf via Microsoft Word, then merges the seven PDFs into one file next to
rem  this script:
rem      Arabic_Workbooks_Complete.pdf
rem
rem  Any failing step stops the build. Nothing is merged from an archived copy.
rem
rem  Usage:
rem      build-all.bat              full build (docx -> pdf -> merge)
rem      build-all.bat /nobuild     skip node build, convert + merge only
rem ==========================================================================

set "ROOT=%~dp0"
set "OUTPUT=%ROOT%Arabic_Workbooks_Complete.pdf"

set "DO_BUILD=1"
if /i "%~1"=="/nobuild"   set "DO_BUILD=0"
if /i "%~1"=="/no-build"  set "DO_BUILD=0"
if /i "%~1"=="/merge"     set "DO_BUILD=0"

echo.
echo ==========================================================
echo   Arabic Workbooks - Full Build
echo ==========================================================
echo.

where node >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] Node.js not found in PATH. Install Node.js 16 or newer.
    goto :fail
)

where powershell >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] powershell.exe not found in PATH.
    goto :fail
)

rem The seven projects share the root node_modules, so one install covers them.
if not exist "%ROOT%node_modules\docx" goto :install_root
if not exist "%ROOT%node_modules\pdf-lib" goto :install_root
if not exist "%ROOT%node_modules\svg2img" goto :install_root
goto :after_root_install

:install_root
echo   [INFO] Installing shared dependencies (docx, pdf-lib, svg2img) ...
pushd "%ROOT%"
call npm install --no-audit --no-fund
rem "if errorlevel" reads the live errorlevel, unlike %ERRORLEVEL% in a block
if errorlevel 1 (
    popd
    echo   [ERROR] npm install failed.
    goto :fail
)
popd
echo.

:after_root_install

pushd "%ROOT%"

if "%DO_BUILD%"=="1" (
    echo [1/3] Building workbooks ^(.docx^)...
    echo.
    call :build "Arabic_Alphabet_Workbook" docx svg2img
    if errorlevel 1 (popd & goto :fail)
    call :build "Arabic_Harakat_Workbook" docx
    if errorlevel 1 (popd & goto :fail)
    call :build "Arabic_Syllable_Analysis_Workbook" docx
    if errorlevel 1 (popd & goto :fail)
    call :build "Arabic_Word_Analysis_Workbook" docx svg2img
    if errorlevel 1 (popd & goto :fail)
    call :build "Arabic_Letter_Shapes_Workbook" docx svg2img
    if errorlevel 1 (popd & goto :fail)
    call :build "Arabic_Word_Composition_Workbook" docx svg2img
    if errorlevel 1 (popd & goto :fail)
    call :build "Arabic_Stories_Workbook" docx
    if errorlevel 1 (popd & goto :fail)
) else (
    echo [1/3] Skipping node build ^(/nobuild^).
)

echo.
echo [2/3] Converting .docx to .pdf via Microsoft Word...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "docx-to-pdf.ps1"
if errorlevel 1 (
    echo.
    echo   [ERROR] PDF conversion failed. Close Microsoft Word and try again.
    popd
    goto :fail
)

echo.
echo [3/3] Merging PDFs...
echo.
node merge-pdfs.js
if errorlevel 1 (
    popd
    goto :fail
)

if not exist "%OUTPUT%" (
    echo.
    echo   [ERROR] Output PDF was not created: %OUTPUT%
    popd
    goto :fail
)

popd
echo.
echo ==========================================================
echo   DONE
echo   %OUTPUT%
echo ==========================================================
echo.
if not defined BUILD_ALL_NO_PAUSE pause
endlocal
exit /b 0


rem ---------------------------------------------------------------- build sub
rem  :build <project> <module> [<module> ...]
rem  Installs into the project only when a module it needs is in neither the
rem  shared root node_modules nor the project one.
:build
set "PROJ=%~1"
shift
set "NEED_INSTALL=0"
:check_deps
if "%~1"=="" goto :do_build
if exist "%ROOT%node_modules\%~1" goto :next_dep
if exist "%ROOT%%PROJ%\node_modules\%~1" goto :next_dep
set "NEED_INSTALL=1"
:next_dep
shift
goto :check_deps

:do_build
if "%NEED_INSTALL%"=="1" (
    echo   [INFO] Installing %PROJ% dependencies ...
    pushd "%ROOT%%PROJ%"
    call npm install --no-audit --no-fund
    if errorlevel 1 (
        popd
        echo.
        echo   [ERROR] npm install failed in %PROJ%
        exit /b 1
    )
    popd
)
pushd "%ROOT%%PROJ%"
node create-workbook.js
if errorlevel 1 (
    popd
    echo.
    echo   [ERROR] Build failed in %PROJ%
    exit /b 1
)
popd
exit /b 0


rem -------------------------------------------------------------------- fail
:fail
set "RC=1"
if not defined BUILD_ALL_NO_PAUSE pause
endlocal & exit /b %RC%
