@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1

rem ==========================================================================
rem  images-to-book.bat
rem  Builds a book out of every image in the images\ folder next to this
rem  script, using the same full-bleed cover page as the four workbook
rem  projects: one A5 page per image, then a PDF via Microsoft Word.
rem
rem  By default the images come from images\ next to this script and the output
rem  lands there too. Drop the images in and run the script; --images and --out
rem  only exist so the desktop GUI can point the tool somewhere else.
rem
rem  Usage:
rem      images-to-book.bat                 A5, DOCX + PDF
rem      images-to-book.bat --size a4       A4 instead of A5
rem      images-to-book.bat --no-pdf        DOCX only, no Word needed
rem      images-to-book.bat --images <dir> --out <dir>   other folders
rem ==========================================================================

set "ROOT=%~dp0"
rem images-to-book.bat changes into its own folder below, so remember where the
rem user ran it from: relative --images / --out paths mean folders next to them.
set "IMAGES_TO_BOOK_CWD=%CD%"

echo.
echo ==========================================================
echo   Images to Book  -  full-bleed cover style
echo ==========================================================
echo.

where node >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] Node.js not found in PATH. Install Node.js 16 or newer.
    goto :fail
)

rem The projects share the root node_modules, so one install covers this tool.
if exist "%ROOT%..\node_modules\docx" goto :after_deps

echo   [INFO] Installing docx into the shared root node_modules ...
pushd "%ROOT%.."
call npm install --no-audit --no-fund
rem "if errorlevel" reads the live errorlevel, unlike %ERRORLEVEL% in a block
if errorlevel 1 (
    popd
    echo   [ERROR] npm install failed.
    goto :fail
)
popd
echo.

:after_deps
rem The dot matters: %ROOT% ends in a backslash, and a quoted path that ends
rem in \" is read as an escaped quote, which makes pushd fail.
pushd "%ROOT%."
node images-to-book.js %*
set "RC=%ERRORLEVEL%"
popd

echo.
if not "%RC%"=="0" goto :fail

echo ==========================================================
echo   DONE
echo   (the output path is printed above)
echo ==========================================================
echo.
if not defined IMAGES_TO_BOOK_NO_PAUSE pause
endlocal
exit /b 0

rem -------------------------------------------------------------------- fail
:fail
echo.
echo   [ERROR] Build failed. See the message above.
echo.
if not defined IMAGES_TO_BOOK_NO_PAUSE pause
endlocal & exit /b 1
