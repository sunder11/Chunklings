@echo off
setlocal enabledelayedexpansion

REM Define path to LibreOffice
set "LibreOfficePath=C:\Program Files\LibreOffice\program"
set "path=%LibreOfficePath%;%path%"

REM Define output directory
set "outputDir=C:\AI\1"

REM Create output folder if it doesn't exist
if not exist "%outputDir%" mkdir "%outputDir%"

REM Loop recursively through all .wpd files
for /r %%F in (*.wpd) do call :ConvertFile "%%~fF"

echo All conversions completed successfully.
pause
exit /b

REM -------------- Conversion function starts here ----------------
:ConvertFile
setlocal enabledelayedexpansion

set "sourceFile=%~1"
set "baseName=%~n1"

REM Create a unique temporary folder
set "uniqueTemp=%temp%\LibreTemp_!random!!random!!random!"
mkdir "!uniqueTemp!"

REM Launch LibreOffice conversion process
soffice.exe -headless -convert-to docx:"MS Word 2007 XML" -outdir "!uniqueTemp!" "!sourceFile!"

REM Wait for conversion to complete (check existence of the file and retry for max 20 seconds)
set "convertedFile=!uniqueTemp!\!baseName!.docx"
set "waitTime=0"
:wait_loop
if exist "!convertedFile!" (
    goto file_ready
) else (
    timeout /t 1 >nul
    set /a waitTime+=1
    if !waitTime! GEQ 20 (
        echo ***ERROR***: Conversion timed out for file "!sourceFile!"
        rd /s /q "!uniqueTemp!" >nul
        goto end_local
    )
    goto wait_loop
)

:file_ready
REM Generate unique file name in output dir (handle duplicates)
set "finalFile=!baseName!.docx"
set "counter=1"
:check_duplicate
if exist "%outputDir%\!finalFile!" (
    set "finalFile=!baseName!_!counter!.docx"
    set /a counter+=1
    goto check_duplicate
)

REM Safely move file to the output directory
move /y "!convertedFile!" "%outputDir%\!finalFile!" >nul
echo Converted "!sourceFile!" to "%outputDir%\!finalFile!"

REM Clean up the unique temporary folder
rd /s /q "!uniqueTemp!" >nul 2>&1

:end_local
endlocal
exit /b