@echo off
setlocal enabledelayedexpansion

rem Set your folders:
set "source_folder=C:\AI\Copypdf"
set "target_folder=C:\AI\2"

rem Make sure the target folder exists
if not exist "%target_folder%" mkdir "%target_folder%"

rem Loop through each PDF file recursively
for /f "delims=" %%F in ('dir "%source_folder%\*.pdf" /s /b') do (
    set "filename=%%~nxF"
    set "name=%%~nF"
    set "ext=%%~xF"
    set "dest=%target_folder%\!filename!"

    rem Initialize counter
    set /a counter=1

    rem Loop to find unique filename if duplicates exist
    for %%Z in ("!dest!") do (
        if exist "%%~fZ" (
            rem if file exists, loop to find new unique name
            set "dest=%target_folder%\!name!_!counter!!ext!"
            for /l %%i in (1,1,1000) do (
                if exist "!dest!" (
                    set /a counter+=1
                    set "dest=%target_folder%\!name!_!counter!!ext!"
                )
            )
        )
    )

    echo Copying "%%F" to "!dest!"
    copy "%%F" "!dest!" >nul
)

echo.
echo All PDF files copied successfully.
pause