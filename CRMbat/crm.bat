@echo off
setlocal enabledelayedexpansion

REM ===== FIELD CONFIGURATION SECTION =====
REM Define ALL field names in your CSV file in the exact order they appear
REM Add or remove fields as needed to match your actual CSV structure (UP TO 50)
set "field_names=FIRST_NAME,LAST_NAME,MY_FEE,STATUS,DATE,CASE_TYPE,NOTES,COSTS,TOTAL,PHONE1,PHONE2,EMAIL,GCLID,EXPIRES"

REM Define which field positions contain the data we need for output
REM Count starting from 1 (first field = 1, second field = 2, etc.)
set "pos_FIRST_NAME=1"
set "pos_LAST_NAME=2"
set "pos_DATE=5"
set "pos_CASE_TYPE=6"
set "pos_NOTES=7"
set "pos_TOTAL=9"
set "pos_PHONE1=10"
set "pos_EMAIL=12"
set "POS_STATUS=4"
set "POS_MY_FEE=3"
set "POS_COSTS=8"

REM =======================================

REM Check if crm.csv exists
if not exist "crm.csv" (
    echo Error: crm.csv file not found!
    pause
    exit /b 1
)

REM Initialize variables
set "line_count=0"
set "record_count=0"
set "output_file=CRM_Records_Output.html"

REM Create HTML file with header
echo ^<!DOCTYPE html^> > "%output_file%"
echo ^<html^> >> "%output_file%"
echo ^<head^> >> "%output_file%"
echo ^<title^>CRM Records Report^</title^> >> "%output_file%"
echo ^<style^> >> "%output_file%"
echo body { font-family: 'Times New Roman', serif; font-size: 12pt; margin: 1in; } >> "%output_file%"
echo .record { margin-bottom: 20px; padding: 15px; border-left: 3px solid #0066cc; background-color: #f9f9f9; } >> "%output_file%"
echo .record-header { font-weight: bold; color: #0066cc; } >> "%output_file%"
echo hr { margin: 20px 0; border: 1px solid #cccccc; } >> "%output_file%"
echo h1 { color: #0066cc; border-bottom: 2px solid #0066cc; padding-bottom: 10px; } >> "%output_file%"
echo ^</style^> >> "%output_file%"
echo ^</head^> >> "%output_file%"
echo ^<body^> >> "%output_file%"
echo ^<h1^>CRM Records Report^</h1^> >> "%output_file%"

echo Processing CSV records from crm.csv...
echo Field configuration: %field_names%
echo Output will be saved to: %output_file%
echo.

REM Read the CSV file line by line
for /f "usebackq delims=" %%a in ("crm.csv") do (
    set /a line_count+=1
    set "current_line=%%a"
    
    REM Skip the header row (first line)
    if !line_count! equ 1 (
        echo Header row detected: !current_line!
        echo.
    ) else (
        call :ParseAndDisplayRecord "!current_line!"
    )
)

REM Close HTML file
echo ^</body^> >> "%output_file%"
echo ^</html^> >> "%output_file%"

echo.
echo Processing complete. Total records processed: !record_count!
echo Output saved to: %output_file%
echo.
echo To convert to .docx:
echo 1. Open %output_file% in Microsoft Word
echo 2. Go to File ^> Save As
echo 3. Choose "Word Document (*.docx)" as the file type
echo 4. Click Save
pause
exit /b 0

:ParseAndDisplayRecord
REM Enhanced CSV parsing function that handles all fields
set "line=%~1"
set "field_num=0"

REM Clear all previous field values (up to 50 fields to be safe)
for /l %%i in (1,1,50) do set "field%%i="

REM Parse CSV with quote handling
set "in_quotes=false"
set "current_field="

:parse_loop
if "!line!"=="" goto :done_parsing

REM Get first character
set "char=!line:~0,1!"
set "line=!line:~1!"

if "!char!"=="," (
    if "!in_quotes!"=="false" (
        call :AddField "!current_field!"
        set "current_field="
    ) else (
        set "current_field=!current_field!!char!"
    )
) else if "!char!"=="""" (
    if "!in_quotes!"=="false" (
        set "in_quotes=true"
    ) else (
        set "in_quotes=false"
    )
) else (
    set "current_field=!current_field!!char!"
)

goto :parse_loop

:done_parsing
REM Add the last field
call :AddField "!current_field!"

REM Extract the specific fields we need using the position mapping
set "FIRST_NAME=!field%pos_FIRST_NAME%!"
set "LAST_NAME=!field%pos_LAST_NAME%!"
set "DATE=!field%pos_DATE%!"
set "PHONE1=!field%pos_PHONE1%!"
set "EMAIL=!field%pos_EMAIL%!"
set "CASE_TYPE=!field%pos_CASE_TYPE%!"
set "TOTAL=!field%pos_TOTAL%!"
set "NOTES=!field%pos_NOTES%!"
set "STATUS=!field%pos_STATUS%!"
set "MY_FEE=!field%pos_MY_FEE%!"
set "COSTS=!field%pos_COSTS%!"

REM Clean up any remaining quotes
set "FIRST_NAME=!FIRST_NAME:"=!"
set "LAST_NAME=!LAST_NAME:"=!"
set "DATE=!DATE:"=!"
set "PHONE1=!PHONE1:"=!"
set "EMAIL=!EMAIL:"=!"
set "CASE_TYPE=!CASE_TYPE:"=!"
set "TOTAL=!TOTAL:"=!"
set "NOTES=!NOTES:"=!"
set "STATUS=!STATUS:"=!"
set "MY_FEE=!MY_FEE:"=!"
set "COSTS=!COSTS:"=!"

REM Count records
set /a record_count+=1

REM Debug output to console (optional - remove if not needed)
echo Processing Record !record_count!:
echo   Name: !FIRST_NAME! !LAST_NAME!
echo   Date: !DATE!
echo   Case Type: !CASE_TYPE!
echo   Total: $!TOTAL!

REM Output to HTML file
echo ^<div class="record"^> >> "%output_file%"
echo ^<div class="record-header"^>CRM Record:^</div^> >> "%output_file%"
echo ^<div class="record-header"^>:^</div^> >> "%output_file%"
echo ^<div class="record-header"^>Begin Record: !record_count!^</div^> >> "%output_file%"
echo ^<div class="record-header"^>:^</div^> >> "%output_file%"
echo ^<div class="record-header"^>NAME: !FIRST_NAME! !LAST_NAME!^</div^> >> "%output_file%"
echo ^<div class="name-line"^>You have a record in your CRM for !FIRST_NAME! !LAST_NAME!. The date of the record is !DATE!. !FIRST_NAME!'s phone number is !PHONE1!. !FIRST_NAME!'s email address is !EMAIL!. !FIRST_NAME!'s case type is !CASE_TYPE!. I quoted !FIRST_NAME! an attorney fee of $!MY_FEE! and costs of !COSTS!. The notes are as follows: !NOTES!. The status of the matter is: !STATUS!. (0=Call, 1=Email, 2=Left Word, 3=Emailed, 4=Made Quote, 5=Retained me).^</div^> >> "%output_file%"
echo ^<div class="record-header"^>:^</div^> >> "%output_file%"
echo ^<div class="record-header"^>End Record: !record_count!^</div^> >> "%output_file%"
echo ^</div^> >> "%output_file%"
echo ^<hr^> >> "%output_file%"
echo   ✓ Record added to HTML file
echo ----------------------------------------
echo.

goto :eof

:AddField
set /a field_num+=1
set "field!field_num!=%~1"
goto :eof