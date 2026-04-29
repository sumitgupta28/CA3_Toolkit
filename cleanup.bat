@echo off
REM Cleanup script — deletes generated output files.
REM Keeps: CA3_Marks/marks.xlsx (asks), students/, source scripts.

echo CA3 Toolkit — Cleanup
echo ----------------------------------------

REM Delete generated folders
for %%D in (output CA3_Marks_Pdf __pycache__) do (
    if exist "%%D\" (
        rmdir /s /q "%%D"
        echo   Deleted: %%D\
    ) else (
        echo   Skipped: %%D\ (not found)
    )
)

REM Delete contents of students\ (keep the folder)
if exist "students\" (
    for /f "delims=" %%F in ('dir /b "students\*" 2^>nul') do (
        if exist "students\%%F\" (
            rmdir /s /q "students\%%F"
        ) else (
            del /q "students\%%F"
        )
    )
    echo   Cleared: students\
) else (
    echo   Skipped: students\ (not found)
)

echo ----------------------------------------

REM Ask about marks.xlsx
if exist "CA3_Marks\marks.xlsx" (
    set /p answer="  Delete CA3_Marks\marks.xlsx? (y/N): "
    if /i "%answer%"=="y" (
        del "CA3_Marks\marks.xlsx"
        echo   Deleted: CA3_Marks\marks.xlsx
    ) else (
        echo   Kept:    CA3_Marks\marks.xlsx
    )
) else (
    echo   Skipped: CA3_Marks\marks.xlsx (not found)
)

echo ----------------------------------------
echo Done.
pause
