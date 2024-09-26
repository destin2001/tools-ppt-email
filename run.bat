@echo off
set "input_folder=./input"

for %%F in ("%input_folder%\*") do (
    echo Ready to process pptx file: {{%%~nxF}}
    pause
    echo Processing file pptx: {{%%~nxF}}
    echo ...
    .venv\Scripts\python main.py "%%F"
    echo ...
    echo Export file pptx: {{%%~nxF}} to Outlook email body has been run successfully!...
)

pause