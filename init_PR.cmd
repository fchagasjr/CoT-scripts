@echo off

CALL :script_head

:: Setup variables
SET month=%date:~0,2%
SET day=%date:~3,2%
SET year=%date:~-4%
SET creation_date=%year%%month%%day%
CALL :set_company
SET /p description="Enter Order Description:"
CALL :confirm "%description%" "purchase description"
SET order_info=PR %creation_date% - %description%
SET order_info=%order_info:"=%

:: Create directories 
CALL :focus_directory "H:\Documents\Purchase Orders"
CALL :focus_directory "%year%"
CALL :focus_directory "%company%"
CALL :focus_directory "%order_info%"
start .
mkdir "Receipts"

:: Finish and exit
CALL :rename_quote
CALL :create_pr
CALL :exit

:script_head
cls
echo ===Purchase Request Folder Creation Script===
echo.
EXIT /B

:set_company
FOR /F "tokens=1,2 delims=," %%g IN (db\companies.csv) DO echo %%g - %%h
SET /p id="Select company by number: "
FOR /F "tokens=1,2 delims=," %%g IN (db\companies.csv) DO (
  IF %id%==%%g SET company="%%h"
)
IF [%company%]==[] (
  echo "Invalid number"
  CALL :set_company
) ELSE (
  CALL :confirm %company% "supplier company"
  )	
)
SET company=%company:"=%
EXIT /B

:focus_directory
IF NOT EXIST "\%~1" mkdir %1
cd %1
EXIT /B

:confirm
call :script_head
SET /p input="%1 was selected as %~2.  Press [ENTER] to continue."
EXIT /B

:rename_quote
CALL :script_head
SET /p input="Move quote to folder and press [ENTER] to continue."
move *.pdf "QT %creation_date%.pdf"
start "" /max "QT %creation_date%.pdf"
EXIT /B

:create_pr
CALL :script_head
start "" /max "H:/Scripts/misc/PR_form.xlsx"
SET /p input="Update request sheet and press [ENTER] to continue."
cscript H:/Scripts/xlsx_to_pdf.vbs "%CD%\PR %creation_date%.pdf"
EXIT /B

:exit
CALL :script_head
SET /p input="%company% - %order_info% was created. Press [ENTER] to close."
GOTO :EOF