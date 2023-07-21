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
SET folder=PR %creation_date% - %company% - %description%
SET folder=%folder:"=%

:: Create directories 
cd "H:\Documents\Purchase Orders"
mkdir %year%
cd %year%
mkdir "%folder%"
cd %folder%
start .
mkdir "Receipts"

:: Finish and exit
CALL :rename_quote
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

:exit
echo.
SET /p input="%folder% was created. Press [ENTER] to close."
GOTO :EOF