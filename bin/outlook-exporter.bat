@ECHO OFF

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
:: outlook-exporter.bat
::
:: Copyright by toolarium, all rights reserved.
::
:: This file is part of the toolarium outlook-exporter.
::
:: The outlook-exporter is free software: you can redistribute it and/or modify
:: it under the terms of the GNU General Public License as published by
:: the Free Software Foundation, either version 3 of the License, or
:: (at your option) any later version.
::
:: The outlook-exporter is distributed in the hope that it will be useful,
:: but WITHOUT ANY WARRANTY; without even the implied warranty of
:: MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
:: GNU General Public License for more details.
::
:: You should have received a copy of the GNU General Public License
:: along with Foobar. If not, see <http://www.gnu.org/licenses/>.
::
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

set PN=%~nx0
set SCRIPT_PATH=%~dp0
set CB_LINE=----------------------------------------------------------------------------------------
set "CB_LINEHEADER=.: "
set "CB_PARAMETERS="
set CB_VERBOSE=false
set "CURRENT_PATH=%PWD%"


:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:CHECK_PARAMETER
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
if %0X==X goto CHECK_POWERSHELL_ACCESS
if .%1==.-h goto HELP
if .%1==.--help goto HELP
if .%1==.--verbose set CB_VERBOSE=true & shift & goto :CHECK_PARAMETER
if not .%1==. set "CB_PARAMETERS=%CB_PARAMETERS% %~1"
shift
goto :CHECK_PARAMETER


:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:HELP
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo %PN% - Outlook exporter.
echo.
echo usage, select current month: 
echo %PN% 
echo.
echo usage, select month of the year, e.g. 7
echo %PN% 7
echo.
echo usage, select month and year, e.g. 7.2025
echo %PN% 7.2025
echo.
echo usage, select specific date in format dd.mm.yyyy, e.g. 2.7.2025
echo %PN% 2.7.2025
echo.
echo usage, select specific date range by date in format dd.mm.yyyy, e.g. 2.7.2025 - 4.7.2025
echo %PN% 2.7.2025 4.7.2025
echo.
echo Overview of the available OPTIONs:
echo  -h, --help           Show this help message.
goto END


:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:CHECK_POWERSHELL_ACCESS
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Check PowerShell Execution Policy
for /f "usebackq tokens=*" %%i in (`powershell -NoProfile -Command "(Get-ExecutionPolicy -Scope CurrentUser)"`) do set ExecPolicy=%%i
SET arg1=%~1
SET arg2=%~2
:: Allowed policies where scripts can run
set AllowedPolicy1=RemoteSigned
set AllowedPolicy2=Unrestricted
set AllowedPolicy3=Bypass
set AllowedPolicy4=Undefined

if /I "%ExecPolicy%"=="%AllowedPolicy1%" goto MAIN
if /I "%ExecPolicy%"=="%AllowedPolicy2%" goto MAIN
if /I "%ExecPolicy%"=="%AllowedPolicy3%" goto MAIN
if /I "%ExecPolicy%"=="%AllowedPolicy4%" goto MAIN

:: If not allowed
echo .: PowerShell script execution is not allowed by your current Execution Policy: %ExecPolicy%
echo    To enable script execution, run PowerShell as Administrator and execute:
echo.
echo    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
echo.
echo .: Exiting with failure status.
exit /b 1


:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:RUN_SCRIPT
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
if .%CB_VERBOSE% == .true echo %CB_LINEHEADER%Execution PowerShell script: "%SCRIPT_PATH%\%1" %CB_PARAMETERS%
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%\ps\%1" %CB_PARAMETERS%

echo.
goto :eof


:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:MAIN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo %CB_LINE%
echo    Outlook Exporter, v1.0
echo %CB_LINE%
if .%CB_VERBOSE% == .true echo %CB_LINEHEADER%Execution Policy is: %ExecPolicy%

set "CONFIG_PATH=%SCRIPT_PATH%\..\config"
if not exist %CONFIG_PATH% (
	echo %CB_LINEHEADER%Create config path... 
	mkdir %CONFIG_PATH% >nul 2>nul
	echo %CB_LINEHEADER%Create default empty config... 
	type NUL >> %CONFIG_PATH%\calendar-attendee-filter.txt
	type NUL >> %CONFIG_PATH%\calendar-subject-filter.txt
)

set "CUSTOMER_FILTER_PATH=config\customer-filter"
if not exist %CUSTOMER_FILTER_PATH% (
	echo %CB_LINEHEADER%Create customer filter path... 
	mkdir %CUSTOMER_FILTER_PATH% >nul 2>nul
	echo %CB_LINEHEADER%Create sample filter files... 
	type NUL >> %CUSTOMER_FILTER_PATH%\sample-filter.txt
	type NUL >> %CUSTOMER_FILTER_PATH%\sample-attendees.txt
	type NUL >> %CUSTOMER_FILTER_PATH%\sample-duration.txt
)

echo.
cd %SCRIPT_PATH%\..
call :RUN_SCRIPT calendar-exporter.ps1
call :RUN_SCRIPT mail-exporter.ps1
call :RUN_SCRIPT customer-filter.ps1
call :RUN_SCRIPT customer-reports.ps1
GOTO END


:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
cd %CURRENT_PATH%
:END
exit /b 0
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::