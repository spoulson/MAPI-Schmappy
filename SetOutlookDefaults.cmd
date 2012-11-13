::
:: Configure MAPI profile only once
::

@echo off
pushd "%~dp0"
setlocal

::
:: If CHECKFLAG file already exists, skip configuration
::
set CHECKFLAG=%temp%\%~n0.txt
if exist "%CHECKFLAG%" goto :Exit

::
:: Configure MAPI profile
:: Abort if unsuccessful
::
MAPIDefaultAddressList "Your Preferred Address List Here"
if errorlevel 1 goto :Exit

MAPIAddrListSearch "Your Preferred Address List Here" Contacts "Global Address List"
if errorlevel 1 goto :Exit

::
:: Create CHECKFLAG file to indicate success
::
echo. > "%CHECKFLAG%"


:Exit
endlocal
popd
