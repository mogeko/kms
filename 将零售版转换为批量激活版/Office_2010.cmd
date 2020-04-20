@echo off

:ADMIN
openfiles >nul 2>nul ||(
  echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
  echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
  "%temp%\getadmin.vbs" >nul 2>&1
  goto:eof
)
del /f /q "%temp%\getadmin.vbs" >nul 2>nul
pushd "%~dp0"

set "ospp=%ProgramFiles%\Microsoft Office\Office14\ospp.vbs"
if not exist "%ospp%" (
set "ospp=%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs"
)

echo Are you sure you want to change your Office?
pause

REM Retail2VOL
echo ---Processing--------------------------
echo ---------------------------------------
cscript "%ospp%" /inslic:"..\root\Licenses14\ProPlusVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProPlusVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProPlusVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProPlusVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProPlusVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProPlusVL_MAK-ul-phn.xrm-ms"

cscript "%ospp%" /inslic:"..\root\Licenses14\ProjectProVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProjectProVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProjectProVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProjectProVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProjectProVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\ProjectProVL_MAK-ul-phn.xrm-ms"

cscript "%ospp%" /inslic:"..\root\Licenses14\VisioProVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\VisioProVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\VisioProVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\VisioProVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\VisioProVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\VisioProVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses14\VisioProVL_MAK-ul-phn.xrm-ms"
echo ---------------------------------------
echo ---------------------------------------
echo ---Exiting-----------------------------
