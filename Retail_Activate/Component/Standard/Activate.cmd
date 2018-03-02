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

cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
cscript ospp.vbs /sethst:kms.03k.org
cscript ospp.vbs /act
