@echo off
chcp 65001

REM 获得管理员权限
:ADMIN
openfiles >nul 2>nul ||(
  echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
  echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
  "%temp%\getadmin.vbs" >nul 2>&1
  goto:eof
)
del /f /q "%temp%\getadmin.vbs" >nul 2>nul
pushd "%~dp0"

REM ospp.vbs 路径
set "ospp=%ProgramFiles%\Microsoft Office\Office16\ospp.vbs"
if not exist "%ospp%" (
set "ospp=%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs"
)

echo What do you want?
echo;
echo   1. 激活 Windows (仅限 VOL 版)
echo   2. 激活 Office 2016 (零售版)
echo   3. 激活 Office 2016 (VOL 版)
echo;
echo   4. 激活 Office Project Pro 2016 (零售版)
echo   5. 激活 Office Visio Pro 2016 (零售版)
echo;
echo   6. 将零售版 Office 2016 转换为 VOL 版 Office, 不激活
echo   7. 将零售版 Office Project Pro 2016 转换为 VOL 版 Office, 不激活
echo   8. 将零售版 Office Visio Pro 2016 转换为 VOL 版 Office, 不激活
echo;
echo   0. 退出
echo;
echo;
echo 本脚本仅支持激活 VOL 版 Windows, 零售版 Windows 请另寻它法。  
echo 本脚本默认激活的是 2016 版的 Office (包括 Visio 和 Project), 其他版本的 Office 请从 Atom 文件夹中寻找相关脚本激活 
echo 需要激活 Vol 版的 Office Project 2016 或 Office Visio 2016 请选择 [3]
echo;

set /p flag=请输入 (数字): 

if %flag%==1 goto:Windows_Activate
if %flag%==2 goto:Office_Retail2VOL
if %flag%==3 goto:Office_VOL_Activate
if %flag%==4 goto:Project_Retail2VOL
if %flag%==5 goto:Visio_Retail2VOL
if %flag%==6 goto:Office_Retail2VOL
if %flag%==7 goto:Project_Retail2VOL
if %flag%==8 goto:Visio_Retail2VOL
if %flag%==0 exit
exit


REM 激活 Windows  
:Windows_Activate

echo Are you sure you want to activate your Windows?
pause
slmgr /skms kms.03k.org
slmgr /ato
exit


REM 激活 Office 2016 (零售版)  
:Office_VOL_Activate

echo Are you sure you want to activate your Office?
pause
cscript "%ospp%" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript "%ospp%" /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT
cscript "%ospp%" /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
cscript "%ospp%" /sethst:kms.03k.org
cscript "%ospp%" /act
echo Current Office status:
cscript "%ospp%" /dstatus
pause
exit


REM 将零售版 Office 转换为 VOL 版  
:Office_Retail2VOL
echo Are you sure you want to change your Office?
pause
cscript "%ospp%" /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-phn.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-phn.xrm-ms"
if %flag%==2 goto:Office_VOL_Activate
if %flag%==6 exit


REM 将零售版 Project 转换为 VOL 版  
:Project_Retail2VOL

echo Are you sure you want to change and activate your Office?
pause
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-phn.xrm-ms"
if %flag%==4 goto:Office_VOL_Activate
if %flag%==7 exit


REM 将零售版 Visio 转换为 VOL 版  
:Visio_Retail2VOL

echo Are you sure you want to change and activate your Office?
pause
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-pl.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-ppd.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-oob.xrm-ms"
cscript "%ospp%" /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-phn.xrm-ms"
if %flag%==5 goto:Office_VOL_Activate
if %flag%==8 exit