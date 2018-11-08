@echo off
chcp 65001

echo What do you want?
echo;
echo   1. 激活 Windows (仅限 VOL 版)
echo   2. 激活 Office 2016 (零售版)
echo   3. 激活 Office 2016 (VOL 版)
echo;
echo   4. 激活 Office Project 2016
echo   5. 激活 Office Visio 2016
echo;
echo   6. 将零售版 Office 转换为 VOL 版 Office，不激活
echo;
echo   0. 退出
echo;
echo;
echo 本脚本仅支持激活 VOL 版 Windows，零售版 Windows 请另寻它法。  
echo 本脚本默认你使用的是 2016 版的 Office (包括 Visio 和 Project)，其他版本的 Office 请参考 README 更改相关脚本再激活。
echo;

set /p flag=请输入 (数字): 

if %flag%==1 call Windows_Activate\Windows_Activate.cmd

if %flag%==2 call Office_Activate\Retail_Activate\Retail2VOL+Activate.cmd

if %flag%==3 call Office_Activate\VOL_Activate.cmd

if %flag%==4 call Office_Activate\Retail_Activate\Component\ProjectPro.cmd

if %flag%==5 call Office_Activate\Retail_Activate\Component\VisioPro.cmd

if %flag%==6 call Office_Activate\Retail_Activate\Retail2VOL_Only.cmd

