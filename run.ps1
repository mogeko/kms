# 强制以管理员权限以运行
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

$script:LICENCE = @"
**********************************************
* Author:  Mogeko                            *
* Github:  https://github.com/Mogeko/KMS     *
* LICENCE: GNU General Public License v3.0   *
**********************************************
"@
$script:KMS_SERVER = 'kms.mogeko.me'

if (Test-Path .\data.json) {
    $script:MENU_INFO = Get-Content .\data.json | ConvertFrom-Json
}
    

function Main {
    param (
        $main
    )
    $flag = $true;
    while ($flag) {
        # Drew Menu
        Invoke-Expression 'cls';
        Write-Host $script:LICENCE;
        Write-Host "$nl";
        Write-Host $main.description;
        Write-Host "$nl";
        for ($i = 0; $i -lt $main.option.Count; $i++) {
            "    {0}) {1}" -f ($i+1), $main.option[$i].Name;
        }
        Write-Host "    $nl";
        Write-Host "    0) 返回上级菜单 (没有则退出程序)";
        Write-Host "    q) 退出程序";
        Write-Host "    $nl";
        # INPUT
        $input = Read-Host "输入序号 (可多选)";
        $input = $input.ToCharArray() | Sort-Object -Unique;
        # Menu logic
        foreach($inputItem in $input) {
            if ($inputItem -eq "q") {
                Invoke-Expression 'cls';
                exit 0;
            } elseif($inputItem -eq "0") {
                Invoke-Expression 'cls';
                $flag = $false;
            } elseif ($main.option[$inputItem-49].submenu) {
                Main  $main.option[$inputItem-49].submenu;
            } elseif ($main.option[$inputItem-49].action -eq "ActivateWindows") {
                ActivateWindows $main.option[$inputItem-49].name;
            } elseif ($main.option[$inputItem-49].action -eq "ActivateOffice") {
                ActivateOffice $main.Option[$inputItem-49].name $main.option[$inputItem-49].key $main.option[$inputItem-49].osppdir;
            }
        }
    }
}

function ActivateWindows {
    param (
        [String]$name
    )
    $cmd = 'slmgr'
    $WARNING = @"
作者用爱发电，出了任何问题都是不会负责的
回车后就没有回头路了，你是否真的想好了？
"@
    Invoke-Expression 'cls';
    "正在{0}..." -f $name
    "使用 KMS Server: {0}" -f $script:KMS_SERVER;
    Write-Host $WARNING -f red
    Invoke-Expression 'Pause'

    Invoke-Expression "$cmd /skms $script:KMS_SERVER";
    Invoke-Expression "$cmd /ato";
}

function CheckOSPP {
    param (
        [String]$verion
    )
    if (Test-Path "$env:ProgramFiles\Microsoft Office\$verion\ospp.vbs") {
        return "$env:ProgramFiles\Microsoft Office\$verion\ospp.vbs"
    } elseif (Test-Path "$env:ProgramFiles(x86)\Microsoft Office\$verion\ospp.vbs") {
        return "$env:ProgramFiles(x86)\Microsoft Office\$verion\ospp.vbs"
    } else {
        Invoke-Expression 'cls'
        Write-Host "错误！未找到 ospp.vbs 文件" -f red
        exit 1
    }
}

function ActivateOffice {
    param (
        [String]$name, [String]$key, [String]$osppPath
    )
    $cmd = 'cscript'
    $ospp = CheckOSPP $osppPath
    $WARNING = @"
作者用爱发电，出了任何问题都是不会负责的
回车后就没有回头路了，你是否真的想好了？
如果 ospp.vbs 位置不对请立即 Ctrl+C 强制退出
"@
    Invoke-Expression 'cls';
    "正在激活 {0}..." -f $name
    "ospp.vbs 位置: {0}" -f $ospp;
    "使用 KMS Server: {0}" -f $script:KMS_SERVER;
    "使用 KMS Key: {0}" -f $key
    Write-Host $WARNING -f red
    Invoke-Expression 'Pause'

    Invoke-Expression "$cmd '$ospp' /inpkey:$key"
    Invoke-Expression "$cmd '$ospp' /sethst:$script:KMS_SERVER"
    Invoke-Expression "$cmd '$ospp' /act"

    Write-Host "$nl"
    Write-Host "激活完成！！"-f green
    Write-Host "当前 $name 的激活状态:" -f green
    Write-Host "$nl"
    Invoke-Expression "$cmd '$ospp' /dstatus"
    Invoke-Expression 'Pause'

}

Main $script:MENU_INFO
