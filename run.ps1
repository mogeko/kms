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
$script:MENU_INFO = @{
    Description = @"
本脚本支持激活所有版本的 Office 和批量激活版的 Windows
输入序号来选择你想要执行的函数，如果你想同时执行几个函数，支持多选
例如，输入: 123; 则表示选择: "选项1", "选项2" 和"选项3"
"@;
    Option = (@{
        Name = "激活 Windows (批量激活版)"
        Action = "ActivateWindows"
    },
    @{
        Name = "激活 Office"
        SubMenu = @{
            Description = "请选择你 Office 的版本号"
            Option = (@{
                Name = "Office 2010"
                SubMenu = @{
                    Description = "请选择你 Office 的版本"
                    Option = (@{
                        Name = "Office Professional Plus 2010"
                        Action = "ActivateOffice"
                        key = "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB"
                        osppPath = "office14"
                    },
                    @{
                        Name = "Office Standard 2010"
                        Action = "ActivateOffice"
                        key = "V7QKV-4XVVR-XYV4D-F7DFM-8R6BM"
                        osppPath = "office14"
                    },
                    @{
                        Name = "Office Home and Business 2010"
                        Action = "ActivateOffice"
                        key = "D6QFG-VBYP2-XQHM7-J97RH-VVRCK"
                        osppPath = "office14"
                    })
                }
            },
            @{
                Name = "Office 2013"
                SubMenu = @{
                    Description = "请选择你 Office 的版本"
                    Option = (@{
                        Name = "Office 2013 Professional Plus"
                        Action = "ActivateOffice"
                        key = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
                        osppPath = "office15"
                    },
                    @{
                        Name = "Office 2013 Standard"
                        Action = "ActivateOffice"
                        key = "KBKQT-2NMXY-JJWGP-M62JB-92CD4"
                        osppPath = "office15"
                    })
                }
            },
            @{
                Name = "Office 2016"
                SubMenu = @{
                    Description = "请选择你 Office 的版本"
                    Option = (@{
                        Name = "Office 专业增强版 2016"
                        Action = "ActivateOffice"
                        key = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
                        osppPath = "office16"
                    },
                    @{
                        Name = "Office Standard 2016"
                        Action = "ActivateOffice"
                        key = "JNRGM-WHDWX-FJJG3-K47QV-DRTFM"
                        osppPath = "office16"
                    })
                }
            },
            @{
                Name = "Office 2019"
                SubMenu = @{
                    Description = "请选择你 Office 的版本"
                    Option = (@{
                        Name = "Office 专业增强版 2019"
                        Action = "ActivateOffice"
                        key = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
                        osppPath = "office16"
                    },
                    @{
                        Name = "Office 标准版 2019"
                        Action = "ActivateOffice"
                        key = "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK"
                        osppPath = "office16"
                    })
                }
            })
        }
    },
    @{
        Name = "激活 Project"
        SubMenu = @{
            Description = "请选择你 Project 的版本号"
            Option = (@{
                Name = "Project 2013"
                SubMenu = @{
                    Description = "请选择你 Visio 的版本"
                    Option = (@{
                        Name = "Project 2013 Professional"
                        Action = "ActivateOffice"
                        key = "FN8TT-7WMH6-2D4X9-M337T-2342K"
                        osppPath = "office15"
                    },
                    @{
                        Name = "Project 2013 Standard"
                        Action = "ActivateOffice"
                        key = "6NTH3-CW976-3G3Y2-JK3TX-8QHTT"
                        osppPath = "office15"
                    })
                }
            },
            @{
                Name = "Project 2016"
                SubMenu = @{
                    Description = "请选择你 Visio 的版本"
                    Option = (@{
                        Name = "Project Professional 2016"
                        Action = "ActivateOffice"
                        key = "YG9NW-3K39V-2T3HJ-93F3Q-G83KT"
                        osppPath = "office16"
                    },
                    @{
                        Name = "Project Standard 2016"
                        Action = "ActivateOffice"
                        key = "GNFHQ-F6YQM-KQDGJ-327XX-KQBVC"
                        osppPath = "office16"
                    })
                }
            },
            @{
                Name = "Project 2019"
                SubMenu = @{
                    Description = "请选择你 Visio 的版本"
                    Option = (@{
                        Name = "Project 专业版 2019"
                        Action = "ActivateOffice"
                        key = "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"
                        osppPath = "office16"
                    },
                    @{
                        Name = "Project 标准版 2019"
                        Action = "ActivateOffice"
                        key = "C4F7P-NCP8C-6CQPT-MQHV9-JXD2M"
                        osppPath = "office16"
                    })
                }
            })
        }
    },
    @{
        Name = "激活 Visio"
        SubMenu = @{
            Description = "请选择你 Visio 的版本号"
            Option = (@{
                Name = "Visio 2010"
                SubMenu = @{
                    Description = "请选择你 Visio 的版本"
                    Option = (@{
                        Name = "Visio Premium 2010"
                        Action = "ActivateOffice"
                        key = "D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ"
                        osppPath = "office14"
                    },
                    @{
                        Name = "Visio Professional 2010"
                        Action = "ActivateOffice"
                        key = "7MCW8-VRQVK-G677T-PDJCM-Q8TCP"
                        osppPath = "office14"
                    },
                    @{
                        Name = "Visio Standard 2010"
                        Action = "ActivateOffice"
                        key = "767HD-QGMWX-8QTDB-9G3R2-KHFGJ"
                        osppPath = "office14"
                    })
                }
            },
            @{
                Name = "Visio 2013"
                SubMenu = @{
                    Description = "请选择你 Visio 的版本"
                    Option = (@{
                        Name = "Visio 2013 Professional"
                        Action = "ActivateOffice"
                        key = "C2FG9-N6J68-H8BTJ-BW3QX-RM3B3"
                        osppPath = "office15"
                    },
                    @{
                        Name = "Visio 2013 Standard"
                        Action = "ActivateOffice"
                        key = "J484Y-4NKBF-W2HMG-DBMJC-PGWR7"
                        osppPath = "office15"
                    })
                }
            },
            @{
                Name = "Visio 2016"
                SubMenu = @{
                    Description = "请选择你 Visio 的版本"
                    Option = (@{
                        Name = "Visio Professional 2016"
                        Action = "ActivateOffice"
                        key = "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK"
                        osppPath = "office16"
                    },
                    @{
                        Name = "Visio Standard 2016"
                        Action = "ActivateOffice"
                        key = "7WHWN-4T7MP-G96JF-G33KR-W8GF4"
                        osppPath = "office16"
                    })
                }
            },
            @{
                Name = "Visio 2019"
                SubMenu = @{
                    Description = "请选择你 Visio 的版本"
                    Option = (@{
                        Name = "Visio 专业版 2019"
                        Action = "ActivateOffice"
                        key = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB"
                        osppPath = "office16"
                    },
                    @{
                        Name = "Visio 标准版 2019"
                        Action = "ActivateOffice"
                        key = "7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2"
                        osppPath = "office16"
                    })
                }
            })
        }
    })
}
    

function Main {
    param (
        [hashtable]$main
    )
    $flag = $true;
    while ($flag) {
        # Drew Menu
        Invoke-Expression 'cls';
        Write-Host $script:LICENCE;
        Write-Host "$nl";
        Write-Host $main.Description;
        Write-Host "$nl";
        for ($i = 0; $i -lt $main.Option.Count; $i++) {
            "    {0}) {1}" -f ($i+1), $main.Option[$i].Name;
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
            } elseif ($main.Option[$inputItem-49].SubMenu) {
                Main  $main.Option[$inputItem-49].SubMenu;
            } elseif ($main.Option[$inputItem-49].Action -eq "ActivateWindows") {
                ActivateWindows $main.Option[$inputItem-49].Name;
            } elseif ($main.Option[$inputItem-49].Action -eq "ActivateOffice") {
                ActivateOffice $main.Option[$inputItem-49].Name $main.Option[$inputItem-49].key $main.Option[$inputItem-49].osppPath;
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
