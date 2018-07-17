**不提倡使用盗版软件！！！**

**不提倡使用盗版软件！！！**

**不提倡使用盗版软件！！！**

激活 Windows 的脚本只支持激活批量授权版（VOL 版）的Windows，零售版（Retail）的 Windows 请另寻它法。

Office 方面，这里提供了两种激活 Office 的脚本，Retail 版和 VOL 版，分别对应两种版本 Office，零售版（Retail 版）和批量授权版（VOL 版）。

零售版的 Office 无法使用 KMS 直接激活，要先转换为批量授权版的 Office，再激活。

**零售版的 Office 和批量授权版的 Office 在功能上没有任何区别，只是激活方式不同。**

# KMS 服务器可用性测试

[检测 KMS 服务器是否挂了](https://03k.org/go/kmscheck.php)

服务器挂掉了也不用担心，只需要去 Google / 百度 搜索可以用的 KMS 服务器地址，然后替换掉原脚本中挂掉的服务器地址（`kms.03k.org`）就可以了。

# 使用方法

## Windows 激活

双击脚本激活。（脚本需要管理员权限，双击脚本，它自己会弹出 UAC 提示授权）

## Office 激活

先安装 Office，最好安装批量授权版。

然后选择对应的脚本激活就可以了。（脚本同样需要管理员权限，同样会自己弹出 UAC 提示授权）

### 判断你安装的 Office 是哪个版本

用管理员权限打开命令行（CMD），来到Office的安装目录，输入命令：`cscript ospp.vbs /dstatus `

如果输出的信息中包含下面这句话说明你安装的是零售版

```
LICENSE DESCRIPTION: Office 15, RETAIL(Grace) channel
```

如果输出的信息中包含下面这句话说明你安装的是批量授权版

```
LICENSE DESCRIPTION: Office 15, VOLUME_KMSCLIENT channel
```

# 脚本介绍

简单介绍一下各个脚本

```
.
| ├── Office_Activate
| | ├── Component
| | | ├── ProjectPro.cmd    # 将零售版的 Office Project Plus 转换为批量授权版，然后激活
| | | ├── ProjectStd.cmd    # 将零售版的 Office Project 转换为批量授权版，然后激活
| | | ├── VisioPro.cmd    # 将零售版的 Office Visio Plus 转换为批量授权版，然后激活
| | | └── VisioStd.cmd    # 将零售版的 Office Visio 转换为批量授权版，然后激活
| | ├── Retail2VOL+Activate.cmd    # 将零售版的 Office 转换为批量授权版，然后激活
| | └── Retail2VOL_Only.cmd    # 仅仅将零售版的 Office 转换为批量授权版，不激活
| ├── VOL_Activate.cmd    # 激活批量授权版的 Office
| └── 查看 Office 状态.cmd    # 查看 Office 状态
├── Windows_Activate.cmd    # 激活批量授权版的 Windows
├── KMS 服务可用性测试.url    # 检测 KMS 服务器是否挂了
└── README.md
```



