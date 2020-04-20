# KMS 激活脚本
#### 不提倡使用盗版软件！！！

激活 Windows 的脚本只支持激活批量授权版（VOL 版）的Windows, 零售版（Retail）的 Windows 请另寻它法。

Office 方面, 这里提供了两种激活 Office 的脚本, Retail 版和 VOL 版, 分别对应两种版本 Office, 零售版（Retail 版）和批量授权版（VOL 版）。

零售版的 Office 无法使用 KMS 直接激活, 要先转换为批量授权版的 Office, 再激活。

**零售版的 Office 和批量授权版的 Office 在功能上没有任何区别, 只是激活方式不同。**

KMS 激活服务器地址: `kms.mogeko.me`

服务器挂掉了也不用担心, 只需要去 Google/百度 搜索可以用的 KMS 服务器地址, 然后替换掉原脚本中挂掉的服务器地址（`kms.mogeko.me`）就可以了。

## 使用方法

双击 `Run.ps1`, 选择合适的服务即可激活相应的软件。

### 判断你安装的 Office 是哪个版本

用管理员权限打开命令行（CMD）, 来到Office的安装目录, 输入命令：`cscript ospp.vbs /dstatus `

如果输出的信息中包含下面这句话说明你安装的是零售版

```
LICENSE DESCRIPTION: Office 15, RETAIL(Grace) channel
```

如果输出的信息中包含下面这句话说明你安装的是批量授权版

```
LICENSE DESCRIPTION: Office 15, VOLUME_KMSCLIENT channel
```

## 脚本介绍

简单介绍一下各个脚本:

```
.
├── 将零售版转换为批量激活版
| ├── 查看 Office 状态.cmd     # 查看 Office 状态
| ├── Office_2010.cmd         # 将零售版 Office 2010 转换为批量激活版
| ├── Office_2013.cmd         # 将零售版 Office 2013 转换为批量激活版
| ├── Office_2016.cmd         # 将零售版 Office 2016 转换为批量激活版
| ├── Office_2019.cmd         # 将零售版 Office 2019 转换为批量激活版
├── LICENSE                   # LICENSE
├── README.md                 # README
└── Run.ps1                   # 激活脚本
```

## LICENCE

[**GNU General Public License v3.0**](https://github.com/Mogeko/KMS/blob/master/LICENSE)