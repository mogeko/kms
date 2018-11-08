**不提倡使用盗版软件！！！**

**不提倡使用盗版软件！！！**

**不提倡使用盗版软件！！！**

激活 Windows 的脚本只支持激活批量授权版（VOL 版）的Windows, 零售版（Retail）的 Windows 请另寻它法。

Office 方面, 这里提供了两种激活 Office 的脚本, Retail 版和 VOL 版, 分别对应两种版本 Office, 零售版（Retail 版）和批量授权版（VOL 版）。

零售版的 Office 无法使用 KMS 直接激活, 要先转换为批量授权版的 Office, 再激活。

**零售版的 Office 和批量授权版的 Office 在功能上没有任何区别, 只是激活方式不同。**

---

- **V2.0.0 更新一个引导脚本 `Run.cmd`**  
  直接双击运行 `Run.cmd`, 然后选择合适的服务即可激活相应的软件

- **V3.0.0 重构整个项目, 使 `Run.cmd` 可以单独运行**  
  史诗级加强 `Run.cmd`, 只需要这一个脚本就可以激活 Office 2016 和 Windows (VOL 版), 不依赖别的脚本  
  使所用脚本解耦


# KMS 服务器可用性测试

[检测 KMS 服务器是否挂了](https://03k.org/go/kmscheck.php)

服务器挂掉了也不用担心, 只需要去 Google / 百度 搜索可以用的 KMS 服务器地址, 然后替换掉原脚本中挂掉的服务器地址（`kms.03k.org`）就可以了。

# 使用方法

双击 `Run.cmd`, 选择合适的服务即可激活相应的软件。

或者, 再目录中选择合适的脚本单独运行。

## 判断你安装的 Office 是哪个版本

用管理员权限打开命令行（CMD）, 来到Office的安装目录, 输入命令：`cscript ospp.vbs /dstatus `

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
├── Atom    # 供单独使用的脚本
| ├── Office_2016
| | ├── Office_Project_Pro.cmd    # 将零售版的 Office Project Plus 转换为批量授权版, 然后激活
| | ├── Office_Project_Std.cmd    # 将零售版的 Office Project 转换为批量授权版, 然后激活
| | ├── Office_Retail2VOL_Only.cmd    # 仅仅将零售版的 Office 转换为批量授权版, 不激活
| | ├── Office_Retail2VOL+Activate.cmd    # 将零售版的 Office 转换为批量授权版, 然后激活
| | ├── Office_Visio_Pro.cmd    # 将零售版的 Office Visio Plus 转换为批量授权版, 然后激活
| | ├── Office_Visio_Std.cmd    # 将零售版的 Office Visio 转换为批量授权版, 然后激活
| | ├── Office_VOL_Activate.cmd    # 激活批量授权版的 Office
| | └── 查看 Office 状态.cmd    # 查看 Office 状态
| ├── Office_2013
| | └── 与 Office_2016 同理, 就不单独介绍了
| ├── Office_2010
| | └── 与 Office_2016 同理
| └── Windows_Activate.cmd    # 激活批量授权版的 Windows
├── KMS 服务可用性测试.url    # 检测 KMS 服务器是否挂了
├── Run.cmd    # 激活脚本
└── README.md
```



