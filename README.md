# KMS 激活脚本

#### 不提倡使用盗版软件！！！

激活 Windows 的脚本只支持激活批量授权版（VOL 版）的 Windows, 零售版（Retail）的
Windows 请另寻它法。

Office 方面, 这里提供了两种激活 Office 的脚本, Retail 版和 VOL 版, 分别对应两种
版本 Office, 零售版（Retail 版）和批量授权版（VOL 版）。

零售版的 Office 无法使用 KMS 直接激活, 要先转换为批量授权版的 Office, 再激活。

**零售版的 Office 和批量授权版的 Office 在功能上没有任何区别, 只是激活方式不同
。**

KMS 激活服务器地址: `kms.mogeko.me`

服务器挂掉了也不用担心, 只需要去 Google/百度 搜索可以用的 KMS 服务器地址, 然后替
换掉原脚本中挂掉的服务器地址（`kms.mogeko.me`）就可以了。

## 使用方法

下载并安装 CA 证书: [`mogeko.cer`](https://github.com/Mogeko/kms/releases/download/v2.1.2/mogeko.cer)

下载并运行 [`Run.ps1`](https://github.com/Mogeko/kms/releases/download/v2.1.2/run.ps1), 选择合适的服务即可激活相应的软件。

[为什么需要安装 CA 证书?](https://github.com/Mogeko/kms/wiki/%E4%B8%BA%E4%BB%80%E4%B9%88%E8%A6%81%E5%AE%89%E8%A3%85-CA-%E8%AF%81%E4%B9%A6%EF%BC%9F)

### 判断你安装的 Office 是哪个版本

用管理员权限打开命令行（CMD）, 来到 Office 的安装目录, 输入命令
：`cscript ospp.vbs /dstatus`

如果输出的信息中包含下面这句话说明你安装的是零售版

```
LICENSE DESCRIPTION: Office 15, RETAIL(Grace) channel
```

如果输出的信息中包含下面这句话说明你安装的是批量授权版

```
LICENSE DESCRIPTION: Office 15, VOLUME_KMSCLIENT channel
```

## data.json

储存激活用的密钥等信息的 [`data.json`](https://mogeko.github.io/kms/data.json)
部署在本项目的 Github Pages 上，在每次启动时通过网络动态获取。

如果你不希望动态获取 `data.json` 或者想要自定义 `data.json`，你可以将
[`data.json`](https://mogeko.github.io/kms/data.json) 下载下来，放在 `run.ps1`
脚本的同目录下 **(不推荐这样做)**

## LICENCE

[**GNU General Public License v3.0**](https://github.com/Mogeko/KMS/blob/master/LICENSE)
