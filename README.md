# Win11Debloat

[![GitHub Release](https://img.shields.io/github/v/release/Raphire/Win11Debloat?style=for-the-badge&label=Latest%20release)](https://github.com/Raphire/Win11Debloat/releases/latest)
[![Join the Discussion](https://img.shields.io/badge/Join-the%20Discussion-2D9F2D?style=for-the-badge&logo=github&logoColor=white)](https://github.com/Raphire/Win11Debloat/discussions)
[![Static Badge](https://img.shields.io/badge/Documentation-_?style=for-the-badge&logo=bookstack&color=grey)](https://github.com/Raphire/Win11Debloat/wiki/)

Win11Debloat 是一款轻量级、易于使用的 PowerShell 脚本，让您可以快速清理和自定义 Windows 体验，无需安装！您可以使用它来移除预装应用、禁用遥测、移除侵入性界面元素等等。无需费力地逐一检查所有设置或逐个移除应用。Win11Debloat 让整个过程变得快速而简单！

该脚本还包含许多系统管理员和高级用户会喜欢的功能，例如强大的命令行界面、Windows 审核模式支持以及对其他 Windows 用户进行更改的能力。您还可以轻松导出和导入首选设置，从而快速在所有系统上应用相同的设置。更多详情请参阅我们的 [Wiki](https://github.com/Raphire/Win11Debloat/wiki)。

![Win11Debloat 菜单](/Assets/Images/menu.png)

#### 这个脚本对您有帮助吗？请考虑支持一下

[![爱发电](https://img.shields.io/badge/爱发电-支持作者-946ce6?style=for-the-badge)](https://afdian.com/a/qingxwa)

## 使用方法

> [!Warning]
> 我们已尽最大努力确保此脚本不会意外破坏任何操作系统功能，但使用风险自负！如果遇到任何问题，请在[此处](https://github.com/Raphire/Win11Debloat/issues)报告。

### 快捷方式

通过 PowerShell 自动下载并运行脚本。

1. 打开 PowerShell 或终端。
2. 将以下命令复制并粘贴到 PowerShell 中：

```PowerShell
irm https://raw.githubusercontent.com/scavin/Win11Debloat/master/Get_CN.ps1 | iex
```

3. 等待脚本自动下载并启动 Win11Debloat。
4. 仔细阅读并按照屏幕上的说明操作。

此方式支持命令行参数来自定义脚本行为。更多信息请点击[此处](https://github.com/Raphire/Win11Debloat/wiki/Command%E2%80%90line-Interface#parameters)。

### 传统方式

<details>
  <summary>手动下载并运行脚本。</summary><br/>

  1. [下载最新版本的脚本](https://github.com/Raphire/Win11Debloat/releases/latest)，并将 .ZIP 文件解压到您想要的位置。
  2. 导航到 Win11Debloat 文件夹。
  3. 双击 `Run.bat` 文件启动脚本。注意：如果控制台窗口立即关闭且没有任何反应，请尝试下面的高级方式。
  4. 接受 Windows UAC 提示以管理员身份运行脚本，这是脚本正常工作所必需的。
  5. 仔细阅读并按照屏幕上的说明操作。
</details>

### 高级方式

<details>
  <summary>手动下载脚本并通过 PowerShell 运行。推荐高级用户使用。</summary><br/>

  1. [下载最新版本的脚本](https://github.com/Raphire/Win11Debloat/releases/latest)，并将 .ZIP 文件解压到您想要的位置。
  2. 以管理员身份打开 PowerShell 或终端。
  3. 输入以下命令临时启用 PowerShell 执行策略：

  ```PowerShell
  Set-ExecutionPolicy Unrestricted -Scope Process -Force
  ```

  4. 在 PowerShell 中，导航到文件解压的目录。例如：`cd c:\Win11Debloat`
  5. 输入以下命令运行脚本：

  ```PowerShell
  .\Win11Debloat.ps1
  ```

  6. 仔细阅读并按照屏幕上的说明操作。

  此方式支持命令行参数来自定义脚本行为。更多信息请点击[此处](https://github.com/Raphire/Win11Debloat/wiki/Command%E2%80%90line-Interface#parameters)。
</details>

## 功能概述

以下是 Win11Debloat 提供的主要功能和特性的概述。您可以访问 [Wiki](https://github.com/Raphire/Win11Debloat/wiki) 了解更多详情。

> [!Tip]
> Win11Debloat 所做的所有更改都可以轻松恢复，几乎所有应用都可以通过 Microsoft Store 重新安装。您可以访问 [Wiki](https://github.com/Raphire/Win11Debloat/wiki/Reverting-Changes) 了解有关恢复更改的更多信息。

#### 应用卸载

- 移除各种预装应用。点击[此处](https://github.com/Raphire/Win11Debloat/wiki/App-Removal)了解更多信息。

#### 隐私与建议内容

- 禁用遥测、诊断数据、活动历史记录、应用启动跟踪和定向广告。
- 禁用 Windows、锁屏和 Microsoft Edge 中的提示、技巧、建议和广告。
- 禁用 Windows 定位服务、应用位置访问和「查找我的设备」位置跟踪。
- 隐藏设置「主页」页面上的 Microsoft 365 广告，或完全隐藏「主页」页面。

#### AI 功能

- 禁用并移除 Microsoft Copilot、Windows Recall 和 Click To Do。
- 阻止 AI 服务（WSAIFabricSvc）自动启动。
- 禁用 Edge、画图和记事本中的 AI 功能。

#### 系统

- 禁用共享和移动文件的「拖拽托盘」。
- 恢复旧版 Windows 10 样式的右键菜单。
- 关闭「提高指针精确度」（鼠标加速）。
- 禁用粘滞键快捷键。
- 禁用存储感知自动磁盘清理。
- 禁用快速启动以确保完全关机。
- 禁用 BitLocker 自动设备加密。
- 禁用现代待机期间的网络连接以减少电池消耗。

#### Windows 更新

- 阻止 Windows 在更新可用后立即获取。
- 阻止登录后更新自动重启。
- 禁用与其他电脑共享下载的更新（传递优化）。
- 阻止 Windows 自动安装 LG Monitor App、Alienware Command Center 等设备配套应用。

#### 外观

- 为系统和应用启用深色模式。
- 禁用透明效果、动画和视觉效果。

#### 开始菜单与搜索

- 自定义开始菜单：移除已固定应用、隐藏推荐内容、自定义「所有应用」部分。
- 禁用开始菜单中的手机连接移动设备集成。
- 禁用 Windows 搜索中的 Bing 网页搜索、Copilot 集成和 Microsoft Store 应用建议。

#### 任务栏

- 更改任务栏对齐方式。
- 自定义或隐藏任务栏按钮，如搜索栏、任务视图等。
- 禁用任务栏和锁屏上的小组件。
- 在任务栏右键菜单中启用「结束任务」选项以快速强制关闭应用。
- 为任务栏应用区域启用「最后活动点击」行为。这允许您重复点击任务栏中的应用图标来切换该应用打开窗口之间的焦点。
- 自定义任务栏上应用按钮的显示方式。

#### 文件资源管理器

- 更改文件资源管理器的默认打开位置。
- 显示已知文件类型的文件扩展名。
- 显示隐藏的文件、文件夹和驱动器。
- 从文件资源管理器导航窗格中隐藏主页、图库或 OneDrive 部分。
- 从文件资源管理器导航窗格中隐藏重复的可移动驱动器条目，仅保留「此电脑」下的条目。
- 将所有常用文件夹（桌面、下载等）添加回文件资源管理器中的「此电脑」。
- 更改文件资源管理器中驱动器盘符的位置或可见性。

#### 多任务

- 禁用窗口贴靠。
- 禁用拖拽或贴靠窗口时的贴靠助手和贴靠布局建议。
- 更改贴靠窗口或按 Alt+Tab 时是否显示标签页。

#### 可选 Windows 功能

- 启用 Windows 沙盒，一个用于在隔离环境中安全运行应用程序的轻量级桌面环境。
- 启用 Windows 子系统 for Linux，允许您直接在 Windows 上运行 Linux 环境。

#### 其他

- 禁用 Xbox Game Bar 集成和游戏/屏幕录制。如果您卸载了 Xbox Game Bar，这也会禁用 `ms-gamingoverlay`/`ms-gamebar` 弹窗。
- 禁用 Brave 浏览器中的膨胀功能（AI、加密货币、新闻等）。

#### 高级功能

- 支持[对其他用户应用更改](https://github.com/Raphire/Win11Debloat/wiki/Advanced-Features#running-as-another-user)，而不是当前登录的用户。
- [Sysprep 模式](https://github.com/Raphire/Win11Debloat/wiki/Advanced-Features#sysprep-mode)可将更改应用于 Windows 默认用户配置文件，确保所有新用户都会自动应用这些更改。

## 贡献

欢迎各种形式的贡献！请查看我们的[贡献指南](https://github.com/Raphire/Win11Debloat/blob/master/.github/CONTRIBUTING.md)以获取详细的入门说明和最佳实践。

## 许可证

Win11Debloat 使用 MIT 许可证。有关更多信息，请参阅 LICENSE 文件。
