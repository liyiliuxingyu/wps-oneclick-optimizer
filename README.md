# WPS Office 一键优化脚本

> 适用：Windows 10/11 + WPS Office 中国个人版（12.x 实测：12.1.0.29690，安装于自定义目录同样可用）
> 实测环境：Windows 11，WPS Office 12.1.0.29690（个人版 / 冬季更新）
>
> **只需要一个文件：`WPS一键优化.bat`**。双击运行 → 中文菜单选 1/2/3/0；优化、检测、恢复三合一，全部逻辑内嵌在这一个 bat 里（用记事本打开可看到并修改里面的 PowerShell 代码）。

> 📦 **打包下载（zip）**：[Releases · v2.2](https://github.com/liyiliuxingyu/wps-oneclick-optimizer/releases/tag/v2.2)（`WPS-Optimize.zip`）。

---

## 一、这个脚本做什么

| # | 目标 | 实现方式 | 生效范围 |
|---|------|----------|----------|
| 1 | **关闭截图中的 13 个“功能定制”项**（3 个推荐项 + 10 个系统右键菜单隐藏项） | 调用官方 **WPS Office 配置工具（ksomisc.exe）→ 功能定制**，用 Windows UI 自动化自动勾选并点“确定”（等价人工操作） | 关闭 WPS 首页推荐、右键菜单中的 WPS 扩展/新建拓展/云文档/桌面拓展/回收站/解压/释放C盘空间/清理重复文件/编辑/AI 等入口 |
| 2 | **不登录也能用（跳过强制登录）** | 写入 WPS 官方“强制登录”开关：`enableForceLoginForFirstInstallDevice=false`、`enableForceLoginForHasInstallDevice=false` | 减少/跳过“首次安装设备强制登录”“已安装设备强制登录”弹窗，离线可用本地功能 |
| 3 | **关闭 WPS 自启动项** | ① 禁用计划任务 `WpsUpdateTask_*`、`WpsUpdateLogonTask_*`、`WpsWakeWnsLogonTask*`；② `wpscloudsvr`（WPS Office Cloud Service）停止并设为“禁用”；③ 清理注册表 Run/RunOnce 与启动文件夹中的 WPS 项 | 开机不再自动检查更新、不再后台常驻 WPS 云服务 |
| 4 | **关闭 WPS 网盘 / 文件传输助手自启动** | 关闭 `HKCU\Software\Kingsoft\wpscloud` 下网盘/同步/传输相关开关，`miniTransferAssistant\globalEnable=false`，配合云服务禁用 | WPS 网盘、文件传输助手不再随开机常驻、不再自动同步 |
| 5 | **关闭资源管理器里的 WPS 入口** | 删除 WPS 云盘的资源管理器命名空间注册（「此电脑」和桌面的「WPS云盘」图标），并删除 WPS 云盘右键扩展（`qingshellext` / `qingnse` 等 Handler） | 资源管理器不再出现「WPS云盘」图标与云盘右键入口（此电脑 / 桌面 / 右键菜单） |
| 6 | **关闭启动时自动打开的「稻壳AI模板」页** | 写入 WPS 官方开关 `HKCU\Software\Kingsoft\Office\6.0\plugins\kdocerpage\docerTabShow = 0`（REG_DWORD） | 每次启动 WPS 时不再自动打开「稻壳AI模板」标签页（实测重启 WPS 后仍生效） |
| 7 | **关闭 WPS 启动首页右侧面板** | 写入 WPS 官方开关 `HKCU\Software\Kingsoft\Office\6.0\plugins\kstartpage\rightAppInfo\rightPanelWidgetVisibleGlobal = "false"`（REG_SZ） | WPS 首页右侧的「精选推荐 / 功能推荐」面板不再显示（实测重启 WPS 后仍生效） |

另外脚本还有：

- **自动备份**：注册表、计划任务 XML、服务启动类型、Run 启动项、启动文件夹快捷方式，全部备份到脚本目录 `backup\wps-opt-<时间>\`；
- **一键恢复**：菜单选 `3` 即可从**最早的一次备份（优化前的原始状态）**完整还原（包括 13 项勾选状态，因为其开关就在 `HKCU\Software\Kingsoft` 备份中；资源管理器入口等会从后续备份补齐；即使多次 Apply 也能恢复到原始状态）；
- **只读检测**：菜单选 `2`，不改任何东西，只输出当前状态；
- **幂等**：重复运行不会重复破坏，已经是目标状态会提示“已是关闭状态”。

---

## 二、怎么用（推荐步骤）

**只需要这一个文件：`WPS一键优化.bat`**（双击即可，不需要其他文件）。

1. **先保存并关闭 WPS 文档**（脚本会检测 WPS 是否在运行并提示；不关也能跑，但关掉更稳）。
2. 双击 **`WPS一键优化.bat`**，出现中文菜单：
   - `1` 一键优化（会弹出 UAC 提权，点“是”）
   - `2` 只读检测（不改任何东西，不需要管理员）
   - `3` 一键恢复（从最早备份恢复原始状态，会弹 UAC）
   - `0` 退出
3. 选 `1` 后等待完成，最后打印汇总表；按回车可回到菜单，再选 `0` 退出。
4. 右键菜单 / 资源管理器图标类改动，**重启资源管理器或重启电脑**后完全生效。

> 也支持命令行直接调用（不进菜单）：
> - 检测：`WPS一键优化.bat -Mode Scan`
> - 执行：`WPS一键优化.bat -Mode Apply`
> - 恢复：`WPS一键优化.bat -Mode Revert`
> - 可选参数：`-RestartExplorer`（执行完自动重启资源管理器）、`-SkipUiAutomation`、`-SkipExplorerEntries`、`-SkipService`、`-SkipScheduledTasks`、`-SkipCloud`、`-SkipLogin`
>
> **WPS 安装位置自动查找**：脚本会依次从 注册表安装信息 → 卸载项 → 文件关联 → 运行中的 WPS 进程 → 系统服务 → 开始菜单快捷方式 → PATH(`where.exe`) → 固定磁盘浅层扫描 中定位 `ksomisc.exe`，不写死安装路径。
> 如果 WPS 装在非常规目录、或注册表被清理导致自动找不到，可用 `-WpsPath` 手动指定（安装目录 / office6 目录 / ksomisc.exe 任选其一）：
> `WPS一键优化.bat -Mode Apply -WpsPath "D:\Apps\WPS Office\12.1.0.29690"`

---

## 三、文件说明

| 文件 / 目录 | 作用 |
|------|------|
| **`WPS一键优化.bat`** | **唯一需要的文件**：双击运行，内置 优化 / 检测 / 恢复 三合一菜单；完整 PowerShell 逻辑以文本形式内嵌在文件末尾，可用记事本查看和修改 |
| `README.md` | 本说明（可选，方便查阅） |
| `backup\wps-opt-<时间>\` | 每次优化自动生成：注册表备份、计划任务 XML、服务启动类型、资源管理器入口备份、manifest.json |
| `logs\WPS-Optimize_<时间>.log` | 每次运行完整日志 |

---

## 四、原理与实测依据（重要）

WPS 12.x 的“功能定制”页 13 项中，**大部分状态并不是普通注册表值**，而是 WPS 的 **KIC 加密开关**（保存在 `HKCU\Software\Kingsoft\Office\6.0\Common\kic\switch\user`，值为加密串）。实测“取消勾选→确定”与“重新勾选→确定”后对比注册表，可以确认：

- 只有这几项有**明文开关**（脚本在 UI 自动化失败时也会写入它们作为降级）：
  - 隐藏WPS新建拓展菜单项 → `HKCU\Software\Kingsoft\Office\6.0\Common\ksomisc\VasContextMenuShellNewUser = false`，并删除新建菜单注册项 `HKCU\Software\Classes\.pdfwpsshellnew`、`.pptxwpsaicreateshellnew`、`WPP.PPTX.AI`；
  - 隐藏编辑拓展菜单项 → `...\Office\6.0\kwpsshellext\contextmenu\photo\supportedit = false`；
  - 隐藏WPS清理重复文件功能菜单项 → `...\Office\6.0\plugins\kcdiskclean\contextMenuVisible = 0`；
  - 关闭场景化功能推荐 → `...\Office\6.0\plugins\kcdiskclean\wpsSceneRecEnable = false`；
  - 文档属性扩展页 → `...\Office\6.0\kwpsshellext\disableDocPropPage = 1`。
- 其余项（扩展功能、云文档、桌面拓展、回收站、解压、释放C盘空间、WPS AI、模板/在线素材推荐）由 **KIC 加密开关**控制，**无法用脚本直接伪造**，否则会写坏状态。因此脚本采用“**驱动官方配置工具 UI 自动勾选**”的方式——和用户手工点配置工具完全等价，这也是最稳妥的做法。

实测验证：脚本会自动勾选 13 项并点“确定”，然后**重新打开配置工具复核**，汇总表里会打印每项结果；正常应显示 13/13 已勾选。

其他实测到的位置：

- 免登录：`HKCU\Software\Kingsoft\Office\6.0\plugins\officespace\flogin`
  - `enableForceLoginForFirstInstallDevice`（REG_SZ，"false" = 关闭）
  - `enableForceLoginForHasInstallDevice`（REG_SZ，"false" = 关闭）
- WPS 网盘/传输助手：`HKCU\Software\Kingsoft\wpscloud`
  - `fileTransfer` / `UploadSync` / `DesktopSync` / `SendBy` / `SpecifyFolderSync` / `uploadFile` / `shareFile`
  - `kicFileTransferSwitch` / `kicUploadFileSwitch` / `kicDesktopSyncSwitch` / `kicShareFileSwitch` / `kicSendBySwitch` / `kicUploadSyncSwitch`
  - `miniTransferAssistant\globalEnable` / `defaultEnable`（文件传输助手总开关）
- 自启动：
  - 计划任务：`WpsUpdateTask_<用户名>`、`WpsUpdateLogonTask_<用户名>`（动作：`ksolaunch.exe /wpsupdate -from=task -source=task`）
  - 服务：`wpscloudsvr`（显示名 "WPS Office Cloud Service"，二进制在 `C:\Program Files (x86)\Kingsoft\office6\wpscloudsvr.exe`）
  - 本机干净安装默认**没有** Run/RunOnce/启动文件夹项，脚本仍会兜底扫描清理。

---

### WPS 安装路径自动发现（不写死路径）

脚本定位 `ksomisc.exe` 的顺序（找到即停，并在日志/检测结果里打印“发现方式”）：

1. `-WpsPath` 手动指定；
2. 注册表 `HKCU/HKLM(32/64)\Software\Kingsoft\Office\6.0\Common` 的 `InstallRoot`；
3. 卸载项 `InstallLocation / UninstallString / DisplayIcon`（DisplayName 含 WPS Office/Kingsoft Office/WPS）；
4. 文件关联（`KWPS.Document.9`、`WPS.Doc.6` 等的 `shell\open\command`）与 `App Paths\wps.exe`；
5. 正在运行的 `wps.exe / et.exe / wpp.exe / ksomisc.exe` 进程路径；
6. 系统服务（名字/显示名含 WPS、Kingsoft）的 `ImagePath`；
7. 开始菜单快捷方式（名称含 WPS/Kingsoft/金山）的指向；
8. `where.exe wps.exe / ksomisc.exe / wpsoffice.exe`；
9. 最后扫描固定磁盘：常见目录 + 各磁盘深度 ≤2 下名称含 `WPS/Kingsoft/金山` 的目录，以及 `apps/software/Programs` 等父目录。

仍然找不到时，检测结果会列出“已尝试的查找方式”，并提示用 `-WpsPath` 手动指定。

### 「稻壳AI模板」启动页的实测位置（12.1.0.29690）

- 开关：`HKCU\Software\Kingsoft\Office\6.0\plugins\kdocerpage\docerTabShow`（REG_DWORD：**0 = 不再默认展示**，1 或不存在 = 每次启动默认展示）
- 同键下还有 `docerPageTitlePreview = 稻壳AI模板`、`docerPageTitleVisible = true`，可辅助确认这是「稻壳AI模板」页。
- 实测路径：在 WPS 标题栏「稻壳AI模板」标签上右键 → 点「不再默认展示」后，该值立即变为 0、标签消失；重启 WPS 后不再出现。脚本直接写这个值，效果等同于点这一下。

### WPS 启动首页右侧面板的实测位置（12.1.0.29690）

- 开关：`HKCU\Software\Kingsoft\Office\6.0\plugins\kstartpage\rightAppInfo\rightPanelWidgetVisibleGlobal`（REG_SZ：**"false" = 不显示右侧面板**，"true" = 显示）
- 实测路径：WPS 首页右侧面板标题「精选推荐」右上角点 ×（界面提示“可在此重新打开右侧面板”）后，该值变为 "false"、面板消失；重启 WPS 后仍保持关闭。脚本直接写这个值，效果等同于点 ×。
- 面板内容（今日签到、WPS Office 焕新升级、功能推荐等）属于 WPS 首屏推广，关闭后不影响 WPS 新建/打开文档等本地功能。

### 资源管理器 WPS 入口的实测位置（12.1.0.29690）

- 「此电脑」里的 WPS云盘：`HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{EEEEFCF7-867B-4FA2-9ABD-884CF531B600}`
- 桌面上的 WPS云盘：`HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{EEEEFCF7-867B-4FA2-9ABD-884CF531B602}`
- 这两个命名空间指向 `qingnse.dll`，目标文件夹是 `Documents\WPSDrive`；脚本只删除“入口”注册，**不会删除你本地的 WPSDrive 文件夹和里面的文件**。
- 云盘右键扩展：`HKCU\Software\Classes\{*|Directory|Directory\Background|lnkfile}\shellex\ContextMenuHandlers\{Open With qingshellext|QingNseContextMenu}` 等，脚本会动态枚举并备份后删除。
- 删除前脚本会把每个入口的注册表项导出到 `backup\wps-opt-<时间>\explorer-*.reg`，并记录在 `explorer-entries.json`，可用 `Revert` 一键恢复。

## 五、验证清单（运行后可自查）

1. 打开任意文件夹 → 右键：WPS 的“扩展功能/新建拓展/云文档/桌面拓展/回收站/解压/释放C盘空间/清理重复文件/编辑/WPS AI”等项应消失或大幅减少（重启资源管理器后）；
2. 打开 WPS → 右上角设置/首页：推荐、模板、在线素材、场景化推荐应关闭；
3. 打开“WPS Office 配置工具 → 功能定制”：13 项应全部勾选；
4. `任务计划程序` → 任务计划程序库：`WpsUpdateTask_*`、`WpsUpdateLogonTask_*` 状态为“已禁用”；
5. `services.msc` → `WPS Office Cloud Service`：启动类型“禁用”、状态“已停止”；
6. `任务管理器 → 启动应用`：无 WPS 相关项（脚本运行后）。
7. 运行 `WPS一键优化.bat` 并选 `2`，一次性看到以上所有状态。

---

8. 打开「此电脑」和桌面：不应再出现「WPS云盘」图标；右键菜单里也没有 WPS 云盘的上传/同步/发送等入口。

9. WPS 启动后标题栏不再出现「稻壳AI模板」标签页（脚本写入 `docerTabShow=0`）。

10. WPS 启动首页右侧不再出现「精选推荐 / 功能推荐」面板（脚本写入 `rightPanelWidgetVisibleGlobal=false`）。

## 六、注意事项与限制（请务必阅读）

1. **个人版功能边界（合规说明）**：WPS 个人版的**云文档、WPS AI、在线模板、PDF 会员功能**等本身就要求登录账号/会员，脚本无法也不应“破解”这些收费功能。本脚本能做到的是：跳过**强制登录弹窗**、关闭云服务/更新/推荐常驻，让**本地编辑、打印、格式转换等免费能力无需登录即可正常使用**。
2. **WPS 升级/自愈可能恢复设置**：WPS 有自愈机制，升级后可能重建计划任务、恢复部分配置。重新运行本脚本即可；如计划任务反复被重建，说明 WPS 自动更新仍在工作，可在官方“配置工具 → 升级设置”里改为手动更新。
3. **修改注册表有风险**：脚本会先自动备份（`backup\`），并支持 `Revert` 一键恢复；但仍建议重要资料先备份后再操作。
4. **需要管理员权限**：计划任务、系统服务、HKLM 部分需要管理员；双击 `WPS一键优化.bat` 后选 `1`/`3` 会自动申请 UAC（选 `2` 检测不需要）。
5. **需要桌面会话**：功能定制自动勾选依赖 Windows UI 自动化（不能在没有交互桌面的服务/远程无桌面环境使用）；如果失败，脚本会降级写入明文开关并提示你手动在配置工具里勾选剩余项。
6. **重启资源管理器**：右键菜单类改动，重启资源管理器或重启电脑后完全生效（脚本不会自动重启资源管理器，避免打断你正在做的事）。
7. 本脚本不修改、不删除 WPS 程序文件，不涉及破解补丁，不联网下载任何内容；恢复时从本机备份还原。

---

8. **资源管理器入口可能被 WPS 重建**：WPS 启动/升级时有可能重新注册「WPS云盘」命名空间或云盘右键扩展。重跑本脚本即可；彻底避免需要保持 WPS 云服务禁用（脚本已默认做）。
9. **清理入口不会删除文件**：脚本只删注册表中的“显示入口”，`Documents\WPSDrive` 等本地文件夹及其内容不会被删除。

10. **WPS 运行时会回写部分开关**：实测启动 WPS 后 `enableForceLoginForHasInstallDevice` 会变回 `true`、`miniTransferAssistant\globalEnable` 会变回 `true`（云服务已在脚本里禁用）。这不影响「稻壳AI模板」的持久关闭（`docerTabShow` 重启后实测仍生效）；如果在意登录/传输助手开关，可在关闭 WPS 后重新运行一次脚本。

11. **脚本不写死 WPS 安装路径**：会自动搜索（见上文“WPS 安装路径自动发现”）；如果提示“未找到 ksomisc.exe”，用 `-WpsPath` 指定安装目录即可，例如 `WPS一键优化.bat -Mode Apply -WpsPath "E:\Software\WPS Office\12.1.0.xxxxx"`。

## 七、常见问题

**Q：双击 bat 一闪而过？**
A：请先运行 `WPS一键优化.bat -Mode Scan` 查看报错，或打开 `logs\` 下最新日志；若解包失败，请确认 bat 文件未被改动（它是 UTF-8 文本，且必须保持完整）。

**Q：提示“未检测到 WPS Office / 未找到 ksomisc.exe”怎么办？**
A：脚本已内置多来源自动查找（注册表、卸载项、文件关联、进程、服务、开始菜单、PATH、磁盘扫描），一般都能找到。若仍提示找不到，可手动指定：`WPS一键优化.bat -Mode Apply -WpsPath "D:\Apps\WPS Office\12.1.0.29690"`（也可以直接给 office6 目录或 ksomisc.exe 的完整路径）。

**Q：功能定制自动勾选失败/窗口没弹出来？**
A：确认不是无桌面环境、没有被安全软件拦截；也可以直接打开“开始菜单 → WPS Office → WPS Office 配置工具 → 功能定制”手动勾选 13 项，其它部分脚本已自动完成。

**Q：为什么我还需要点 UAC？**
A：禁用计划任务、禁用服务属于系统级操作，必须管理员权限；勾选功能定制与注册表 HKCU 部分本身不需要管理员。

**Q：原来的 `WPS-Optimize.ps1` / `Optimize-WPS.bat` / `Scan-WPS.bat` / `Revert-WPS.bat` 还需要吗？**
A：不需要了。本版已合并为一个 `WPS一键优化.bat`：优化、检测、恢复都在菜单里，备份和日志仍自动生成在当前文件夹。

**Q：恢复后右键菜单没变化？**
A：重启资源管理器或重启电脑即可。

---

免责声明：本脚本为本地配置优化工具，只调用 WPS 官方配置工具与官方公开的本地配置项，不包含破解、绕过授权或修改程序文件的内容。使用前请自行备份重要数据，风险自负。
