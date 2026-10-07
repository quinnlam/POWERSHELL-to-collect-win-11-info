# POWERSHELL-to-collect-win-11-info
POWERSHELL脚本 收集电脑V2


Get-LocalSystemReport.ps1 — 功能说明
用途：在一般用户帐号（不需要系统管理员权限）下执行，收集本机系统信息， 产生一份直向(Portrait)、适合打印的单一 HTML 报告 （报告文件名固定为 Local Computer Report.html）。

基本参数
参数	说明	默认值
-OutputPath	报告输出目录	脚本自身所在目录（$PSScriptRoot）
更新历史筛选版脚本文件名：Get-LocalSystemReport-QualityFeatureUpdates.ps1。下方原始执行示例保留；运行更新版时请替换为此文件名，也可以将更新版重命名为 Get-LocalSystemReport.ps1。
执行方式：
.\Get-LocalSystemReport.ps1

或：

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Get-LocalSystemReport.ps1


报告头部
•	标题：Local Computer Report
•	信息行（粗体显示）：Computer: ... | Generated: ... | Collected by: ... | IP Address: ... 
o	IP Address 栏位包含每张网卡的 IP / 子网掩码 / 网关 / DNS，多张网卡用 ;; 分隔。
•	签名栏：信息行下方提供 5 个签名栏位，供列印后手写签名： DGM　CISO　ISO　IT MANAGER　IT
•	区块显示/隐藏勾选面板：列出全部 9 个区块，预设全部打勾（= 全部显示）； 取消勾选某区块即可把该区块从报告中隐藏。此面板本身列印时会自动隐藏， 有 "Show All / Hide All" 快捷按钮。

报告内容（共 9 个区块，依序分页打印）
1.	About This PC 计算机名称、厂牌型号、操作系统版本/组建、CPU、内存、域/工作组、BIOS 序号、 开机时间、目前登入用户等。
2.	Local User Accounts 列出本机每一个使用者帐号的名称、所属的本机群组、启用/停用状态。
所属群组是直接针对每个用户本身查询（ADSI IADsUser.Groups()）， 不是反过来枚举所有群组成员再做名字配对，避免因配对失败或群组内有 无法解析的帐户而导致整列空白。
3.	Administrators Group Members 本机 Administrators 群组内的所有成员（名称、类型、来源）。
4.	BitLocker Status 每一个磁盘的 BitLocker 加密状态（Volume Status）与保护状态（Protection Status，On/Off）。不含加密百分比（Encrypted %，已依需求移除）。
5.	Antivirus Status 侦测到的防毒软件名称、启用状态；若为 Windows Defender，额外显示即时防护 状态与病毒码版本/最新更新时间。
6.	Installed Software 已安装软件清单（名称、版本、发布者、安装日期），来源为登录表 Uninstall 清单。
7.	Shared Folders
o	本机所有共享文件夹（名称、本机路径、描述）。
o	每个共享的共享权限明细（帐户名称、Allow/Deny、权限等级： Full / Change / Read）。
o	若目前帐户权限不足以查询共享权限明细，该共享仍会显示在清单中，权限栏 会注明"可能需要系统管理员权限"。
8.	Screen Saver Settings
o	屏幕保护程序是否启用（Enabled/Disabled）。
o	等待几分钟后自动启动（Wait Time）。
o	"在恢复时显示登录屏幕"是否勾选（On Resume, Display Logon Screen — 恢复时是否需要重新登录）。

9. Windows Quality / Feature Updates (Latest 10)
收集本机 Windows 更新历史中符合筛选条件的质量更新与功能更新，按时间由新到旧显示，合计最多 10 条；不是两种类型各显示 10 条。
显示范围：Windows 累积更新、安全更新、预览更新、服务堆栈更新；.NET / .NET Framework 安全更新及累积更新；标题可识别的 Windows 功能更新和启用包更新。
排除范围：Microsoft Defender Antivirus 安全智能更新、Defender 平台更新、恶意软件删除工具更新，以及不符合质量更新或功能更新标题规则的记录（例如一般驱动更新）。此排除仅影响第 9 节；第 5 节 Antivirus Status 仍保留 Defender 防护状态、病毒码版本和更新时间。
显示字段：Date / Time（日期时间）、Type（Quality / Feature）、Update Title（完整标题）、KB（从标题提取的编号）、Operation（Installation / Uninstallation）、Result（执行结果）。标题中原有的月份、预览标识和 Build 编号会保留；未包含的版本信息不会自行补入。
统计逻辑：使用 Windows Update Agent 的 Microsoft.Update.Session COM 接口查询本机历史。每批读取最多 100 条，先排除 Defender 等无关记录，再取最近 10 条匹配记录；不足 10 条时显示实际数量，不用其他类型凑数。
历史记录范围：保留成功、失败、已中止、安装和卸载等事件；同一 KB 的不同尝试不会去重。Result 可显示 Not started、In progress、Succeeded、Succeeded with errors、Failed、Aborted 或 Unknown。
报告交互：新增第 9 个显示／隐藏复选框，默认勾选；支持 Show All / Hide All，并在打印时作为独立分页区块显示。
查询异常：没有匹配记录时显示提示；接口访问或历史查询失败时显示错误原因，不中断其他报告区块。该功能不调用在线更新扫描、下载或安装方法，不依赖第三方 PowerShell 模块。
分类限制：当前按更新标题关键词筛选，并非直接读取 Windows 设置页面的“质量更新／功能更新”分组。特殊标题、语言差异或未被规则覆盖的更新可能漏选，也可能存在误分类；报告不承诺与设置页面完全一致。
用途限制：此区块是更新历史摘要，不是当前已安装补丁清单，也不是补丁合规结论。标题和日期均取自本机返回记录；界面表头为英文，但完整更新标题保留原语言。
设计原则 / 技术限制
•	设计目标是在一般用户权限下运行，不强制要求「以系统管理员身份执行」；所有信息来源都 尽量挑选一般用户可读取的管道（CIM/WMI、Get-LocalUser、ADSI (WinNT Provider)、Get-BitLockerVolume、SecurityCenter2 WMI、Windows Defender Cmdlet、登录表 Uninstall 清单、Get-SmbShare、HKCU:\Control Panel\Desktop）。 新增更新历史查询使用 Windows Update Agent COM 接口；实际可读取范围仍受系统环境和访问权限限制。
•	任何一个区块若因权限不足、功能未安装、或系统不支援而无法取得信息，会在该 区块显示原因说明，不会让整份脚本执行中断。
•	报告为纯 HTML 单一文件，不依赖任何第三方 PowerShell 模块，双击即可用 浏览器打开；也可以用 Excel 直接打开该 .html 档。
•	HTML 固定标题、表头和状态文字使用英文；来自系统的动态内容（例如 Windows 更新完整标题）保留原语言。PowerShell 脚本内部注解与终端机提示信息维持中文，方便维护者阅读。
•	CSS 已设定 @page { size: portrait; }，并对各大区块加了分页规则 （page-break-before: always），方便直接用浏览器「打印」产生直向、分页 清楚的纸本报告；签名栏与信息行会一起印出，区块勾选面板打印时自动隐藏。

已知限制（供未来扩充参考）
•	共享文件夹的权限是「共享层级权限」，不是 NTFS 权限；实际存取权限通常是 两者取交集（较严格者生效），如需要 NTFS 权限，可参考另一支脚本 Get-FolderPermissions.ps1（专门处理 NTFS 权限矩阵报告）。
•	部分系统上 Get-SmbShareAccess 可能仍需要系统管理员权限才能取得明细。
•	区块显示/隐藏的勾选状态只存在于浏览器当前分页，重新整理页面会恢复成 脚本产生当下的预设状态（全部显示），不会被永久保存。
•	本脚本不含 Microsoft Outlook 相关检测（先前版本曾实作过 Reading Pane / 自动已读回执检测，后续已依需求整段移除）。

此文件仅为功能说明，供之后延续开发/修改需求时快速对照使用，非正式规格书。

