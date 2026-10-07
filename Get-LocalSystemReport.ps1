<#
.SYNOPSIS
    在「一般使用者帳號」(不需要系统管理员权限) 下收集本机系统信息，产生一份
    适合打印的直向(Portrait) HTML 报告。

.收集内容
    1. About This PC        - 计算机名称、型号、操作系统、CPU、内存、域/工作组等
    2. Local User Accounts  - 每个本机用户帐号名称、所属的本机群组、启用/停用状态
    3. Administrators Group - 本机 Administrators 群组内的所有成员
    4. BitLocker Status     - 每个磁盘的 BitLocker 加密状态 (Volume Status) 与
                              保护状态 (Protection Status，On/Off)
    5. Antivirus Status     - 防毒软件名称、启用状态、最新病毒码更新日期
    6. Installed Software   - 已安装软件清单 (名称/版本/发布者/安装日期)
    7. Shared Folders       - 本机所有共享文件夹 (含名称/本机路径/描述)，以及每个
                              共享的「共享权限」明细 (哪个帐户/群组、允许或拒绝、
                              读取/变更/完全控制)。
    8. Screen Saver         - 屏幕保护程序是否启用、等待几分钟后自动启动、
                              以及「在恢复时显示登录屏幕」是否勾选 (恢复时是否
                              需要重新登录)。

.说明
    - 全程使用一般用户权限可读取的信息来源 (Get-CimInstance、Get-LocalUser、
      Get-LocalGroupMember、Get-BitLockerVolume、SecurityCenter2 WMI、
      Windows Defender Cmdlet、登录表 Uninstall 清单)，不需要用系统管理员帐号执行。
    - 如果某一项信息在目前的权限或系统环境下无法取得 (例如没有安装 BitLocker
      模块、没有防毒软件、登录表读取被拒绝等)，该项目会显示原因说明，而不会
      让整个脚本中断。
    - 报告是单一 HTML 文件，直向(Portrait)版面，各大区块预设分页打印，方便
      直接列印成册。
    - 共享文件夹的「共享权限」(Share Permissions，不同于 NTFS 权限) 在部分系统上
      可能需要系统管理员权限才能查询到明细；如果目前的帐户权限不足，该共享仍会
      显示在清单中 (名称/路径/描述)，但权限栏位会注明"需要系统管理员权限"。

.用法
    直接在目标计算机上执行 (不需要「以系统管理员身份执行」):
        .\Get-LocalSystemReport.ps1
    也可以指定报告输出目录 (默认为脚本自身所在目录):
        .\Get-LocalSystemReport.ps1 -OutputPath "C:\Reports"
#>

param(
    # 报表输出目录，默认是脚本自身所在目录
    [string]$OutputPath = $PSScriptRoot
)

$ErrorActionPreference = 'Continue'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

Add-Type -AssemblyName System.Web

$reportTitle = "Local Computer Report"
$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$htmlPath    = Join-Path $OutputPath "$reportTitle.html"

function Enc([string]$text) {
    if ($null -eq $text) { return "" }
    return [System.Web.HttpUtility]::HtmlEncode($text)
}

# Get-LocalGroupMember 遇到群组内有「无法解析的 SID」(例如已离职/已删除的网域帐号)
# 时，会让整个群组查询直接失败 (抛出例外)，导致该群组的所有正常成员都读不到。
# 这个函式会先尝试 Get-LocalGroupMember，失败的话改用 ADSI (WinNT Provider)
# 逐一枚举成员，对无法解析的项目只会略过该笔，不会让整个群组查询失败。
function Get-GroupMembersSafe {
    param([string]$GroupName)

    $result = New-Object System.Collections.Generic.List[PSObject]

    try {
        $members = Get-LocalGroupMember -Group $GroupName -ErrorAction Stop
        foreach ($m in $members) {
            $result.Add([PSCustomObject]@{
                Name            = $m.Name
                ObjectClass     = [string]$m.ObjectClass
                PrincipalSource = [string]$m.PrincipalSource
            })
        }
        return $result
    }
    catch {
        # 往下改用 ADSI 后备方案
    }

    try {
        $groupObj = [ADSI]"WinNT://$env:COMPUTERNAME/$GroupName,group"
        $members = @($groupObj.psbase.Invoke("Members"))
        foreach ($m in $members) {
            try {
                $adsPath   = $m.GetType().InvokeMember("ADsPath", 'GetProperty', $null, $m, $null)
                $classType = $m.GetType().InvokeMember("Class", 'GetProperty', $null, $m, $null)

                # ADsPath 格式通常是: WinNT://DOMAIN/Name 或 WinNT://COMPUTERNAME/Name
                $segments     = $adsPath -split '/'
                $shortName    = $segments[-1]
                $sourceDomain = if ($segments.Count -ge 2) { $segments[-2] } else { "" }

                $principalSource = if ($sourceDomain -eq $env:COMPUTERNAME) { "Local" } else { "ActiveDirectory" }
                $objClass = if ($classType -eq 'Group') { 'Group' } else { 'User' }

                $result.Add([PSCustomObject]@{
                    Name            = "$sourceDomain\$shortName"
                    ObjectClass     = $objClass
                    PrincipalSource = $principalSource
                })
            }
            catch {
                # 单一成员解析失败就跳过这一笔，不影响群组内其他成员
            }
        }
    }
    catch {
        # 群组本身查询失败 (例如群组不存在)，回传空清单
    }

    return $result
}

# 直接查询「某个本机用户」自己所属的群组清单 (IADsUser::Groups())。
# 这个方法是反过来从使用者本身出发去问「你在哪些群组里」，而不是枚举每个群组
# 的全部成员、再用名字去配对——后者容易因为名字格式 (是否带 "电脑名\")、
# 大小写、或其他成员解析失败而配对不到，造成 Member Of 栏位空白。
function Get-UserGroupsSafe {
    param([string]$UserName)

    $groups = New-Object System.Collections.Generic.List[string]

    try {
        $userObj = [ADSI]"WinNT://$env:COMPUTERNAME/$UserName,user"
        $groupsCollection = $userObj.psbase.Invoke("Groups")
        foreach ($g in $groupsCollection) {
            try {
                $name = $g.GetType().InvokeMember("Name", 'GetProperty', $null, $g, $null)
                if ($name) { $groups.Add($name) }
            }
            catch {
                # 单一群组名称解析失败就跳过，不影响其他群组
            }
        }
        return $groups
    }
    catch {
        # ADSI 查不到此使用者 (少见)，往下改用 Get-LocalGroupMember 逐一检查当作后备
    }

    try {
        $allGroups = Get-LocalGroup -ErrorAction Stop
        foreach ($grp in $allGroups) {
            try {
                $isMember = Get-LocalGroupMember -Group $grp.Name -Member $UserName -ErrorAction Stop
                if ($isMember) { $groups.Add($grp.Name) }
            }
            catch {
                # 该群组查询失败或此用户不在此群组内，略过
            }
        }
    }
    catch { }

    return $groups
}

Write-Host "正在收集本机系统信息，请稍候..." -ForegroundColor Cyan

# =================================================================
# 0. IP 信息 (IP / Subnet / Gateway / DNS，用于报告头部的 meta 信息行)
# =================================================================
Write-Host " - 收集 IP 信息..." -ForegroundColor DarkCyan

$ipListDisplay = "-"
try {
    $adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction Stop

    $adapterBlocks = New-Object System.Collections.Generic.List[string]

    foreach ($ad in $adapters) {
        # IPAddress / IPSubnet 是并行数组，同一个索引互相对应；只取 IPv4 (排除 IPv6 格式)
        $ipv4Entries = New-Object System.Collections.Generic.List[string]
        if ($ad.IPAddress) {
            for ($i = 0; $i -lt $ad.IPAddress.Count; $i++) {
                $ip = $ad.IPAddress[$i]
                if ($ip -match '^\d{1,3}(\.\d{1,3}){3}$') {
                    $subnet = if ($ad.IPSubnet -and $ad.IPSubnet.Count -gt $i) { $ad.IPSubnet[$i] } else { "-" }
                    $ipv4Entries.Add("$ip/$subnet")
                }
            }
        }

        if ($ipv4Entries.Count -eq 0) { continue }

        $gateway = if ($ad.DefaultIPGateway) { ($ad.DefaultIPGateway -join ", ") } else { "-" }
        $dns     = if ($ad.DNSServerSearchOrder) { ($ad.DNSServerSearchOrder -join ", ") } else { "-" }

        $block = "IP: $($ipv4Entries -join ', ')  |  Gateway: $gateway  |  DNS: $dns"
        $adapterBlocks.Add($block)
    }

    if ($adapterBlocks.Count -gt 0) {
        $ipListDisplay = ($adapterBlocks -join "  ;;  ")
    }
}
catch {
    $ipListDisplay = "Unable to retrieve"
}


# =================================================================
# 1. About This PC
# =================================================================
Write-Host " - 收集 About This PC 信息..." -ForegroundColor DarkCyan

$aboutRows = New-Object System.Collections.Generic.List[string]

try {
    $cs  = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
    $os  = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue

    $ramGB = [math]::Round(($cs.TotalPhysicalMemory / 1GB), 2)
    $domainOrWorkgroup = if ($cs.PartOfDomain) { "$($cs.Domain) (Domain)" } else { "$($cs.Workgroup) (Workgroup)" }
    $lastBoot = $os.LastBootUpTime
    $installDate = $os.InstallDate

    $aboutPairs = [ordered]@{
        "Computer Name"      = $env:COMPUTERNAME
        "Manufacturer"       = $cs.Manufacturer
        "Model"              = $cs.Model
        "Operating System"   = $os.Caption
        "OS Version / Build" = "$($os.Version) (Build $($os.BuildNumber))"
        "System Type"        = $os.OSArchitecture
        "Processor"          = $cpu.Name
        "Installed RAM"      = "$ramGB GB"
        "Domain / Workgroup" = $domainOrWorkgroup
        "BIOS Serial Number" = $bios.SerialNumber
        "OS Install Date"    = $installDate
        "Last Boot Time"     = $lastBoot
        "Current User"       = "$env:USERDOMAIN\$env:USERNAME"
    }

    foreach ($key in $aboutPairs.Keys) {
        $aboutRows.Add("<tr><td class='labelCol'>$(Enc $key)</td><td>$(Enc([string]$aboutPairs[$key]))</td></tr>")
    }
}
catch {
    $aboutRows.Add("<tr><td colspan='2' class='errCell'>Unable to retrieve system information: $(Enc $_.Exception.Message)</td></tr>")
}

# =================================================================
# 2. Local User Accounts + 3. Administrators Group Members
# =================================================================
Write-Host " - 收集本机用户帐号与群组信息..." -ForegroundColor DarkCyan

$userRows  = New-Object System.Collections.Generic.List[string]
$adminRows = New-Object System.Collections.Generic.List[string]

try {
    $localUsers = Get-LocalUser -ErrorAction Stop | Sort-Object Name
    foreach ($u in $localUsers) {
        $groupList = Get-UserGroupsSafe -UserName $u.Name
        $groups = if ($groupList.Count -gt 0) { ($groupList | Sort-Object -Unique) -join ", " } else { "-" }
        $status = if ($u.Enabled) { "Enabled" } else { "Disabled" }
        $statusClass = if ($u.Enabled) { "hasperm" } else { "noperm" }
        $userRows.Add("<tr><td class='labelCol'>$(Enc $u.Name)</td><td>$(Enc $groups)</td><td class='$statusClass'>$(Enc $status)</td></tr>")
    }
}
catch {
    $userRows.Add("<tr><td colspan='3' class='errCell'>Unable to retrieve local user accounts: $(Enc $_.Exception.Message)</td></tr>")
}

try {
    $admins = Get-GroupMembersSafe -GroupName "Administrators"
    foreach ($a in $admins) {
        $adminRows.Add("<tr><td class='labelCol'>$(Enc $a.Name)</td><td>$(Enc $a.ObjectClass)</td><td>$(Enc $a.PrincipalSource)</td></tr>")
    }
    if ($admins.Count -eq 0) {
        $adminRows.Add("<tr><td colspan='3'>No members found.</td></tr>")
    }
}
catch {
    $adminRows.Add("<tr><td colspan='3' class='errCell'>Unable to retrieve Administrators group members: $(Enc $_.Exception.Message)</td></tr>")
}

# =================================================================
# 4. BitLocker Status
# =================================================================
Write-Host " - 收集 BitLocker 状态..." -ForegroundColor DarkCyan

$bitlockerRows = New-Object System.Collections.Generic.List[string]

try {
    $volumes = Get-BitLockerVolume -ErrorAction Stop
    if ($volumes.Count -eq 0) {
        $bitlockerRows.Add("<tr><td colspan='3'>No volumes found.</td></tr>")
    }
    foreach ($v in $volumes) {
        $statusClass = if ($v.ProtectionStatus -eq 'On') { "hasperm" } else { "noperm" }
        $bitlockerRows.Add("<tr><td class='labelCol'>$(Enc $v.MountPoint)</td><td>$(Enc([string]$v.VolumeStatus))</td><td class='$statusClass'>$(Enc([string]$v.ProtectionStatus))</td></tr>")
    }
}
catch {
    $bitlockerRows.Add("<tr><td colspan='3' class='errCell'>Unable to retrieve BitLocker status (module not available, feature not present, or insufficient permission): $(Enc $_.Exception.Message)</td></tr>")
}

# =================================================================
# 5. Antivirus Status
# =================================================================
Write-Host " - 收集防毒软件状态..." -ForegroundColor DarkCyan

$avRows = New-Object System.Collections.Generic.List[string]
$avFound = $false

try {
    $avProducts = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction Stop
    foreach ($av in $avProducts) {
        $avFound = $true
        $displayName = $av.displayName

        # productState 是 hex bitmask，位于中间的一个字节代表是否启用(实时保护)
        $stateHex = "{0:X6}" -f $av.productState
        $enabledFlag = $stateHex.Substring(2,2)
        $isEnabled = ($enabledFlag -eq "10" -or $enabledFlag -eq "11")
        $enabledText = if ($isEnabled) { "Enabled" } else { "Disabled / Unknown" }
        $enabledClass = if ($isEnabled) { "hasperm" } else { "noperm" }

        $lastUpdate = "-"
        try {
            if ($av.timestamp) {
                $lastUpdate = [datetime]::ParseExact($av.timestamp.Substring(0,14), "yyyyMMddHHmmss", $null).ToString("yyyy-MM-dd HH:mm:ss")
            }
        }
        catch { $lastUpdate = $av.timestamp }

        $avRows.Add("<tr><td class='labelCol'>$(Enc $displayName)</td><td class='$enabledClass'>$(Enc $enabledText)</td><td>$(Enc $lastUpdate)</td></tr>")
    }
}
catch {
    # SecurityCenter2 在某些 Windows 版本 (例如 Server) 上不存在，属正常现象
}

# 补充 Windows Defender 自身的详细状态 (若存在)
try {
    $mp = Get-MpComputerStatus -ErrorAction Stop
    $avFound = $true
    $rtpText  = if ($mp.RealTimeProtectionEnabled) { "Enabled" } else { "Disabled" }
    $rtpClass = if ($mp.RealTimeProtectionEnabled) { "hasperm" } else { "noperm" }
    $avRows.Add("<tr><td class='labelCol'>Windows Defender - Real-time Protection</td><td class='$rtpClass'>$(Enc $rtpText)</td><td>-</td></tr>")

    $sigDate = if ($mp.AntivirusSignatureLastUpdated) { $mp.AntivirusSignatureLastUpdated.ToString("yyyy-MM-dd HH:mm:ss") } else { "-" }
    $avRows.Add("<tr><td class='labelCol'>Windows Defender - Signature Version</td><td>$(Enc $mp.AntivirusSignatureVersion)</td><td>$(Enc $sigDate)</td></tr>")
}
catch {
    # Windows Defender 模块不存在或已被第三方防毒软件取代，属正常现象
}

if (-not $avFound) {
    $avRows.Add("<tr><td colspan='3' class='errCell'>Unable to detect any antivirus product on this computer.</td></tr>")
}

# =================================================================
# 6. Installed Software
# =================================================================
Write-Host " - 收集已安装软件清单..." -ForegroundColor DarkCyan

$softwareRows = New-Object System.Collections.Generic.List[string]

try {
    $uninstallPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $software = Get-ItemProperty -Path $uninstallPaths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne "" } |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate -Unique |
        Sort-Object DisplayName

    if (-not $software -or $software.Count -eq 0) {
        $softwareRows.Add("<tr><td colspan='4'>No installed software found.</td></tr>")
    }
    else {
        foreach ($s in $software) {
            $installDateDisplay = "-"
            if ($s.InstallDate -and $s.InstallDate -match '^\d{8}$') {
                try {
                    $installDateDisplay = [datetime]::ParseExact($s.InstallDate, "yyyyMMdd", $null).ToString("yyyy-MM-dd")
                }
                catch { $installDateDisplay = $s.InstallDate }
            }
            elseif ($s.InstallDate) {
                $installDateDisplay = $s.InstallDate
            }
            $softwareRows.Add("<tr><td class='labelCol'>$(Enc $s.DisplayName)</td><td>$(Enc $s.DisplayVersion)</td><td>$(Enc $s.Publisher)</td><td>$(Enc $installDateDisplay)</td></tr>")
        }
    }
}
catch {
    $softwareRows.Add("<tr><td colspan='4' class='errCell'>Unable to retrieve installed software list: $(Enc $_.Exception.Message)</td></tr>")
}

# =================================================================
# 7. Shared Folders (共享文件夹 与 共享权限明细)
# =================================================================
Write-Host " - 收集本机共享文件夹信息..." -ForegroundColor DarkCyan

$shareRows = New-Object System.Collections.Generic.List[string]

$shares = $null
$smbAvailable = $false

try {
    # Get-SmbShare 在一般用户权限下多半可以列出共享，但部分系统需要系统管理员权限
    $shares = Get-SmbShare -ErrorAction Stop | Where-Object { $_.ShareType -eq 'FileSystemDirectory' }
    $smbAvailable = $true
}
catch {
    try {
        # 备用方案: Win32_Share 通常一般用户也可以查询到 (但没有权限明细)
        $shares = Get-CimInstance -ClassName Win32_Share -ErrorAction Stop | Where-Object { $_.Type -eq 0 }
    }
    catch {
        $shares = $null
    }
}

if (-not $shares -or ($shares | Measure-Object).Count -eq 0) {
    $shareRows.Add("<tr><td colspan='4'>No shared folders found on this computer.</td></tr>")
}
else {
    foreach ($sh in $shares) {
        $shareName = $sh.Name
        $sharePath = $sh.Path
        $shareDesc = $sh.Description

        $permEntries = New-Object System.Collections.Generic.List[string]

        if ($smbAvailable) {
            try {
                $accessList = Get-SmbShareAccess -Name $shareName -ErrorAction Stop
                foreach ($acc in $accessList) {
                    $permEntries.Add("$(Enc $acc.AccountName): $(Enc $acc.AccessControlType.ToString()) - $(Enc $acc.AccessRight.ToString())")
                }
            }
            catch {
                $permEntries.Add("<span class='errCell'>Unable to retrieve permissions for this share (may require administrator privileges): $(Enc $_.Exception.Message)</span>")
            }
        }
        else {
            $permEntries.Add("<span class='errCell'>Share permission details require administrator privileges and are not available under a standard user account.</span>")
        }

        if ($permEntries.Count -eq 0) {
            $permEntries.Add("-")
        }

        $permDisplay = ($permEntries -join "<br/>")

        $shareRows.Add("<tr><td class='labelCol'>$(Enc $shareName)</td><td>$(Enc $sharePath)</td><td>$(Enc $shareDesc)</td><td>$permDisplay</td></tr>")
    }
}

# =================================================================
# 8. Screen Saver Settings (屏幕保护程序: 是否启用 / 等待分钟数 / 恢复时是否需要登录)
# =================================================================
Write-Host " - 收集屏幕保护程序设置..." -ForegroundColor DarkCyan

$screenSaverRows = New-Object System.Collections.Generic.List[string]

try {
    $desktopKey = 'HKCU:\Control Panel\Desktop'

    $ssActiveRaw  = (Get-ItemProperty -Path $desktopKey -Name 'ScreenSaveActive'   -ErrorAction SilentlyContinue).ScreenSaveActive
    $ssTimeoutRaw = (Get-ItemProperty -Path $desktopKey -Name 'ScreenSaveTimeOut'  -ErrorAction SilentlyContinue).ScreenSaveTimeOut
    $ssSecureRaw  = (Get-ItemProperty -Path $desktopKey -Name 'ScreenSaverIsSecure' -ErrorAction SilentlyContinue).ScreenSaverIsSecure
    $ssExeRaw     = (Get-ItemProperty -Path $desktopKey -Name 'SCRNSAVE.EXE'       -ErrorAction SilentlyContinue).'SCRNSAVE.EXE'

    $isActive = ($ssActiveRaw -eq '1' -or $ssActiveRaw -eq 1)
    $activeText  = if ($isActive) { "Enabled" } else { "Disabled" }
    $activeClass = if ($isActive) { "hasperm" } else { "noperm" }
    $screenSaverRows.Add("<tr><td class='labelCol'>Screen Saver</td><td class='$activeClass'>$(Enc $activeText)</td></tr>")

    if ($ssTimeoutRaw) {
        $minutes = [math]::Round([int]$ssTimeoutRaw / 60, 1)
        $screenSaverRows.Add("<tr><td class='labelCol'>Wait Time</td><td>$minutes minute(s)</td></tr>")
    }
    else {
        $screenSaverRows.Add("<tr><td class='labelCol'>Wait Time</td><td>-</td></tr>")
    }

    $isSecure = ($ssSecureRaw -eq '1' -or $ssSecureRaw -eq 1)
    $secureText  = if ($isSecure) { "Yes - Login required on resume" } else { "No - Login not required on resume" }
    $secureClass = if ($isSecure) { "hasperm" } else { "noperm" }
    $screenSaverRows.Add("<tr><td class='labelCol'>On Resume, Display Logon Screen</td><td class='$secureClass'>$(Enc $secureText)</td></tr>")

    if ($ssExeRaw) {
        $screenSaverRows.Add("<tr><td class='labelCol'>Screen Saver Program</td><td>$(Enc $ssExeRaw)</td></tr>")
    }
}
catch {
    $screenSaverRows.Add("<tr><td colspan='2' class='errCell'>Unable to retrieve screen saver settings: $(Enc $_.Exception.Message)</td></tr>")
}

# 9. Windows quality / feature update history
$windowsUpdateRows = New-Object System.Collections.Generic.List[string]
$session = $null
$searcher = $null
$history = $null
$entry = $null
try {
    $session = New-Object -ComObject Microsoft.Update.Session -ErrorAction Stop
    $searcher = $session.CreateUpdateSearcher()
    $total = $searcher.GetTotalHistoryCount()
    # 分批向前读取，筛选后才取最近 10 条，Defender 不占名额。
    for ($offset = 0; $offset -lt $total -and $windowsUpdateRows.Count -lt 10; $offset += 100) {
        $history = $searcher.QueryHistory($offset, [Math]::Min(100, $total - $offset))
        for ($i = 0; $i -lt $history.Count -and $windowsUpdateRows.Count -lt 10; $i++) {
            $entry = $history.Item($i)
            try {
                $title = [string]$entry.Title
                $exclude = $title -match '(?i)Defender|Security Intelligence|安全智能|安全情報|安全情报|安全性情報|Antimalware|Malicious Software Removal|恶意软件删除|惡意軟體移除|KB890830|KB2267602|KB915597'
                $feature = $title -match '(?i)Feature update.*Windows|Windows.*Feature update|功能更新.*Windows|Windows.*功能更新|Enablement Package|启用包|啟用套件|Upgrade to Windows|升级到 Windows|升級至 Windows'
                $quality = (($title -match '(?i)Windows|\.NET') -and ($title -match '(?i)Cumulative Update|Security Update|Update for|Servicing Stack|累积更新|累積更新|累计更新|安全更新|安全性更新|预览更新|預覽更新|更新预览|更新預覽|服务堆栈|服務堆疊')) -or ($title -match '(?i)^\s*\d{4}-\d{2}\s+(Preview Update|Security Update|Update|预览更新|預覽更新|安全更新|安全性更新|更新)')
                if ($exclude -or (-not $feature -and -not $quality)) { continue }
                $kind = if ($feature) { 'Feature' } else { 'Quality' }
                $kb = ([regex]::Matches($title, '(?i)\bKB\d+\b') | ForEach-Object { $_.Value.ToUpperInvariant() } | Select-Object -Unique) -join ', '
                if (-not $kb) { $kb = '-' }
                $date = ([datetime]$entry.Date).ToString('yyyy-MM-dd HH:mm:ss')
                $operation = switch ([int]$entry.Operation) { 1 {'Installation'} 2 {'Uninstallation'} default {'Unknown'} }
                $result = switch ([int]$entry.ResultCode) { 0 {'Not started'} 1 {'In progress'} 2 {'Succeeded'} 3 {'Succeeded with errors'} 4 {'Failed'} 5 {'Aborted'} default {'Unknown'} }
                $windowsUpdateRows.Add("<tr><td>$(Enc $date)</td><td>$(Enc $kind)</td><td style='overflow-wrap:anywhere;'>$(Enc $title)</td><td>$(Enc $kb)</td><td>$(Enc $operation)</td><td>$(Enc $result)</td></tr>")
            } finally {
                if ($null -ne $entry) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($entry); $entry = $null }
            }
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($history)
        $history = $null
    }
    if ($windowsUpdateRows.Count -eq 0) {
        $windowsUpdateRows.Add("<tr><td colspan='6'>No matching Windows quality / feature update history records found.</td></tr>")
    }
} catch {
    $windowsUpdateRows.Add("<tr><td colspan='6' class='errCell'>Unable to retrieve update history: $(Enc $_.Exception.Message)</td></tr>")
} finally {
    foreach ($obj in @($entry, $history, $searcher, $session)) {
        if ($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
        }
    }
}
$windowsUpdateHtml = $windowsUpdateRows -join "`n"

# =================================================================
# 组合 HTML 报告
# =================================================================
Write-Host "正在生成 HTML 报告..." -ForegroundColor Cyan

$aboutHtml     = $aboutRows -join "`n"
$userHtml      = $userRows -join "`n"
$adminHtml     = $adminRows -join "`n"
$bitlockerHtml = $bitlockerRows -join "`n"
$avHtml        = $avRows -join "`n"
$softwareHtml  = $softwareRows -join "`n"
$shareHtml     = $shareRows -join "`n"
$screenSaverHtml = $screenSaverRows -join "`n"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>$reportTitle</title>
<style>
    @page { size: portrait; margin: 15mm; }
    body { font-family: "Segoe UI", "Microsoft YaHei", Arial, sans-serif; margin: 20px; background:#f7f7f9; color:#222; }
    h1 { font-size: 22px; margin-bottom: 4px; }
    h2 { font-size: 16px; margin: 0 0 10px 0; color: #1F4E79; border-bottom: 2px solid #1F4E79; padding-bottom: 4px; }
    .meta { color: #333; font-size: 12px; font-weight: bold; margin-bottom: 10px; }
    .signature-line { display: flex; flex-wrap: wrap; gap: 22px; margin: 0 0 20px 0; font-size: 13px; color: #333; }
    .sig-item { display: flex; align-items: flex-end; white-space: nowrap; }
    .sig-blank { display: inline-block; width: 120px; border-bottom: 1px solid #333; margin-left: 6px; height: 16px; }
    .panel { background:#fff; border: 1px solid #ddd; border-radius: 6px; padding: 12px 16px; margin-bottom: 20px; }
    .panel h2.section-title { font-size: 14px; margin: 0 0 8px 0; color: #333; border: none; padding: 0; }
    .panel-buttons { margin-bottom: 10px; }
    .panel-buttons button { font-size: 12px; padding: 4px 10px; margin-right: 8px; cursor: pointer; }
    .toggle-list { display: flex; flex-wrap: wrap; gap: 6px 18px; }
    .toggle-item { font-size: 13px; white-space: nowrap; }
    .toggle-item input { margin-right: 4px; }
    .section { background:#fff; border: 1px solid #ddd; border-radius: 6px; padding: 16px 20px; margin-bottom: 20px; }
    table { border-collapse: collapse; width: 100%; }
    thead { display: table-header-group; }
    tr { page-break-inside: avoid; }
    th, td { border: 1px solid #ddd; padding: 6px 10px; font-size: 12px; text-align: left; vertical-align: top; }
    th { background: #2f5597; color: #fff; }
    td.labelCol { font-weight: bold; background: #f0f4fa; white-space: nowrap; }
    td.hasperm { color: #1a7f37; font-weight: bold; }
    td.noperm { color: #b00020; font-weight: bold; }
    td.errCell { color: #b00020; font-style: italic; }
    .page-section { page-break-before: always; }
    @media print {
        body { background: #fff; margin: 0; }
        .section { border: none; box-shadow: none; }
        .panel { display: none; }
    }
</style>
</head>
<body>

<h1>$reportTitle</h1>
<div class="meta">Computer: $(Enc $env:COMPUTERNAME) &nbsp;|&nbsp; Generated: $generatedAt &nbsp;|&nbsp; Collected by: $(Enc "$env:USERDOMAIN\$env:USERNAME") &nbsp;|&nbsp; IP Address: $(Enc $ipListDisplay)</div>

<div class="signature-line">
    <span class="sig-item">DGM:<span class="sig-blank"></span></span>
    <span class="sig-item">CISO:<span class="sig-blank"></span></span>
    <span class="sig-item">ISO:<span class="sig-blank"></span></span>
    <span class="sig-item">IT MANAGER:<span class="sig-blank"></span></span>
    <span class="sig-item">IT:<span class="sig-blank"></span></span>
</div>

<div class="panel">
    <h2 class="section-title">Show / Hide Sections (checked sections will be displayed below; default is all shown)</h2>
    <div class="panel-buttons">
        <button type="button" onclick="setAllSections(true)">Show All</button>
        <button type="button" onclick="setAllSections(false)">Hide All</button>
    </div>
    <div class="toggle-list" id="sectionToggleList">
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-1" checked onchange="applySectionFilters()"> 1. About This PC</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-2" checked onchange="applySectionFilters()"> 2. Local User Accounts</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-3" checked onchange="applySectionFilters()"> 3. Administrators Group Members</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-4" checked onchange="applySectionFilters()"> 4. BitLocker Status</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-5" checked onchange="applySectionFilters()"> 5. Antivirus Status</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-6" checked onchange="applySectionFilters()"> 6. Installed Software</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-7" checked onchange="applySectionFilters()"> 7. Shared Folders</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-8" checked onchange="applySectionFilters()"> 8. Screen Saver Settings</label>
        <label class="toggle-item"><input type="checkbox" class="sectionChk" data-target="section-9" checked onchange="applySectionFilters()"> 9. Windows Quality / Feature Updates (Latest 10)</label>
    </div>
</div>

<div class="section" id="section-1">
    <h2>1. About This PC</h2>
    <table>
        <tbody>
        $aboutHtml
        </tbody>
    </table>
</div>

<div class="section page-section" id="section-2">
    <h2>2. Local User Accounts</h2>
    <table>
        <thead><tr><th style="width:25%">User Name</th><th style="width:45%">Member Of (Local Groups)</th><th style="width:15%">Status</th></tr></thead>
        <tbody>
        $userHtml
        </tbody>
    </table>
</div>

<div class="section" id="section-3">
    <h2>3. Administrators Group Members</h2>
    <table>
        <thead><tr><th style="width:40%">Name</th><th style="width:20%">Type</th><th style="width:20%">Source</th></tr></thead>
        <tbody>
        $adminHtml
        </tbody>
    </table>
</div>

<div class="section page-section" id="section-4">
    <h2>4. BitLocker Status</h2>
    <table>
        <thead><tr><th style="width:20%">Drive</th><th style="width:40%">Volume Status</th><th style="width:40%">Protection Status</th></tr></thead>
        <tbody>
        $bitlockerHtml
        </tbody>
    </table>
</div>

<div class="section" id="section-5">
    <h2>5. Antivirus Status</h2>
    <table>
        <thead><tr><th style="width:40%">Product</th><th style="width:25%">Status</th><th style="width:25%">Last Update</th></tr></thead>
        <tbody>
        $avHtml
        </tbody>
    </table>
</div>

<div class="section page-section" id="section-6">
    <h2>6. Installed Software</h2>
    <table>
        <thead><tr><th style="width:35%">Name</th><th style="width:15%">Version</th><th style="width:30%">Publisher</th><th style="width:15%">Install Date</th></tr></thead>
        <tbody>
        $softwareHtml
        </tbody>
    </table>
</div>

<div class="section page-section" id="section-7">
    <h2>7. Shared Folders</h2>
    <table>
        <thead><tr><th style="width:15%">Share Name</th><th style="width:25%">Local Path</th><th style="width:20%">Description</th><th style="width:40%">Permissions (Account: Type - Right)</th></tr></thead>
        <tbody>
        $shareHtml
        </tbody>
    </table>
</div>

<div class="section page-section" id="section-8">
    <h2>8. Screen Saver Settings</h2>
    <table>
        <tbody>
        $screenSaverHtml
        </tbody>
    </table>
</div>

<div class="section page-section" id="section-9">
<h2>9. Windows Quality / Feature Updates (Latest 10)</h2>
<p style="font-size:12px;">Latest 10 matching events, newest first. Includes Windows / .NET quality updates and Windows feature updates. Defender intelligence updates are excluded. Classification uses titles and may differ from Settings. Includes failed attempts and uninstallations; this is not a current installed-patch inventory.</p>
<table><thead><tr><th>Date / Time</th><th>Type</th><th>Update Title</th><th>KB</th><th>Operation</th><th>Result</th></tr></thead>
<tbody>$windowsUpdateHtml</tbody></table>
</div>

<script>
function applySectionFilters() {
    document.querySelectorAll('.sectionChk').forEach(function(chk) {
        var target = document.getElementById(chk.getAttribute('data-target'));
        if (target) {
            target.style.display = chk.checked ? '' : 'none';
        }
    });
}

function setAllSections(checked) {
    document.querySelectorAll('.sectionChk').forEach(function(chk) {
        chk.checked = checked;
    });
    applySectionFilters();
}

document.addEventListener('DOMContentLoaded', applySectionFilters);
</script>

</body>
</html>
"@

$html | Out-File -FilePath $htmlPath -Encoding UTF8

Write-Host "HTML 报告已生成: $htmlPath" -ForegroundColor Green
Write-Host "完成。" -ForegroundColor Cyan