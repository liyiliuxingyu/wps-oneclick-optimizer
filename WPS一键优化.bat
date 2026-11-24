@echo off
chcp 65001 >nul 2>&1
setlocal EnableExtensions
title WPS Office Optimize (single file)
set "SELF=%~f0"
set "BASE=%~dp0"
set "PS1=%TEMP%\WPS-Optimize-single.ps1"
set "WPS_OPT_BASE=%BASE%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:SELF,[Text.Encoding]::UTF8); $m='#::WPS_PS_'+'BEGIN_9F3A::'; $i=$t.LastIndexOf($m); if($i -lt 0){exit 9}; $code=$t.Substring($i+$m.Length); [IO.File]::WriteAllText($env:PS1,$code,[Text.UTF8Encoding]::new($true)); exit 0"
if errorlevel 1 (
    echo [ERROR] Failed to unpack the embedded PowerShell script.
    pause
    exit /b 1
)
if "%~1"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
)
set "RC=%ERRORLEVEL%"
del "%PS1%" >nul 2>&1
if not "%RC%"=="0" pause
exit /b %RC%
#::WPS_PS_BEGIN_9F3A::
<#
.SYNOPSIS
    WPS Office 一键优化（单文件版：优化 / 检测 / 恢复 三合一）
    关闭推荐与稻壳AI模板弹窗、隐藏系统右键菜单、免强制登录、禁用自启动、关闭 WPS 网盘与文件传输助手、清理资源管理器入口

.DESCRIPTION
    本脚本针对 Windows 上的 WPS Office（个人版 12.x 实测通过），一键完成：
      1. 功能定制：勾选关闭官方“WPS Office 配置工具 → 功能定制”中的 13 个推荐 / 系统右键菜单项
         （前 3 项推荐 + 10 项系统右键菜单隐藏项）。这些设置在 WPS 中大部分以“加密开关（KIC）”保存，
         无法用注册表伪造，因此脚本调用官方配置工具 ksomisc.exe，用 Windows UI 自动化替你勾选并点“确定”，
         等价于人工操作；若 UI 自动化不可用，则自动降级为写入已知的明文开关（并提示剩余项需手动处理）。
      2. 免登录/离线：写入 WPS 官方的强制登录开关，跳过“首次安装设备/已安装设备强制登录”，
         减少登录弹窗；同时脚本会关闭 WPS 云服务，避免账号登录服务常驻。
      3. 关闭 WPS 自启动：禁用 WPS 更新计划任务（WpsUpdateTask / WpsUpdateLogonTask / WpsWakeWnsLogonTask）、
         停用并禁用 WPS Office Cloud Service（wpscloudsvr）、清理注册表 Run/RunOnce 与启动文件夹中 WPS 相关项。
      4. 关闭 WPS 网盘 / 文件传输助手：关闭 WPS 云盘常驻开关与传输助手总开关，配合云服务禁用实现“不开机自启、不常驻”。
      5. 关闭启动弹窗：写入 WPS 官方开关 docerTabShow=0，关闭每次启动 WPS 时自动打开的「稻壳AI模板」标签页；
         同时写入 rightPanelWidgetVisibleGlobal=false，关闭 WPS 启动首页右侧的「精选推荐 / 功能推荐」面板。

    所有改动前会自动备份（注册表、计划任务 XML、服务启动类型、启动项）；可用 -Mode Revert 一键恢复。

.PARAMETER Mode
    Menu   显示中文菜单（默认；双击单文件脚本时使用）。
    Apply  执行优化（自动备份、可重复执行/幂等）。
    Scan   只检测、不改动，输出当前各项状态。
    Revert 从脚本目录下 backup 中“最早一次备份”恢复原始状态。

.PARAMETER SkipUiAutomation
    跳过“功能定制”自动勾选（只做注册表与自启动部分；适合无人值守/无桌面环境）。

.PARAMETER SkipService
    跳过 WPS 云服务（wpscloudsvr）的停用与禁用。

.PARAMETER SkipScheduledTasks
    跳过 WPS 计划任务的禁用。

.PARAMETER SkipCloud
    跳过 WPS 网盘 / 文件传输助手相关注册表开关的关闭。

.PARAMETER SkipLogin
    跳过“免强制登录”注册表开关的写入。

.PARAMETER Elevated
    内部参数：表示当前已是提权后的实例，避免重复提权。

.PARAMETER Pause
    结束时等待回车（单文件 .bat 启动时会自动带上）。

.PARAMETER WpsPath
    可选：手动指定 WPS 安装目录、office6 目录或 ksomisc.exe 的完整路径。
    一般不需要填——脚本会自动从 注册表安装信息 / 卸载项 / 文件关联 / 运行进程 /
    系统服务 / 开始菜单快捷方式 / PATH / 固定磁盘浅层扫描 中查找 WPS。
    自动找不到时（例如 WPS 装在非常规目录、注册表被清理），再手动指定，例如：
    WPS一键优化.bat -Mode Apply -WpsPath "D:\Apps\WPS Office\12.1.0.29690"

.EXAMPLE
    WPS一键优化.bat
    双击运行，出现菜单后按提示选择 1/2/3/0。

.EXAMPLE
    WPS一键优化.bat -Mode Apply
    不进菜单，直接执行一键优化（适合命令行、批处理）。

.EXAMPLE
    WPS一键优化.bat -Mode Scan
    只读检测当前状态，不做任何修改。

.NOTES
    · 建议先保存并关闭 WPS 文档再运行；「功能定制」需要交互式桌面会话（Windows UI 自动化）。
    · 安装路径自动发现：依次从 注册表 InstallRoot、卸载项、文件关联、运行中的 WPS 进程、
      系统服务、开始菜单快捷方式、PATH(where.exe)、固定磁盘浅层扫描 中查找 ksomisc.exe；
      都找不到时可用 -WpsPath 手动指定安装目录/office6/ksomisc.exe。
    · 系统右键菜单项与资源管理器集成的改动，重启资源管理器或重启电脑后完全生效。
    · WPS 个人版的云文档、WPS AI、在线模板、PDF 会员等功能本身需要登录/会员才可以合法使用，
      本脚本只负责“免强制登录 + 关闭云/更新/推荐常驻”，不包含任何破解、破解补丁或绕过付费的内容。
    · WPS 升级或“自愈”后可能恢复部分设置，重新运行本脚本即可。
#>
[CmdletBinding()]
param(
    [ValidateSet('Menu', 'Apply', 'Scan', 'Revert')]
    [string]$Mode = 'Menu',
    [switch]$SkipUiAutomation,
    [switch]$SkipService,
    [switch]$SkipScheduledTasks,
    [switch]$SkipCloud,
    [switch]$SkipLogin,
    [switch]$SkipExplorerEntries,
    [switch]$RestartExplorer,
    [string]$WpsPath,
    [switch]$Elevated,
    [switch]$Pause
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$script:ScriptVersion = '2.2'
$script:ScriptDir = if ($env:WPS_OPT_BASE) { $env:WPS_OPT_BASE.TrimEnd([char]92) } elseif ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$script:LogDir = Join-Path $script:ScriptDir 'logs'
$script:BackupRoot = Join-Path $script:ScriptDir 'backup'
$script:LogFile = $null
$script:Summary = New-Object System.Collections.ArrayList
$script:UiReady = $false
$script:WpsInfo = $null
$script:BackupDir = $null
$script:SkipFinalPause = $false
$script:LogBuffer = New-Object System.Collections.ArrayList
$script:CreatedValues = New-Object System.Collections.ArrayList
$script:LogFlushThreshold = 60

# ----------------------------------------------------------------------------
# 基础工具函数
# ----------------------------------------------------------------------------
function Write-CliLine {
    param([string]$Text = '', [string]$Color = '')
    try {
        $oldColor = $null
        if ($Color) {
            try { $oldColor = [Console]::ForegroundColor } catch { }
            switch ($Color) {
                'Green' { try { [Console]::ForegroundColor = [ConsoleColor]::Green } catch { } }
                'Yellow' { try { [Console]::ForegroundColor = [ConsoleColor]::Yellow } catch { } }
                'Red' { try { [Console]::ForegroundColor = [ConsoleColor]::Red } catch { } }
                'Cyan' { try { [Console]::ForegroundColor = [ConsoleColor]::Cyan } catch { } }
                'DarkGray' { try { [Console]::ForegroundColor = [ConsoleColor]::DarkGray } catch { } }
                default { }
            }
        }
        [Console]::WriteLine($Text)
        if ($null -ne $oldColor) { try { [Console]::ForegroundColor = $oldColor } catch { } }
    } catch {
        try { [Console]::WriteLine($Text) } catch { }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERR', 'STEP')][string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = '[' + $ts + '][' + $Level + '] ' + $Message
    switch ($Level) {
        'OK'   { Write-CliLine -Text $line -Color 'Green' }
        'WARN' { Write-CliLine -Text $line -Color 'Yellow' }
        'ERR'  { Write-CliLine -Text $line -Color 'Red' }
        'STEP' { Write-CliLine ''; Write-CliLine -Text ('===== ' + $Message + ' =====') -Color 'Cyan' }
        default { Write-CliLine -Text $line }
    }
    if ($script:LogFile) {
        $null = $script:LogBuffer.Add($line)
        if ($script:LogBuffer.Count -ge $script:LogFlushThreshold) { Save-LogBuffer }
    }
}

function Save-LogBuffer {
    if (-not $script:LogFile) { return }
    if ($script:LogBuffer.Count -eq 0) { return }
    try {
        Add-Content -Path $script:LogFile -Value $script:LogBuffer -Encoding UTF8
        $script:LogBuffer.Clear()
    } catch { }
}

function Add-Result {
    param(
        [string]$Category,
        [string]$Item,
        [string]$Status,
        [string]$Detail = ''
    )
    $null = $script:Summary.Add([pscustomobject]@{
        Category = $Category
        Item     = $Item
        Status   = $Status
        Detail   = $Detail
    })
}

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Invoke-SelfElevate {
    $argParts = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'), '-Mode', $Mode, '-Pause', '-Elevated')
    if ($SkipUiAutomation) { $argParts += '-SkipUiAutomation' }
    if ($SkipService) { $argParts += '-SkipService' }
    if ($SkipScheduledTasks) { $argParts += '-SkipScheduledTasks' }
    if ($SkipCloud) { $argParts += '-SkipCloud' }
    if ($SkipLogin) { $argParts += '-SkipLogin' }
    $argString = $argParts -join ' '
    Write-Log '当前不是管理员权限，正在请求提升（将弹出 UAC 确认框）...' 'WARN'
    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argString -Verb RunAs -Wait -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log ('提权失败或已取消：' + $_.Exception.Message) 'WARN'
        Write-Log '将继续以当前（普通）权限运行；计划任务/服务/部分系统项可能无法修改。' 'WARN'
        return $false
    }
}

function Initialize-Environment {
    if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Force -Path $script:LogDir | Out-Null }
    if (-not $script:LogFile) {
        $script:LogFile = Join-Path $script:LogDir ('WPS-Optimize_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')
    }
}

# ----------------------------------------------------------------------------
# WPS 安装路径自动发现（找不到时会依次尝试多种来源，最后扫描固定磁盘）
# ----------------------------------------------------------------------------
function Find-FirstFile {
    param([string]$Pattern)
    try {
        $items = @(Get-Item -Path $Pattern -ErrorAction Stop)
        foreach ($it in $items) {
            if ($it -and (-not $it.PSIsContainer)) { return [string]$it.FullName }
        }
    } catch { }
    return $null
}

function New-WpsLayout {
    param([string]$Ksomisc)
    $dir = Split-Path -Parent $Ksomisc
    $root = $dir
    if ((Split-Path -Leaf $dir) -ieq 'office6') { $root = Split-Path -Parent $dir }
    return [pscustomobject]@{ Ksomisc = $Ksomisc; InstallRoot = $root }
}

function Get-WpsLayoutFromRoot {
    param([string]$Root)
    if ([string]::IsNullOrWhiteSpace($Root)) { return $null }
    $r = $Root.Trim().Trim('"').TrimEnd([char]92)
    if ([string]::IsNullOrWhiteSpace($r)) { return $null }
    if (-not [System.IO.Directory]::Exists($r)) { return $null }
    # 直接命中：根目录下 ksomisc.exe / office6\ksomisc.exe
    foreach ($f in @(
            (Join-Path $r 'ksomisc.exe'),
            (Join-Path $r 'office6\ksomisc.exe')
        )) {
        if ([System.IO.File]::Exists($f)) { return (New-WpsLayout -Ksomisc $f) }
    }
    $children = @()
    try { $children = [System.IO.Directory]::GetDirectories($r) } catch { return $null }
    # 一层子目录：...\<版本目录>\ksomisc.exe 或 ...\<任意>\office6\ksomisc.exe
    foreach ($child in $children) {
        $f1 = Join-Path $child 'ksomisc.exe'
        if ([System.IO.File]::Exists($f1)) { return (New-WpsLayout -Ksomisc $f1) }
        $f2 = Join-Path $child 'office6\ksomisc.exe'
        if ([System.IO.File]::Exists($f2)) { return (New-WpsLayout -Ksomisc $f2) }
    }
    # 两层子目录：...\<父目录>\<版本目录>\office6\ksomisc.exe
    foreach ($child in $children) {
        $grand = @()
        try { $grand = [System.IO.Directory]::GetDirectories($child) } catch { $grand = @() }
        foreach ($g in $grand) {
            $f3 = Join-Path $g 'office6\ksomisc.exe'
            if ([System.IO.File]::Exists($f3)) { return (New-WpsLayout -Ksomisc $f3) }
        }
    }
    return $null
}

function Resolve-WpsCandidate {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim()
    # 命令串（"xxx.exe" args / xxx.exe args）先去掉引号与参数
    if ($p -match '^"([^"]+)"') { $p = $Matches[1] }
    elseif ($p -match '^([^\s"]+\.exe)') { $p = $Matches[1] }
    # 去掉图标索引 ,0
    $p = ($p -replace ',(%?\d+)$', '').Trim().Trim('"').Trim()
    if ([string]::IsNullOrWhiteSpace($p)) { return $null }
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    $isContainer = $false
    try { $isContainer = (Test-Path -LiteralPath $p -PathType Container) } catch { }
    if ($isContainer) {
        $layout = Get-WpsLayoutFromRoot -Root $p
        if ($layout) { return $layout }
        $leaf = Split-Path -Leaf $p.TrimEnd([char]92)
        if ($leaf -ieq 'office6') { return (Get-WpsLayoutFromRoot -Root (Split-Path -Parent $p.TrimEnd([char]92))) }
        return $null
    }
    # 传入的是文件：先看所在目录，再逐级向上找安装根
    $dir = Split-Path -Parent $p
    $tryDirs = @($dir)
    $up1 = Split-Path -Parent $dir
    if ($up1) { $tryDirs += $up1 }
    $up2 = Split-Path -Parent $up1
    if ($up2) { $tryDirs += $up2 }
    foreach ($d in $tryDirs) {
        $layout = Get-WpsLayoutFromRoot -Root $d
        if ($layout) { return $layout }
    }
    return $null
}

function Get-WpsQuickCandidates {
    $list = New-Object System.Collections.ArrayList
    # 1) Kingsoft Office 注册表 InstallRoot（HKCU / HKLM / WOW6432）
    foreach ($vk in @(
            'HKCU:\Software\Kingsoft\Office\6.0\common',
            'HKLM:\SOFTWARE\Kingsoft\Office\6.0\Common',
            'HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office\6.0\Common'
        )) {
        if (-not (Test-Path $vk)) { continue }
        $kv = Get-ItemProperty -Path $vk -ErrorAction SilentlyContinue
        if ($kv -and $kv.InstallRoot) {
            $null = $list.Add([pscustomobject]@{ Method = '注册表 InstallRoot'; Path = [string]$kv.InstallRoot })
        }
    }
    # 2) 卸载项 InstallLocation / UninstallString / DisplayIcon（.NET 快速枚举）
    $uninstallRoots = @(
        @{ Hive = [Microsoft.Win32.Registry]::CurrentUser; Sub = 'Software\Microsoft\Windows\CurrentVersion\Uninstall' },
        @{ Hive = [Microsoft.Win32.Registry]::LocalMachine; Sub = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' },
        @{ Hive = [Microsoft.Win32.Registry]::LocalMachine; Sub = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' }
    )
    foreach ($root in $uninstallRoots) {
        $key = $null
        try { $key = $root.Hive.OpenSubKey($root.Sub) } catch { }
        if (-not $key) { continue }
        try {
            foreach ($name in $key.GetSubKeyNames()) {
                $sk = $null
                try { $sk = $key.OpenSubKey($name) } catch { }
                if (-not $sk) { continue }
                try {
                    $displayName = [string]$sk.GetValue('DisplayName')
                    if ($displayName -match 'WPS Office|Kingsoft Office|WPS') {
                        foreach ($prop in @('InstallLocation', 'UninstallString', 'DisplayIcon')) {
                            $val = [string]$sk.GetValue($prop)
                            if ($val) { $null = $list.Add([pscustomobject]@{ Method = ('卸载项 ' + $prop); Path = $val }) }
                        }
                    }
                } finally { $sk.Close() }
            }
        } finally { $key.Close() }
    }
    # 3) 文件关联 / App Paths（用 .NET 读 HKCR/HKCU/HKLM Classes，避免 HKCR: 驱动器差异）
    $progIds = @('KWPS.Document.9', 'WPS.Doc.6', 'KWPS.PDF.9', 'KET.Workbook.9', 'WPP.PPT.6')
    $classesHives = @(
        @{ Hive = [Microsoft.Win32.Registry]::ClassesRoot; Sub = '' },
        @{ Hive = [Microsoft.Win32.Registry]::CurrentUser; Sub = 'Software\Classes\' },
        @{ Hive = [Microsoft.Win32.Registry]::LocalMachine; Sub = 'SOFTWARE\Classes\' }
    )
    foreach ($progId in $progIds) {
        foreach ($hive in $classesHives) {
            $cmdKey = $null
            try { $cmdKey = $hive.Hive.OpenSubKey($hive.Sub + $progId + '\shell\open\command') } catch { }
            if (-not $cmdKey) { continue }
            try {
                $cmd = [string]$cmdKey.GetValue('')
                if ($cmd) { $null = $list.Add([pscustomobject]@{ Method = '文件关联'; Path = $cmd }) }
            } finally { $cmdKey.Close() }
        }
    }
    foreach ($appPath in @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\wps.exe',
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\wps.exe'
        )) {
        if (-not (Test-Path $appPath)) { continue }
        $val = (Get-ItemProperty -Path $appPath -ErrorAction SilentlyContinue).'(default)'
        if ($val) { $null = $list.Add([pscustomobject]@{ Method = 'App Paths'; Path = [string]$val }) }
    }
    # 4) 正在运行的 WPS 进程
    foreach ($procName in @('wps', 'et', 'wpp', 'ksomisc', 'wpsoffice', 'wpspdf')) {
        Get-Process -Name $procName -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                if ($_.Path) { $null = $list.Add([pscustomobject]@{ Method = ('运行中的进程 ' + $procName + '.exe'); Path = [string]$_.Path }) }
            } catch { }
        }
    }
    # 5) 系统服务（.NET 枚举 Services 注册表）
    $svcKey = $null
    try { $svcKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SYSTEM\CurrentControlSet\Services') } catch { }
    if ($svcKey) {
        try {
            foreach ($svcName in $svcKey.GetSubKeyNames()) {
                $sk = $null
                try { $sk = $svcKey.OpenSubKey($svcName) } catch { }
                if (-not $sk) { continue }
                try {
                    $display = [string]$sk.GetValue('DisplayName')
                    if (($svcName -match 'wps|kingsoft') -or ($display -match 'WPS|Kingsoft')) {
                        $img = [string]$sk.GetValue('ImagePath')
                        if ($img) { $null = $list.Add([pscustomobject]@{ Method = ('系统服务 ' + $svcName); Path = $img }) }
                    }
                } finally { $sk.Close() }
            }
        } finally { $svcKey.Close() }
    }
    # 6) 开始菜单快捷方式
    $wshell = $null
    try { $wshell = New-Object -ComObject WScript.Shell } catch { }
    if ($wshell) {
        foreach ($startDir in @(
                (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'),
                (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs')
            )) {
            if (-not (Test-Path $startDir)) { continue }
            Get-ChildItem -Path $startDir -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match 'WPS|Kingsoft|金山' } |
                ForEach-Object {
                    try {
                        $shortcut = $wshell.CreateShortcut($_.FullName)
                        if ($shortcut.TargetPath) { $null = $list.Add([pscustomobject]@{ Method = ('开始菜单快捷方式 ' + $_.Name); Path = [string]$shortcut.TargetPath }) }
                    } catch { }
                }
        }
    }
    # 7) PATH 中 where.exe
    foreach ($exeName in @('wps.exe', 'ksomisc.exe', 'wpsoffice.exe')) {
        try {
            $whereOut = & where.exe $exeName 2>$null
            foreach ($line in @($whereOut)) {
                if ($line) { $null = $list.Add([pscustomobject]@{ Method = 'PATH 搜索'; Path = [string]$line }) }
            }
        } catch { }
    }
    return $list
}

function Get-ChildDirsFast {
    param([string]$Root, [int]$MaxDepth = 2, [int]$TimeoutSec = 12)
    $out = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Root)) { return $out }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue(@($Root, 0))
    while ($queue.Count -gt 0) {
        if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) { break }
        $item = $queue.Dequeue()
        $dir = [string]$item[0]
        $depth = [int]$item[1]
        if ($depth -ge $MaxDepth) { continue }
        try {
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories($dir)) {
                $null = $out.Add($sub)
                $queue.Enqueue(@($sub, $depth + 1))
            }
        } catch { }
    }
    return $out
}

function Search-WpsOnDisk {
    $found = New-Object System.Collections.ArrayList
    $roots = New-Object System.Collections.ArrayList
    foreach ($pf in @($env:ProgramFiles, [Environment]::GetEnvironmentVariable('ProgramFiles(x86)'))) {
        if (-not $pf) { continue }
        foreach ($sub in @('WPS Office', 'Kingsoft\WPS Office', 'Kingsoft\office6', 'Kingsoft\WPS', 'WPS')) {
            $null = $roots.Add((Join-Path $pf $sub))
        }
    }
    if ($env:LOCALAPPDATA) {
        foreach ($sub in @('Kingsoft\WPS Office', 'WPS Office', 'Programs\WPS Office')) { $null = $roots.Add((Join-Path $env:LOCALAPPDATA $sub)) }
    }
    if ($env:ProgramData) { $null = $roots.Add((Join-Path $env:ProgramData 'Kingsoft\WPS Office')) }
    # 扫描固定磁盘浅层目录（深度 <= 2），只挑名字含 WPS/Kingsoft/金山 的目录（每个磁盘有时间上限）
    foreach ($drive in [System.IO.DriveInfo]::GetDrives()) {
        if ($drive.DriveType -ne [System.IO.DriveType]::Fixed) { continue }
        if (-not $drive.IsReady) { continue }
        $driveRoot = $drive.RootDirectory.FullName
        Write-Log ('  磁盘搜索：' + $driveRoot) 'INFO'
        $dirs = Get-ChildDirsFast -Root $driveRoot -MaxDepth 2 -TimeoutSec 12
        foreach ($d in $dirs) {
            $leaf = [System.IO.Path]::GetFileName([string]$d)
            if ($leaf -match 'WPS|Kingsoft|金山') { $null = $roots.Add([string]$d) }
        }
    }
    $seen = @{}
    foreach ($rootPath in $roots) {
        if ([string]::IsNullOrWhiteSpace([string]$rootPath)) { continue }
        $key = ([string]$rootPath).ToLower()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $layout = Resolve-WpsCandidate -Path ([string]$rootPath)
        if ($layout) {
            $null = $found.Add([pscustomobject]@{ Method = '磁盘搜索'; Path = $layout.Ksomisc })
            break
        }
    }
    return $found
}

function Complete-WpsInfo {
    param($Info, $Layout, [string]$Method, [string]$FallbackVersion)
    $Info.Ksomisc = [string]$Layout.Ksomisc
    $Info.InstallRoot = [string]$Layout.InstallRoot
    $Info.Method = $Method
    $Info.Found = $true
    try {
        $ver = (Get-Item -LiteralPath $Layout.Ksomisc).VersionInfo
        if ($ver -and $ver.ProductVersion) { $Info.Version = [string]$ver.ProductVersion }
    } catch { }
    if (-not $Info.Version) { $Info.Version = $FallbackVersion }
    return $Info
}

function Get-WpsInfo {
    param([string]$ExplicitPath = '')
    $info = [pscustomobject]@{
        InstallRoot = $null
        Version     = $null
        Ksomisc     = $null
        Found       = $false
        Method      = $null
        Searched    = ''
    }
    # 注册表中的版本号（仅用于显示/兜底）
    $regVersion = $null
    foreach ($vk in @(
            'HKCU:\Software\Kingsoft\Office\6.0\common',
            'HKLM:\SOFTWARE\Kingsoft\Office\6.0\Common'
        )) {
        if (-not (Test-Path $vk)) { continue }
        $kv = Get-ItemProperty -Path $vk -ErrorAction SilentlyContinue
        if ($kv -and $kv.Version -and -not $regVersion) { $regVersion = [string]$kv.Version }
    }
    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        Write-Log ('尝试使用 -WpsPath 指定路径：' + $ExplicitPath) 'INFO'
        $layout = Resolve-WpsCandidate -Path $ExplicitPath
        if ($layout) { return (Complete-WpsInfo -Info $info -Layout $layout -Method '手动指定 -WpsPath' -FallbackVersion $regVersion) }
        Write-Log '-WpsPath 指定的路径无效，继续自动搜索...' 'WARN'
    }
    $tried = New-Object System.Collections.ArrayList
    Write-Log '自动查找 WPS 安装位置（注册表/卸载项/文件关联/进程/服务/开始菜单/PATH）...' 'INFO'
    $quickCandidates = Get-WpsQuickCandidates
    foreach ($cand in $quickCandidates) {
        $null = $tried.Add($cand.Method)
        $layout = Resolve-WpsCandidate -Path $cand.Path
        if ($layout) {
            return (Complete-WpsInfo -Info $info -Layout $layout -Method $cand.Method -FallbackVersion $regVersion)
        }
    }
    Write-Log '常规来源未找到 WPS，开始扫描固定磁盘目录（可能需要几十秒）...' 'WARN'
    $diskCandidates = Search-WpsOnDisk
    foreach ($cand in $diskCandidates) {
        $null = $tried.Add($cand.Method)
        $layout = Resolve-WpsCandidate -Path $cand.Path
        if ($layout) {
            return (Complete-WpsInfo -Info $info -Layout $layout -Method $cand.Method -FallbackVersion $regVersion)
        }
    }
    $info.Searched = (($tried | Select-Object -Unique) -join '; ')
    return $info
}

function Test-WpsRunning {
    $names = @('wps', 'et', 'wpp', 'wpspdf', 'wpsoffice', 'wpscenter', 'wpscloudsvr', 'wpscloudlaunch')
    foreach ($n in $names) {
        if (Get-Process -Name $n -ErrorAction SilentlyContinue) { return $true }
    }
    return $false
}

# ----------------------------------------------------------------------------
# 注册表工具
# ----------------------------------------------------------------------------
function Export-RegKey {
    param([string]$RegKey, [string]$FileName)
    if (-not $script:BackupDir) { return $null }
    $file = Join-Path $script:BackupDir $FileName
    $cmd = 'reg export "' + $RegKey + '" "' + $file + '" /y'
    $null = cmd.exe /c $cmd 2>&1
    if (Test-Path $file) { return $FileName }
    Write-Log ('备份注册表失败（已跳过）：' + $RegKey) 'WARN'
    return $null
}

function Set-RegValue {
    param(
        [string]$Path,
        [string]$Name,
        [string]$Value,
        [ValidateSet('String', 'DWord', 'ExpandString')][string]$Type = 'String'
    )
    try {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction Stop | Out-Null }
        if ($Type -eq 'DWord') { $typedValue = [int]$Value } else { $typedValue = [string]$Value }
        New-ItemProperty -Path $Path -Name $Name -Value $typedValue -PropertyType $Type -Force -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Get-RegValueText {
    param([string]$Path, [string]$Name)
    try {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        $prop = $item.PSObject.Properties[$Name]
        if ($prop) { return [string]$prop.Value }
    } catch { }
    return $null
}

# ----------------------------------------------------------------------------
# WPS 配置工具 UI 自动化
# ----------------------------------------------------------------------------
function Get-TargetDefinitions {
    return @(
        [pscustomobject]@{ Group = '在线资源和功能推荐'; Name = '关闭模板推荐'; Partial = '模板推荐' },
        [pscustomobject]@{ Group = '在线资源和功能推荐'; Name = '关闭在线素材资源推荐'; Partial = '在线素材资源推荐' },
        [pscustomobject]@{ Group = '在线资源和功能推荐'; Name = '关闭场景化功能推荐'; Partial = '场景化功能推荐' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS扩展功能菜单项'; Partial = '隐藏WPS扩展功能菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS新建拓展菜单项'; Partial = '隐藏WPS新建拓展菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS云文档功能菜单项'; Partial = '隐藏WPS云文档功能菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS桌面拓展菜单项'; Partial = '隐藏WPS桌面拓展菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS回收站拓展菜单项'; Partial = '隐藏WPS回收站拓展菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS解压拓展菜单项'; Partial = '隐藏WPS解压拓展菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS释放C盘空间功能菜单项'; Partial = '隐藏WPS释放C盘空间功能菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS清理重复文件功能菜单项'; Partial = '隐藏WPS清理重复文件功能菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏编辑拓展菜单项'; Partial = '隐藏编辑拓展菜单项' },
        [pscustomobject]@{ Group = '系统右键菜单'; Name = '隐藏WPS AI拓展菜单项'; Partial = '隐藏WPS AI拓展菜单项' }
    )
}

function Initialize-UiAutomation {
    if ($script:UiReady) { return $true }
    try {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
        Add-Type -AssemblyName UIAutomationTypes -ErrorAction Stop
        $script:UiReady = $true
        return $true
    } catch {
        Write-Log ('Windows UI 自动化组件加载失败：' + $_.Exception.Message) 'WARN'
        $script:UiReady = $false
        return $false
    }
}

function Wait-WpsConfigDialog {
    param([int]$TimeoutSec = 120)
    try { $root = [System.Windows.Automation.AutomationElement]::RootElement } catch { return $null }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            $wins = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
            foreach ($w in $wins) {
                try { if ($w.Current.Name -eq 'WPS Office 配置工具') { return $w } } catch { }
            }
        } catch { }
        Start-Sleep -Milliseconds 500
    }
    return $null
}

function Open-WpsConfigTool {
    param([string]$KsomiscPath)
    if ([string]::IsNullOrWhiteSpace($KsomiscPath) -or (-not (Test-Path $KsomiscPath))) {
        throw ('未找到 WPS 配置工具 ksomisc.exe：' + $KsomiscPath + '（可尝试用 -WpsPath 手动指定 WPS 安装目录后重试）')
    }
    $proc = Start-Process -FilePath $KsomiscPath -ArgumentList '-functioncustomization' -PassThru -ErrorAction Stop
    $win = Wait-WpsConfigDialog -TimeoutSec 120
    if (-not $win) {
        try { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } catch { }
        throw '等待“WPS Office 配置工具”窗口超时（可能被安全软件拦截，或当前 WPS 版本不支持 -functioncustomization）'
    }
    return [pscustomobject]@{ ProcessId = $proc.Id; Window = $win }
}

function Get-WpsConfigCheckboxMap {
    param($Window)
    $map = @{}
    if ($null -eq $Window) { return $map }
    try {
        $all = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    } catch { return $map }
    foreach ($e in $all) {
        try {
            $c = $e.Current
            if ($c.ControlType.ProgrammaticName -ne 'ControlType.CheckBox') { continue }
            if ([string]::IsNullOrWhiteSpace($c.Name)) { continue }
            if (-not $map.ContainsKey($c.Name)) { $map[$c.Name] = $e }
        } catch { }
    }
    return $map
}

function Get-ToggleStateText {
    param($Element)
    try {
        $tp = $Element.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
        return [string]$tp.Current.ToggleState
    } catch { return 'Unknown' }
}

function Find-TargetElement {
    param($Map, $Definition)
    if ($Map.ContainsKey($Definition.Name)) { return $Map[$Definition.Name] }
    foreach ($k in $Map.Keys) {
        if ($k -like ('*' + $Definition.Partial + '*')) { return $Map[$k] }
    }
    return $null
}

function Find-DialogButton {
    param($Window, [string]$NameRegex)
    try {
        $all = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    } catch { return $null }
    foreach ($e in $all) {
        try {
            $c = $e.Current
            if ($c.ControlType.ProgrammaticName -eq 'ControlType.Button' -and $c.Name -match $NameRegex) { return $e }
        } catch { }
    }
    return $null
}

function Invoke-AutomationElement {
    param($Element)
    try {
        $pattern = $Element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        $pattern.Invoke()
        return $true
    } catch { return $false }
}

function Set-CheckboxOn {
    param($Element)
    $before = Get-ToggleStateText $Element
    if ($before -eq 'On') { return 'already' }
    try {
        $tp = $Element.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
        $tp.Toggle()
        Start-Sleep -Milliseconds 180
        $after = Get-ToggleStateText $Element
        if ($after -eq 'On') { return 'toggled' } else { return ('unexpected:' + $after) }
    } catch {
        return ('failed:' + $_.Exception.Message)
    }
}

function Resolve-WpsConfigDialogs {
    param($MainWindow, [int[]]$ProcessIds, [int]$MaxSeconds = 120)
    $deadline = (Get-Date).AddSeconds($MaxSeconds)
    $mainHandle = 0
    try { $mainHandle = $MainWindow.Current.NativeWindowHandle } catch { }
    $clicked = 0
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 600
        $alive = @()
        foreach ($procId in $ProcessIds) {
            if (Get-Process -Id $procId -ErrorAction SilentlyContinue) { $alive += $procId }
        }
        if ($alive.Count -eq 0) { break }
        try {
            $root = [System.Windows.Automation.AutomationElement]::RootElement
            $wins = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
            foreach ($w in $wins) {
                try {
                    $c = $w.Current
                    if ($ProcessIds -notcontains $c.ProcessId) { continue }
                    if ($c.NativeWindowHandle -eq $mainHandle) { continue }
                    $btn = Find-DialogButton -Window $w -NameRegex '^(确定|是|OK|Yes)$'
                    if ($btn) {
                        if (Invoke-AutomationElement $btn) {
                            $clicked++
                            Start-Sleep -Milliseconds 900
                        }
                    }
                } catch { }
            }
        } catch { }
    }
    return $clicked
}

function Close-WpsConfigTool {
    param([int[]]$ProcessIds)
    foreach ($procId in $ProcessIds) {
        try { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue } catch { }
    }
}

function Get-WpsConfigStates {
    param([string]$KsomiscPath)
    if (-not (Initialize-UiAutomation)) { return $null }
    $dlg = $null
    try {
        $dlg = Open-WpsConfigTool -KsomiscPath $KsomiscPath
    } catch {
        Write-Log ('读取 WPS 配置工具状态失败：' + $_.Exception.Message) 'WARN'
        return $null
    }
    Start-Sleep -Milliseconds 800
    $map = Get-WpsConfigCheckboxMap -Window $dlg.Window
    $states = @{}
    foreach ($t in (Get-TargetDefinitions)) {
        $el = Find-TargetElement -Map $map -Definition $t
        if ($el) { $states[$t.Name] = Get-ToggleStateText $el } else { $states[$t.Name] = 'NotFound' }
    }
    Close-WpsConfigTool -ProcessIds @($dlg.ProcessId)
    Start-Sleep -Milliseconds 400
    return $states
}

function Invoke-WpsConfigCleanup {
    param([string]$KsomiscPath)
    $result = [pscustomobject]@{
        Supported = $false
        Applied   = $false
        Items     = @()
        Verified  = $null
        Message   = ''
    }
    if (-not (Initialize-UiAutomation)) { $result.Message = 'UI 自动化不可用'; return $result }
    $dlg = $null
    try {
        $dlg = Open-WpsConfigTool -KsomiscPath $KsomiscPath
    } catch {
        $result.Message = $_.Exception.Message
        return $result
    }
    $result.Supported = $true
    $procIds = @($dlg.ProcessId)
    Start-Sleep -Milliseconds 800
    $map = Get-WpsConfigCheckboxMap -Window $dlg.Window
    if ($map.Count -eq 0) {
        $result.Message = '配置工具窗口中没有找到任何复选框'
        Close-WpsConfigTool -ProcessIds $procIds
        return $result
    }
    $items = New-Object System.Collections.ArrayList
    foreach ($t in (Get-TargetDefinitions)) {
        $entry = [pscustomobject]@{ Group = $t.Group; Name = $t.Name; Before = 'NotFound'; Action = 'not-found'; After = 'NotFound' }
        $el = Find-TargetElement -Map $map -Definition $t
        if (-not $el) { $null = $items.Add($entry); continue }
        $entry.Before = Get-ToggleStateText $el
        if ($entry.Before -eq 'On') {
            $entry.Action = 'already-on'
            $entry.After = 'On'
        } else {
            $r = Set-CheckboxOn $el
            if ($r -eq 'toggled') { $entry.Action = 'toggled'; $entry.After = 'On' }
            elseif ($r -eq 'already') { $entry.Action = 'already-on'; $entry.After = 'On' }
            else { $entry.Action = 'toggle-failed'; $entry.After = $entry.Before }
        }
        $null = $items.Add($entry)
    }
    $result.Items = $items
    $ok = Find-DialogButton -Window $dlg.Window -NameRegex '^确定$'
    if (-not $ok) {
        $result.Message = '未找到“确定”按钮'
        Close-WpsConfigTool -ProcessIds $procIds
        return $result
    }
    if (Invoke-AutomationElement $ok) {
        Write-Log '已点击“确定”，正在等待 WPS 配置工具应用设置（最多 120 秒）...' 'INFO'
        $null = Resolve-WpsConfigDialogs -MainWindow $dlg.Window -ProcessIds $procIds -MaxSeconds 120
        $result.Applied = $true
    } else {
        $result.Message = '点击“确定”失败'
    }
    Start-Sleep -Seconds 2
    Close-WpsConfigTool -ProcessIds $procIds
    Start-Sleep -Seconds 3
    $result.Verified = Get-WpsConfigStates -KsomiscPath $KsomiscPath
    return $result
}

function Invoke-MenuItemsRegistryFallback {
    $ops = @(
        @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\Common\ksomisc'; Name = 'VasContextMenuShellNewUser'; Value = 'false'; Type = 'String' },
        @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\kwpsshellext'; Name = 'disableDocPropPage'; Value = '1'; Type = 'DWord' },
        @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\kwpsshellext\contextmenu\photo'; Name = 'supportedit'; Value = 'false'; Type = 'String' },
        @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kcdiskclean'; Name = 'contextMenuVisible'; Value = '0'; Type = 'DWord' },
        @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kcdiskclean'; Name = 'wpsSceneRecEnable'; Value = 'false'; Type = 'String' }
    )
    $okCount = 0
    foreach ($op in $ops) {
        if (Set-RegValue -Path $op.Path -Name $op.Name -Value $op.Value -Type $op.Type) { $okCount++ }
    }
    $shellNewKeys = @(
        'HKCU:\Software\Classes\.pdfwpsshellnew',
        'HKCU:\Software\Classes\.pptxwpsaicreateshellnew',
        'HKCU:\Software\Classes\WPP.PPTX.AI'
    )
    $removed = 0
    foreach ($k in $shellNewKeys) {
        if (Test-Path $k) {
            Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue
            $removed++
        }
    }
    return [pscustomobject]@{ Ops = $okCount; RemovedKeys = $removed }
}

# ----------------------------------------------------------------------------
# 自启动检测
# ----------------------------------------------------------------------------
function Get-AutostartRunEntries {
    $targets = @(
        @{ Hive = 'HKCU'; Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'; RegPath = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Run' },
        @{ Hive = 'HKCU'; Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce'; RegPath = 'HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce' },
        @{ Hive = 'HKLM'; Path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'; RegPath = 'HKLM\Software\Microsoft\Windows\CurrentVersion\Run' },
        @{ Hive = 'HKLM'; Path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce'; RegPath = 'HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce' },
        @{ Hive = 'HKLM32'; Path = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; RegPath = 'HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run' },
        @{ Hive = 'HKLM32'; Path = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce'; RegPath = 'HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce' }
    )
    $pattern = 'wps|kingsoft|kwps|kso|qingbangong|wpsbox'
    $list = New-Object System.Collections.ArrayList
    foreach ($t in $targets) {
        if (-not (Test-Path $t.Path)) { continue }
        $item = Get-ItemProperty -Path $t.Path -ErrorAction SilentlyContinue
        if (-not $item) { continue }
        foreach ($prop in $item.PSObject.Properties) {
            if ($prop.Name -like 'PS*') { continue }
            $name = [string]$prop.Name
            $value = [string]$prop.Value
            if (($name -match $pattern) -or ($value -match $pattern)) {
                $null = $list.Add([pscustomobject]@{ Hive = $t.Hive; RegPath = $t.RegPath; Path = $t.Path; Name = $name; Value = $value })
            }
        }
    }
    return $list
}

function Get-WpsStartupShortcuts {
    $dirs = @(
        (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'),
        (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup')
    )
    $pattern = 'wps|kingsoft|kwps|kso|qingbangong'
    $list = New-Object System.Collections.ArrayList
    foreach ($d in $dirs) {
        if (-not (Test-Path $d)) { continue }
        Get-ChildItem -Path $d -File -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.Name -match $pattern) {
                $null = $list.Add([pscustomobject]@{ Directory = $d; File = $_.FullName; Name = $_.Name })
            }
        }
    }
    return $list
}

function Get-WpsScheduledTasks {
    $list = New-Object System.Collections.ArrayList
    try {
        $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
            $n = $_.TaskName
            ($n -like 'WpsUpdateTask*') -or ($n -like 'WpsUpdateLogonTask*') -or ($n -like 'WpsWakeWnsLogonTask*') -or ($n -like 'WPS*Update*')
        }
        foreach ($t in $tasks) { $null = $list.Add($t) }
    } catch { }
    return $list
}

function Get-WpsServices {
    $list = New-Object System.Collections.ArrayList
    try {
        Get-Service -ErrorAction SilentlyContinue | Where-Object {
            ($_.Name -match '^(wps|kingsoft)') -or ($_.DisplayName -match 'WPS Office|Kingsoft')
        } | ForEach-Object { $null = $list.Add($_) }
    } catch { }
    return $list
}

# ----------------------------------------------------------------------------
# 资源管理器 WPS 入口（WPS 云盘图标 / 云盘右键扩展）
# 说明：全部使用 .NET 注册表 API 实现，避免 PowerShell 注册表 Cmdlet 在高延迟环境下过慢
# ----------------------------------------------------------------------------
function Get-RegistryHiveObject {
    param([string]$HiveName)
    switch ($HiveName) {
        'HKCU' { return [Microsoft.Win32.Registry]::CurrentUser }
        'HKLM' { return [Microsoft.Win32.Registry]::LocalMachine }
        'HKLM32' { return [Microsoft.Win32.Registry]::LocalMachine }
        default { return $null }
    }
}

function Get-ClsidHintsFast {
    param([string]$Clsid)
    if ([string]::IsNullOrWhiteSpace($Clsid)) { return '' }
    $parts = New-Object System.Collections.ArrayList
    $roots = @(
        @{ Hive = [Microsoft.Win32.Registry]::CurrentUser; Sub = 'Software\Classes\CLSID\' },
        @{ Hive = [Microsoft.Win32.Registry]::LocalMachine; Sub = 'SOFTWARE\Classes\CLSID\' },
        @{ Hive = [Microsoft.Win32.Registry]::LocalMachine; Sub = 'SOFTWARE\WOW6432Node\Classes\CLSID\' }
    )
    foreach ($root in $roots) {
        $key = $null
        try { $key = $root.Hive.OpenSubKey($root.Sub + $Clsid) } catch { }
        if (-not $key) { continue }
        try {
            $null = $parts.Add([string]$key.GetValue(''))
            foreach ($subName in @('DefaultIcon', 'InProcServer32', 'LocalServer32')) {
                $sk = $null
                try { $sk = $key.OpenSubKey($subName) } catch { }
                if ($sk) {
                    $null = $parts.Add([string]$sk.GetValue(''))
                    $sk.Close()
                }
            }
            $inst = $null
            try { $inst = $key.OpenSubKey('Instance\InitPropertyBag') } catch { }
            if ($inst) {
                $null = $parts.Add([string]$inst.GetValue('TargetFolderPath'))
                $inst.Close()
            }
        } finally {
            $key.Close()
        }
    }
    return ($parts -join ' ')
}

function Get-WpsExplorerNamespaceEntries {
    $scopes = @(
        @{ Hive = 'HKCU'; Sub = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'; RegRoot = 'HKCU'; Label = '桌面图标' },
        @{ Hive = 'HKCU'; Sub = 'Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'; RegRoot = 'HKCU'; Label = '此电脑' },
        @{ Hive = 'HKLM'; Sub = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'; RegRoot = 'HKLM'; Label = '桌面图标' },
        @{ Hive = 'HKLM'; Sub = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'; RegRoot = 'HKLM'; Label = '此电脑' },
        @{ Hive = 'HKLM32'; Sub = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace'; RegRoot = 'HKLM'; Label = '桌面图标' },
        @{ Hive = 'HKLM32'; Sub = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'; RegRoot = 'HKLM'; Label = '此电脑' }
    )
    $pattern = 'wps|kingsoft|云盘|wpsdrive|qingbangong|qingnse'
    $list = New-Object System.Collections.ArrayList
    foreach ($scope in $scopes) {
        $hive = Get-RegistryHiveObject -HiveName $scope.Hive
        if (-not $hive) { continue }
        $key = $null
        try { $key = $hive.OpenSubKey($scope.Sub) } catch { }
        if (-not $key) { continue }
        try {
            foreach ($sid in $key.GetSubKeyNames()) {
                if ($sid -notmatch '^\{[0-9A-Fa-f-]+\}$') { continue }
                $display = ''
                $sk = $null
                try { $sk = $key.OpenSubKey($sid) } catch { }
                if ($sk) {
                    try { $display = [string]$sk.GetValue('') } catch { }
                    $sk.Close()
                }
                $hints = Get-ClsidHintsFast -Clsid $sid
                if (($display -match $pattern) -or ($hints -match $pattern)) {
                    $subPath = $scope.Sub + '\' + $sid
                    $null = $list.Add([pscustomobject]@{
                        Type    = 'NameSpace'
                        Hive    = $scope.Hive
                        SubPath = $subPath
                        RegPath = $scope.RegRoot + '\' + $subPath
                        Label   = $scope.Label + '「' + $(if ($display) { $display } else { $sid }) + '」'
                        Name    = $sid
                        Display = $display
                    })
                }
            }
        } finally {
            $key.Close()
        }
    }
    return $list
}

function Get-WpsExplorerShellHandlers {
    $bases = @('*', 'Directory', 'Directory\Background', 'Drive', 'Folder', 'AllFilesystemObjects', 'lnkfile', 'DesktopBackground')
    $roots = @(
        @{ Hive = 'HKCU'; Sub = 'Software\Classes'; RegRoot = 'HKCU' },
        @{ Hive = 'HKLM'; Sub = 'SOFTWARE\Classes'; RegRoot = 'HKLM' },
        @{ Hive = 'HKLM32'; Sub = 'SOFTWARE\WOW6432Node\Classes'; RegRoot = 'HKLM' }
    )
    $namePattern = 'wps|kingsoft|qing|nsemenu'
    $hintPattern = 'wps|kingsoft|qingnse|qingshellext|kwpsshellext|kdesktopshellext|qingdriveshellext|wpscloud|wpsdrive'
    $list = New-Object System.Collections.ArrayList
    foreach ($root in $roots) {
        $hive = Get-RegistryHiveObject -HiveName $root.Hive
        if (-not $hive) { continue }
        foreach ($base in $bases) {
            $sub = $root.Sub + '\' + $base + '\shellex\ContextMenuHandlers'
            $key = $null
            try { $key = $hive.OpenSubKey($sub) } catch { }
            if (-not $key) { continue }
            try {
                foreach ($handlerName in $key.GetSubKeyNames()) {
                    $clsid = ''
                    $sk = $null
                    try { $sk = $key.OpenSubKey($handlerName) } catch { }
                    if ($sk) {
                        try { $clsid = [string]$sk.GetValue('') } catch { }
                        $sk.Close()
                    }
                    $hints = ''
                    if ($clsid -match '^\{[0-9A-Fa-f-]+\}$') { $hints = Get-ClsidHintsFast -Clsid $clsid }
                    if (($handlerName -match $namePattern) -or ($hints -match $hintPattern)) {
                        $subPath = $sub + '\' + $handlerName
                        $null = $list.Add([pscustomobject]@{
                            Type    = 'ShellHandler'
                            Hive    = $root.Hive
                            SubPath = $subPath
                            RegPath = $root.RegRoot + '\' + $subPath
                            Label   = '右键扩展「' + $handlerName + '」'
                            Name    = $handlerName
                            Display = $base
                        })
                    }
                }
            } finally {
                $key.Close()
            }
        }
    }
    return $list
}

function Remove-ExplorerEntryKey {
    param($Entry)
    $hive = Get-RegistryHiveObject -HiveName $Entry.Hive
    if (-not $hive) { return $false }
    $idx = $Entry.SubPath.LastIndexOf('\')
    if ($idx -lt 1) { return $false }
    $parentPath = $Entry.SubPath.Substring(0, $idx)
    $childName = $Entry.SubPath.Substring($idx + 1)
    $pk = $null
    try { $pk = $hive.OpenSubKey($parentPath, $true) } catch { }
    if (-not $pk) { return $false }
    try {
        $pk.DeleteSubKeyTree($childName, $false)
        return $true
    } catch {
        return $false
    } finally {
        $pk.Close()
    }
}

function Invoke-ExplorerEntryCleanup {
    $nsEntries = Get-WpsExplorerNamespaceEntries
    $shEntries = Get-WpsExplorerShellHandlers
    $allEntries = @()
    foreach ($e in $nsEntries) { $allEntries += $e }
    foreach ($e in $shEntries) { $allEntries += $e }
    if ($allEntries.Count -eq 0) {
        Write-Log '未发现资源管理器中的 WPS 云盘图标 / 云盘右键扩展注册。' 'INFO'
        Add-Result -Category '资源管理器入口' -Item 'WPS 云盘图标 / 云盘右键扩展' -Status '未发现' -Detail ''
        return
    }
    $records = New-Object System.Collections.ArrayList
    $okCount = 0
    $failCount = 0
    foreach ($entry in $allEntries) {
        $safeName = 'explorer-' + ($entry.RegPath -replace '[\\/:*?"<>|{}]', '_')
        if ($safeName.Length -gt 140) { $safeName = $safeName.Substring(0, 140) }
        $regFile = $safeName + '.reg'
        $fullBackup = Join-Path $script:BackupDir $regFile
        $null = cmd.exe /c ('reg export "' + $entry.RegPath + '" "' + $fullBackup + '" /y') 2>&1
        if (-not (Test-Path $fullBackup)) {
            $null = Export-RegKey -RegKey $entry.RegPath -FileName $regFile
        }
        $status = '失败'
        if (Remove-ExplorerEntryKey -Entry $entry) {
            $status = '已移除'
            $okCount++
            Write-Log ('已移除资源管理器 WPS 入口：' + $entry.Label) 'OK'
        } else {
            $failCount++
            Write-Log ('移除失败（权限不足或键不存在）：' + $entry.RegPath) 'WARN'
        }
        if (Test-Path (Join-Path $script:BackupDir $regFile)) {
            $null = $records.Add([ordered]@{
                Type    = $entry.Type
                Label   = $entry.Label
                RegPath = $entry.RegPath
                RegFile = $regFile
            })
        }
        Add-Result -Category '资源管理器入口' -Item $entry.Label -Status $status -Detail $entry.RegPath
    }
    if ($script:BackupDir -and $records.Count -gt 0) {
        try {
            $records | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $script:BackupDir 'explorer-entries.json') -Encoding UTF8
        } catch { }
    }
    Write-Log ('资源管理器入口清理完成：成功 ' + $okCount + ' 项，失败 ' + $failCount + ' 项。重启资源管理器或重启电脑后生效。') 'INFO'
}

function Restart-ExplorerSafe {
    if (-not $RestartExplorer) { return }
    Write-Log '正在重启资源管理器以使改动立即生效...' 'INFO'
    try {
        Stop-Process -Name explorer -Force -ErrorAction Stop
        Start-Sleep -Seconds 4
        if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) {
            Start-Process -FilePath 'explorer.exe' -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
        }
        if (Get-Process -Name explorer -ErrorAction SilentlyContinue) {
            Write-Log '资源管理器已重启。' 'OK'
        } else {
            Write-Log '资源管理器未自动恢复，请按 Ctrl+Shift+Esc 打开任务管理器 → 运行新任务 → explorer.exe。' 'WARN'
        }
    } catch {
        Write-Log ('重启资源管理器失败：' + $_.Exception.Message + '，请手动重启。') 'WARN'
    }
}
# ----------------------------------------------------------------------------
# 检测输出
# ----------------------------------------------------------------------------
function Show-CurrentState {
    Write-Log '正在读取 WPS 安装信息...' 'INFO'
    $info = Get-WpsInfo -ExplicitPath $WpsPath
    if (-not $info.Found) {
        Write-Log '未检测到 WPS Office 安装信息，后续步骤可能无法执行。' 'WARN'
        if ($info.Searched) { Write-Log ('已尝试的查找方式：' + $info.Searched) 'WARN' }
        Write-Log '可手动指定路径重试，例如：WPS一键优化.bat -Mode Scan -WpsPath "D:\Apps\WPS Office\12.1.0.29690"' 'WARN'
    } else {
        Write-Log ('WPS 安装目录：' + $info.InstallRoot) 'INFO'
        Write-Log ('WPS 版本：' + $info.Version) 'INFO'
        Write-Log ('配置工具：' + $info.Ksomisc) 'INFO'
        if ($info.Method) { Write-Log ('发现方式：' + $info.Method) 'INFO' }
    }
    if (Test-WpsRunning) {
        Write-Log '检测到 WPS 相关进程正在运行：为保证设置不被临时覆盖，建议保存文档并关闭 WPS 后运行 Apply / Revert。' 'WARN'
    }
    Write-Log '检查 WPS 配置工具“功能定制”13 项状态...' 'INFO'
    $states = $null
    if (-not $SkipUiAutomation) { $states = Get-WpsConfigStates -KsomiscPath $info.Ksomisc }
    if ($states) {
        $rows = @()
        foreach ($t in (Get-TargetDefinitions)) {
            $st = $states[$t.Name]
            $cn = '未知'
            if ($st -eq 'On') { $cn = '已勾选（已关闭/已隐藏）' }
            elseif ($st -eq 'Off') { $cn = '未勾选（仍会显示）' }
            elseif ($st -eq 'NotFound') { $cn = '未找到（WPS 版本差异）' }
            else { $cn = [string]$st }
            $rows += [pscustomobject]@{ 分组 = $t.Group; 设置项 = $t.Name; 状态 = $cn }
        }
        foreach ($row in $rows) {
            Write-CliLine ('  [' + $row.'分组' + '] ' + $row.'设置项' + '：' + $row.'状态')
        }
    } else {
        Write-Log '无法通过 UI 自动化读取“功能定制”状态（可加 -SkipUiAutomation 跳过）。' 'WARN'
    }

    Write-Log '检查免强制登录开关...' 'INFO'
    $flogin = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\officespace\flogin'
    $v1 = Get-RegValueText -Path $flogin -Name 'enableForceLoginForFirstInstallDevice'
    $v2 = Get-RegValueText -Path $flogin -Name 'enableForceLoginForHasInstallDevice'
    if ($null -eq $v1) { $v1 = '（未设置）' }
    if ($null -eq $v2) { $v2 = '（未设置）' }
    Write-Log ('  enableForceLoginForFirstInstallDevice = ' + $v1) 'INFO'
    Write-Log ('  enableForceLoginForHasInstallDevice   = ' + $v2) 'INFO'

    Write-Log '检查 WPS 网盘 / 文件传输助手状态...' 'INFO'
    $wpscloud = 'HKCU:\Software\Kingsoft\wpscloud'
    foreach ($name in @('fileTransfer', 'UploadSync', 'DesktopSync', 'SendBy', 'kicFileTransferSwitch')) {
        $v = Get-RegValueText -Path $wpscloud -Name $name
        if ($null -ne $v) { Write-Log ('  wpscloud\' + $name + ' = ' + $v) 'INFO' }
    }
    $mini = Join-Path $wpscloud 'miniTransferAssistant'
    $g = Get-RegValueText -Path $mini -Name 'globalEnable'
    if ($null -ne $g) { Write-Log ('  wpscloud\miniTransferAssistant\globalEnable = ' + $g) 'INFO' }

    Write-Log '检查「稻壳AI模板」启动页开关...' 'INFO'
    $docerTab = Get-RegValueText -Path 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kdocerpage' -Name 'docerTabShow'
    if ($null -eq $docerTab) {
        Write-Log '  docerTabShow = （未设置，默认会展示稻壳AI模板页）' 'INFO'
    } elseif ([string]$docerTab -eq '0') {
        Write-Log '  docerTabShow = 0（已关闭，启动时不再自动打开稻壳AI模板页）' 'INFO'
    } else {
        Write-Log ('  docerTabShow = ' + $docerTab + '（会默认展示稻壳AI模板页）') 'INFO'
    }

    Write-Log '检查首页右侧面板（精选推荐/功能推荐）开关...' 'INFO'
    $panelValue = Get-RegValueText -Path 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kstartpage\rightAppInfo' -Name 'rightPanelWidgetVisibleGlobal'
    if ($null -eq $panelValue) {
        Write-Log '  rightPanelWidgetVisibleGlobal = （未设置，默认显示右侧面板）' 'INFO'
    } elseif ([string]$panelValue -eq 'false') {
        Write-Log '  rightPanelWidgetVisibleGlobal = false（已关闭，首页不显示右侧面板）' 'INFO'
    } else {
        Write-Log ('  rightPanelWidgetVisibleGlobal = ' + $panelValue + '（会显示右侧面板）') 'INFO'
    }

    Write-Log '检查自启动项...' 'INFO'
    $tasks = Get-WpsScheduledTasks
    if ($tasks.Count -eq 0) {
        Write-Log '  计划任务：无 WPS 更新任务' 'INFO'
    } else {
        foreach ($t in $tasks) { Write-Log ('  计划任务：' + $t.TaskName + ' 状态=' + $t.State) 'INFO' }
    }
    $services = Get-WpsServices
    if ($services.Count -eq 0) {
        Write-Log '  服务：无 WPS 相关服务' 'INFO'
    } else {
        foreach ($s in $services) { Write-Log ('  服务：' + $s.Name + ' 状态=' + $s.Status + ' 启动类型=' + $s.StartType) 'INFO' }
    }
    $runs = Get-AutostartRunEntries
    if ($runs.Count -eq 0) {
        Write-Log '  Run/RunOnce：无 WPS 相关启动项' 'INFO'
    } else {
        foreach ($r in $runs) { Write-Log ('  ' + $r.RegPath + ' -> ' + $r.Name + ' = ' + $r.Value) 'INFO' }
    }
    $links = Get-WpsStartupShortcuts
    if ($links.Count -eq 0) {
        Write-Log '  启动文件夹：无 WPS 相关快捷方式' 'INFO'
    } else {
        foreach ($l in $links) { Write-Log ('  ' + $l.File) 'INFO' }
    }
    Write-Log '检查资源管理器中的 WPS 入口（WPS 云盘图标 / 云盘右键扩展）...' 'INFO'
    $nsEntries = Get-WpsExplorerNamespaceEntries
    $shEntries = Get-WpsExplorerShellHandlers
    if (($nsEntries.Count + $shEntries.Count) -eq 0) {
        Write-Log '  未发现 WPS 云盘图标 / 云盘右键扩展注册' 'INFO'
    } else {
        foreach ($e in $nsEntries) { Write-Log ('  云盘图标：' + $e.RegPath + '（' + $e.Display + '）') 'INFO' }
        foreach ($e in $shEntries) { Write-Log ('  右键扩展：' + $e.RegPath) 'INFO' }
    }
    Write-Log '只读检测完成（未做任何修改）。' 'OK'
}

# ----------------------------------------------------------------------------
# 备份
# ----------------------------------------------------------------------------
function New-OptimizeBackup {
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $dir = Join-Path $script:BackupRoot ('wps-opt-' + $ts)
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $script:BackupDir = $dir
    Write-Log ('正在创建备份：' + $dir) 'INFO'
    $manifest = [ordered]@{
        CreatedAt       = (Get-Date).ToString('s')
        ScriptVersion   = $script:ScriptVersion
        Computer        = $env:COMPUTERNAME
        User            = $env:USERNAME
        Wps             = $script:WpsInfo
        RegistryBackups = @()
        RunValues       = @()
        StartupLinks    = @()
        Tasks           = @()
        Service         = $null
    }
    $regTargets = @(
        @{ Key = 'HKCU\Software\Kingsoft'; File = 'HKCU-Software-Kingsoft.reg' },
        @{ Key = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Run'; File = 'HKCU-Run.reg' },
        @{ Key = 'HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce'; File = 'HKCU-RunOnce.reg' },
        @{ Key = 'HKLM\Software\Microsoft\Windows\CurrentVersion\Run'; File = 'HKLM-Run.reg' },
        @{ Key = 'HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; File = 'HKLM-WOW-Run.reg' }
    )
    $regFiles = @()
    foreach ($rt in $regTargets) {
        $f = Export-RegKey -RegKey $rt.Key -FileName $rt.File
        if ($f) { $regFiles += $f }
    }
    $manifest.RegistryBackups = $regFiles

    foreach ($r in (Get-AutostartRunEntries)) {
        $manifest.RunValues += [ordered]@{ RegPath = $r.RegPath; Path = $r.Path; Name = $r.Name; Value = $r.Value }
    }
    $startupBackupDir = Join-Path $dir 'startup'
    foreach ($l in (Get-WpsStartupShortcuts)) {
        New-Item -ItemType Directory -Force -Path $startupBackupDir | Out-Null
        $dest = Join-Path $startupBackupDir $l.Name
        try {
            Move-Item -Path $l.File -Destination $dest -Force -ErrorAction Stop
            $manifest.StartupLinks += [ordered]@{ Original = $l.File; Backup = $dest }
            Write-Log ('已把启动文件夹项移到备份：' + $l.Name) 'OK'
        } catch {
            Write-Log ('移动启动项失败：' + $l.File) 'WARN'
        }
    }
    foreach ($t in (Get-WpsScheduledTasks)) {
        $xmlFile = 'task-' + ($t.TaskName -replace '[\\/:*?"<>|]', '_') + '.xml'
        $xmlPath = Join-Path $dir $xmlFile
        try {
            $xml = Export-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop
            Set-Content -Path $xmlPath -Value $xml -Encoding UTF8 -ErrorAction Stop
            $manifest.Tasks += [ordered]@{ Name = $t.TaskName; TaskPath = $t.TaskPath; State = [string]$t.State; XmlFile = $xmlFile }
        } catch {
            Write-Log ('导出计划任务失败：' + $t.TaskName) 'WARN'
        }
    }
    foreach ($s in (Get-WpsServices)) {
        $manifest.Service = [ordered]@{ Name = $s.Name; StartType = [string]$s.StartType; Status = [string]$s.Status }
        $svcKey = 'HKLM\SYSTEM\CurrentControlSet\Services\' + $s.Name
        $null = Export-RegKey -RegKey $svcKey -FileName ('service-' + $s.Name + '.reg')
    }
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -Path (Join-Path $dir 'manifest.json') -Encoding UTF8
    Write-Log ('备份完成：' + $dir) 'OK'
    return $dir
}

# ----------------------------------------------------------------------------
# 优化动作
# ----------------------------------------------------------------------------
function Invoke-LoginHardening {
    Write-Log '写入“免强制登录”开关...' 'INFO'
    $flogin = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\officespace\flogin'
    $beforeFirst = Get-RegValueText -Path $flogin -Name 'enableForceLoginForFirstInstallDevice'
    $beforeHas = Get-RegValueText -Path $flogin -Name 'enableForceLoginForHasInstallDevice'
    $ok1 = Set-RegValue -Path $flogin -Name 'enableForceLoginForFirstInstallDevice' -Value 'false' -Type String
    $ok2 = Set-RegValue -Path $flogin -Name 'enableForceLoginForHasInstallDevice' -Value 'false' -Type String
    if ($null -eq $beforeFirst) { $null = $script:CreatedValues.Add([ordered]@{ Path = $flogin; Name = 'enableForceLoginForFirstInstallDevice'; Value = 'false' }) }
    if ($null -eq $beforeHas) { $null = $script:CreatedValues.Add([ordered]@{ Path = $flogin; Name = 'enableForceLoginForHasInstallDevice'; Value = 'false' }) }
    if ($ok1 -and $ok2) {
        Write-Log '已关闭“首次安装设备强制登录 / 已安装设备强制登录”。' 'OK'
        Add-Result -Category '免登录' -Item '跳过强制登录（首次安装设备/已安装设备）' -Status '已写入' -Detail 'HKCU\...\plugins\officespace\flogin'
    } else {
        Write-Log '写入免登录开关失败（可能被安全软件拦截）。' 'WARN'
        Add-Result -Category '免登录' -Item '跳过强制登录' -Status '失败' -Detail '写入注册表失败'
    }
}

function Invoke-CloudHardening {
    Write-Log '关闭 WPS 网盘 / 文件传输助手相关开关...' 'INFO'
    $wpscloud = 'HKCU:\Software\Kingsoft\wpscloud'
    if (-not (Test-Path $wpscloud)) {
        Write-Log '未找到 HKCU\Software\Kingsoft\wpscloud，跳过。' 'WARN'
        Add-Result -Category 'WPS网盘/传输助手' -Item '云盘开关' -Status '跳过' -Detail '注册表键不存在'
        return
    }
    $switchNames = @(
        'fileTransfer', 'uploadFile', 'shareFile', 'SendBy', 'DesktopSync', 'SpecifyFolderSync', 'UploadSync',
        'kicFileTransferSwitch', 'kicUploadFileSwitch', 'kicDesktopSyncSwitch', 'kicShareFileSwitch', 'kicSendBySwitch', 'kicUploadSyncSwitch'
    )
    $changed = 0
    foreach ($name in $switchNames) {
        $old = Get-RegValueText -Path $wpscloud -Name $name
        if ($null -eq $old) { continue }
        if ($old -ne 'false') {
            if (Set-RegValue -Path $wpscloud -Name $name -Value 'false' -Type String) { $changed++ }
        }
    }
    $mini = Join-Path $wpscloud 'miniTransferAssistant'
    if (Test-Path $mini) {
        $beforeGlobal = Get-RegValueText -Path $mini -Name 'globalEnable'
        $beforeDefault = Get-RegValueText -Path $mini -Name 'defaultEnable'
        $null = Set-RegValue -Path $mini -Name 'globalEnable' -Value 'false' -Type String
        $null = Set-RegValue -Path $mini -Name 'defaultEnable' -Value 'false' -Type String
        if ($null -eq $beforeGlobal) { $null = $script:CreatedValues.Add([ordered]@{ Path = $mini; Name = 'globalEnable'; Value = 'false' }) }
        if ($null -eq $beforeDefault) { $null = $script:CreatedValues.Add([ordered]@{ Path = $mini; Name = 'defaultEnable'; Value = 'false' }) }
        $changed++
    }
    $miniState = Get-RegValueText -Path $mini -Name 'globalEnable'
    $fileTransfer = Get-RegValueText -Path $wpscloud -Name 'fileTransfer'
    if ($changed -gt 0) {
        Write-Log ('已关闭 ' + $changed + ' 个 WPS 网盘/传输助手开关。') 'OK'
        Add-Result -Category 'WPS网盘/传输助手' -Item '云盘同步与传输助手开关' -Status '已关闭' -Detail ('共 ' + $changed + ' 项')
    } else {
        Add-Result -Category 'WPS网盘/传输助手' -Item '云盘同步与传输助手开关' -Status '已是关闭状态' -Detail ''
    }
    if ($fileTransfer -eq 'false' -and $miniState -eq 'false') {
        Write-Log '文件传输助手：总开关已关闭（globalEnable=false）。' 'OK'
        Add-Result -Category 'WPS网盘/传输助手' -Item '文件传输助手（传输助手/跨设备）' -Status '已关闭' -Detail 'globalEnable=false'
    }
}

function Invoke-AutostartCleanup {
    if (-not $SkipScheduledTasks) {
        Write-Log '禁用 WPS 更新计划任务...' 'INFO'
        $tasks = Get-WpsScheduledTasks
        if ($tasks.Count -eq 0) {
            Write-Log '未发现 WPS 更新计划任务。' 'INFO'
            Add-Result -Category '自启动' -Item 'WPS 更新计划任务' -Status '未发现' -Detail ''
        } else {
            foreach ($t in $tasks) {
                if ([string]$t.State -eq 'Disabled') {
                    Add-Result -Category '自启动' -Item ('计划任务：' + $t.TaskName) -Status '已是禁用' -Detail ''
                    continue
                }
                try {
                    Disable-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop | Out-Null
                    Write-Log ('已禁用计划任务：' + $t.TaskName) 'OK'
                    Add-Result -Category '自启动' -Item ('计划任务：' + $t.TaskName) -Status '已禁用' -Detail ''
                } catch {
                    Write-Log ('禁用计划任务失败：' + $t.TaskName + '（' + $_.Exception.Message + '）') 'WARN'
                    Add-Result -Category '自启动' -Item ('计划任务：' + $t.TaskName) -Status '失败' -Detail $_.Exception.Message
                }
            }
        }
    }
    if (-not $SkipService) {
        Write-Log '停用并禁用 WPS 云服务（wpscloudsvr）...' 'INFO'
        $services = Get-WpsServices
        if ($services.Count -eq 0) {
            Write-Log '未发现 WPS 相关服务。' 'INFO'
            Add-Result -Category '自启动' -Item 'WPS 云服务' -Status '未发现' -Detail ''
        } else {
            foreach ($s in $services) {
                try {
                    if ($s.Status -eq 'Running') { Stop-Service -Name $s.Name -Force -ErrorAction Stop }
                    Set-Service -Name $s.Name -StartupType Disabled -ErrorAction Stop
                    Write-Log ('已停用并禁用服务：' + $s.Name) 'OK'
                    Add-Result -Category '自启动' -Item ('服务：' + $s.Name) -Status '已禁用' -Detail '启动类型=Disabled'
                } catch {
                    Write-Log ('禁用服务失败：' + $s.Name + '（' + $_.Exception.Message + '）') 'WARN'
                    Add-Result -Category '自启动' -Item ('服务：' + $s.Name) -Status '失败' -Detail $_.Exception.Message
                }
            }
        }
    }
    Write-Log '清理注册表 Run/RunOnce 与启动文件夹中的 WPS 项...' 'INFO'
    $runs = Get-AutostartRunEntries
    if ($runs.Count -eq 0) {
        Add-Result -Category '自启动' -Item '注册表 Run/RunOnce' -Status '无相关项' -Detail ''
    } else {
        foreach ($r in $runs) {
            try {
                Remove-ItemProperty -Path $r.Path -Name $r.Name -ErrorAction Stop
                Write-Log ('已移除启动项：' + $r.RegPath + ' -> ' + $r.Name) 'OK'
                Add-Result -Category '自启动' -Item ('启动项：' + $r.Name) -Status '已移除' -Detail $r.RegPath
            } catch {
                Add-Result -Category '自启动' -Item ('启动项：' + $r.Name) -Status '失败' -Detail $_.Exception.Message
            }
        }
    }
    $links = Get-WpsStartupShortcuts
    if ($links.Count -eq 0) {
        Add-Result -Category '自启动' -Item '启动文件夹快捷方式' -Status '无相关项' -Detail ''
    } else {
        Add-Result -Category '自启动' -Item '启动文件夹快捷方式' -Status '已移到备份' -Detail ('共 ' + $links.Count + ' 个')
    }
}

function Invoke-DocerTabHardening {
    Write-Log '关闭启动时自动打开的「稻壳AI模板」页...' 'INFO'
    $path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kdocerpage'
    $before = Get-RegValueText -Path $path -Name 'docerTabShow'
    $ok = Set-RegValue -Path $path -Name 'docerTabShow' -Value '0' -Type DWord
    if ($ok) {
        if ($null -eq $before) {
            $null = $script:CreatedValues.Add([ordered]@{ Path = $path; Name = 'docerTabShow'; Value = '0' })
        }
        if ($before -eq '0') {
            Write-Log '「稻壳AI模板」启动页已是关闭状态（docerTabShow=0）。' 'OK'
            Add-Result -Category '启动弹窗/推荐' -Item '稻壳AI模板启动页' -Status '已是关闭状态' -Detail 'kdocerpage\docerTabShow=0'
        } else {
            Write-Log '已关闭「稻壳AI模板」启动页（WPS 下次启动不再自动打开）。' 'OK'
            Add-Result -Category '启动弹窗/推荐' -Item '稻壳AI模板启动页' -Status '已关闭' -Detail 'kdocerpage\docerTabShow=0'
        }
    } else {
        Write-Log '写入「稻壳AI模板」启动页开关失败（可能被安全软件拦截）。' 'WARN'
        Add-Result -Category '启动弹窗/推荐' -Item '稻壳AI模板启动页' -Status '失败' -Detail '写入注册表失败'
    }
}

function Invoke-StartPagePanelHardening {
    Write-Log '关闭 WPS 启动首页右侧的「精选推荐 / 功能推荐」面板...' 'INFO'
    $path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kstartpage\rightAppInfo'
    $before = Get-RegValueText -Path $path -Name 'rightPanelWidgetVisibleGlobal'
    $ok = Set-RegValue -Path $path -Name 'rightPanelWidgetVisibleGlobal' -Value 'false' -Type String
    if ($ok) {
        if ($null -eq $before) {
            $null = $script:CreatedValues.Add([ordered]@{ Path = $path; Name = 'rightPanelWidgetVisibleGlobal'; Value = 'false' })
        }
        if ($before -eq 'false') {
            Write-Log '首页右侧面板已是关闭状态（rightPanelWidgetVisibleGlobal=false）。' 'OK'
            Add-Result -Category '启动弹窗/推荐' -Item '首页右侧面板（精选推荐/功能推荐）' -Status '已是关闭状态' -Detail 'kstartpage\rightAppInfo\rightPanelWidgetVisibleGlobal=false'
        } else {
            Write-Log '已关闭首页右侧面板（WPS 下次启动不再显示）。' 'OK'
            Add-Result -Category '启动弹窗/推荐' -Item '首页右侧面板（精选推荐/功能推荐）' -Status '已关闭' -Detail 'kstartpage\rightAppInfo\rightPanelWidgetVisibleGlobal=false'
        }
    } else {
        Write-Log '写入首页右侧面板开关失败（可能被安全软件拦截）。' 'WARN'
        Add-Result -Category '启动弹窗/推荐' -Item '首页右侧面板（精选推荐/功能推荐）' -Status '失败' -Detail '写入注册表失败'
    }
}

function Save-CreatedValues {
    if (-not $script:BackupDir) { return }
    $file = Join-Path $script:BackupDir 'created-values.json'
    try {
        $script:CreatedValues | ConvertTo-Json -Depth 4 | Set-Content -Path $file -Encoding UTF8 -ErrorAction Stop
    } catch { }
}

function Show-Summary {
    Write-Log '汇总结果' 'STEP'
    if ($script:Summary.Count -eq 0) {
        if ($Mode -eq 'Scan') { Write-Log '只读检测完成，未做任何修改。' 'OK' }
        elseif ($Mode -eq 'Revert') { Write-Log '恢复流程未产生可汇总的条目。' 'WARN' }
        else { Write-Log '本次没有执行任何改变。' 'WARN' }
    } else {
        foreach ($row in $script:Summary) {
            Write-CliLine ('  [' + $row.Category + '] ' + $row.Item + '：' + $row.Status + '  ' + $row.Detail)
        }
    }
    Write-CliLine ''
    if ($script:BackupDir) { Write-Log ('备份目录：' + $script:BackupDir) 'INFO' }
    if ($script:LogFile) { Write-Log ('运行日志：' + $script:LogFile) 'INFO' }
    Write-Log '提示：系统右键菜单项、资源管理器集成项的改动，重启资源管理器或重启电脑后完全生效。' 'INFO'
    Write-Log '如需运行后立即生效（自动重启资源管理器），可选择菜单前加参数：-RestartExplorer。' 'INFO'
    Write-Log '如需撤销本次所有改动：重新双击本脚本，选择 3（一键恢复）；或运行 WPS一键优化.bat -Mode Revert。' 'INFO'
    Write-Log 'WPS 个人版的云文档、WPS AI、在线模板、PDF 会员等功能需要登录/会员，脚本不做也不应做破解。' 'INFO'
    Write-Log '若 WPS 升级后设置被恢复，重新运行本脚本即可。' 'INFO'
}

# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
function Invoke-Apply {
    Write-Log ('开始执行 WPS 一键优化（脚本版本 ' + $script:ScriptVersion + '）') 'STEP'
    $script:WpsInfo = Get-WpsInfo -ExplicitPath $WpsPath
    if (-not $script:WpsInfo.Found) {
        Write-Log '未检测到 WPS Office。请确认已安装 WPS 后再运行。' 'ERR'
        if ($script:WpsInfo.Searched) { Write-Log ('已尝试的查找方式：' + $script:WpsInfo.Searched) 'WARN' }
        Write-Log '如 WPS 装在非常规目录，可手动指定路径重试：WPS一键优化.bat -Mode Apply -WpsPath "D:\Apps\WPS Office\12.1.0.29690"' 'WARN'
        Add-Result -Category '环境' -Item '检测 WPS' -Status '失败' -Detail '未找到 ksomisc.exe / 安装信息'
        return
    }
    Write-Log ('WPS 目录：' + $script:WpsInfo.InstallRoot + '（版本 ' + $script:WpsInfo.Version + '，发现方式：' + $script:WpsInfo.Method + '）') 'INFO'
    if (Test-WpsRunning) {
        Write-Log '检测到 WPS 正在运行：建议先保存文档并关闭 WPS，再重新运行以获取最稳妥的效果。' 'WARN'
    }
    $null = New-OptimizeBackup

    Write-Log '第一步：关闭“功能定制”中的 13 个推荐 / 系统右键菜单项' 'STEP'
    $uiResult = $null
    if ($SkipUiAutomation) {
        Write-Log '已按参数跳过 UI 自动化。' 'WARN'
        Add-Result -Category '功能定制' -Item '13 项设置' -Status '已跳过' -Detail '-SkipUiAutomation'
    } else {
        $uiResult = Invoke-WpsConfigCleanup -KsomiscPath $script:WpsInfo.Ksomisc
        if ($uiResult -and $uiResult.Supported -and $uiResult.Items.Count -gt 0) {
            foreach ($it in $uiResult.Items) {
                $status = '未知'
                if ($it.Action -eq 'already-on') { $status = '已是关闭状态' }
                elseif ($it.Action -eq 'toggled') { $status = '已勾选(关闭)' }
                elseif ($it.Action -eq 'not-found') { $status = '未找到(版本差异)' }
                elseif ($it.Action -eq 'toggle-failed') { $status = '操作失败' }
                Add-Result -Category ([string]$it.Group) -Item ([string]$it.Name) -Status $status -Detail ('点击前=' + [string]$it.Before)
            }
            if ($uiResult.Applied) {
                $verifiedOn = 0
                if ($uiResult.Verified) {
                    foreach ($t in (Get-TargetDefinitions)) {
                        if ($uiResult.Verified[$t.Name] -eq 'On') { $verifiedOn++ }
                    }
                }
                if ($verifiedOn -ge 12) {
                    Write-Log ('官方配置工具已应用设置，复核结果：' + $verifiedOn + '/13 项处于“已勾选(关闭)”。') 'OK'
                } else {
                    Write-Log ('官方配置工具已应用设置，但复核仅有 ' + $verifiedOn + '/13 项生效。') 'WARN'
                    Add-Result -Category '功能定制' -Item '复核' -Status '部分未生效' -Detail ($verifiedOn.ToString() + '/13')
                }
            } else {
                Write-Log ('配置工具执行失败：' + $uiResult.Message) 'WARN'
                Add-Result -Category '功能定制' -Item '应用设置' -Status '失败' -Detail ([string]$uiResult.Message)
            }
        }
        $notFoundCount = 0
        if ($uiResult -and $uiResult.Items) {
            $notFoundCount = @($uiResult.Items | Where-Object { $_.Action -eq 'not-found' }).Count
        }
        if ((-not $uiResult) -or (-not $uiResult.Supported) -or ($notFoundCount -gt 0)) {
            Write-Log 'UI 自动化不可用或部分项未找到，尝试写入已知明文开关作为降级方案...' 'WARN'
            $fb = Invoke-MenuItemsRegistryFallback
            Write-Log ('降级方案：写入 ' + $fb.Ops + ' 个明文开关，删除 ' + $fb.RemovedKeys + ' 个“新建”菜单注册项。') 'INFO'
            Add-Result -Category '功能定制' -Item '降级方案（明文开关）' -Status '已执行' -Detail ('写入 ' + $fb.Ops + ' 项；删除注册项 ' + $fb.RemovedKeys + ' 个。加密开关项需在官方配置工具中手动勾选。')
        }
    }

    Write-Log '第二步：关闭启动时自动打开的「稻壳AI模板」页 / 首页右侧面板' 'STEP'
    Invoke-DocerTabHardening
    Invoke-StartPagePanelHardening

    if (-not $SkipLogin) {
        Write-Log '第三步：免强制登录 / 离线使用' 'STEP'
        Invoke-LoginHardening
    }
    if (-not $SkipCloud) {
        Write-Log '第四步：关闭 WPS 网盘 / 文件传输助手常驻开关' 'STEP'
        Invoke-CloudHardening
    }

    Write-Log '第五步：关闭 WPS 自启动项（计划任务 / 云服务 / Run / 启动文件夹）' 'STEP'
    Invoke-AutostartCleanup
    Save-CreatedValues

    if (-not $SkipExplorerEntries) {
        Write-Log '第六步：清理资源管理器中的 WPS 入口（WPS 云盘图标、云盘右键扩展）' 'STEP'
        Invoke-ExplorerEntryCleanup
    } else {
        Write-Log '已按参数跳过资源管理器入口清理。' 'WARN'
        Add-Result -Category '资源管理器入口' -Item 'WPS 云盘图标 / 云盘右键扩展' -Status '已跳过' -Detail '-SkipExplorerEntries'
    }
    Restart-ExplorerSafe

    Write-Log '第七步：复核结果' 'STEP'
    if (-not $SkipLogin) {
        $flogin = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\officespace\flogin'
        $v1 = [string](Get-RegValueText -Path $flogin -Name 'enableForceLoginForFirstInstallDevice')
        $v2 = [string](Get-RegValueText -Path $flogin -Name 'enableForceLoginForHasInstallDevice')
        Write-Log ('  强制登录开关：FirstInstallDevice=' + $v1 + ' / HasInstallDevice=' + $v2) 'INFO'
    }
    $leftTasks = @(Get-WpsScheduledTasks | Where-Object { [string]$_.State -ne 'Disabled' })
    $leftServices = @(Get-WpsServices | Where-Object { [string]$_.StartType -ne 'Disabled' })
    $leftRuns = @(Get-AutostartRunEntries)
    Write-Log ('  剩余未禁用：计划任务 ' + $leftTasks.Count + ' 个，服务 ' + $leftServices.Count + ' 个，Run/RunOnce ' + $leftRuns.Count + ' 个') $(if (($leftTasks.Count -eq 0) -and ($leftServices.Count -eq 0) -and ($leftRuns.Count -eq 0)) { 'OK' } else { 'WARN' })
}

function Invoke-Revert {
    Write-Log '开始从备份恢复 WPS 设置...' 'STEP'
    if (-not (Test-Path $script:BackupRoot)) {
        Write-Log '未找到备份目录，无法恢复。' 'ERR'
        return
    }
    $allDirs = @(Get-ChildItem -Path $script:BackupRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName 'manifest.json') } |
        Sort-Object Name)
    if ($allDirs.Count -eq 0) {
        Write-Log '备份目录为空或缺少 manifest.json，无法恢复。' 'ERR'
        return
    }
    # 恢复点：使用“最早的一次备份”（即优化前的原始状态），避免多次 Apply 后恢复成上次优化状态
    $restorePoint = $allDirs[0]
    $manifest = Get-Content -Path (Join-Path $restorePoint.FullName 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-Log ('恢复点（最早备份）：' + $restorePoint.FullName + '（创建时间 ' + $manifest.CreatedAt + '）') 'INFO'
    if ($allDirs.Count -gt 1) {
        Write-Log ('共发现 ' + $allDirs.Count + ' 个备份，本次使用最早的一个作为“原始状态”恢复点；资源管理器入口等从后续备份补齐。') 'INFO'
    }
    if (Test-WpsRunning) {
        Write-Log '检测到 WPS 正在运行：建议先保存文档并关闭 WPS，再重新执行恢复。' 'WARN'
    }
    foreach ($regFile in $manifest.RegistryBackups) {
        $full = Join-Path $restorePoint.FullName $regFile
        if (-not (Test-Path $full)) { continue }
        $null = cmd.exe /c ('reg import "' + $full + '"') 2>&1
        Write-Log ('已导入注册表备份：' + $regFile) 'OK'
    }
    # 资源管理器 WPS 入口：从包含 explorer-entries.json 的最新备份恢复
    $explorerDir = $null
    for ($i = $allDirs.Count - 1; $i -ge 0; $i--) {
        if (Test-Path (Join-Path $allDirs[$i].FullName 'explorer-entries.json')) { $explorerDir = $allDirs[$i]; break }
    }
    if ($explorerDir) {
        try {
            $explorerRecords = Get-Content -Path (Join-Path $explorerDir.FullName 'explorer-entries.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($rec in @($explorerRecords)) {
                if (-not $rec) { continue }
                $regFull = Join-Path $explorerDir.FullName $rec.RegFile
                if (Test-Path $regFull) {
                    $null = cmd.exe /c ('reg import "' + $regFull + '"') 2>&1
                    Write-Log ('已恢复资源管理器 WPS 入口：' + $rec.Label) 'OK'
                }
            }
        } catch {
            Write-Log ('恢复资源管理器 WPS 入口失败：' + $_.Exception.Message) 'WARN'
        }
    } else {
        Write-Log '备份中没有资源管理器入口记录（首次运行的是旧版本脚本时会出现），跳过入口恢复。' 'WARN'
    }
    # 删除脚本新建的注册表值：优先用最新 created-values.json，并用恢复点 .reg 兜底核对
    $createdDir = $null
    for ($i = $allDirs.Count - 1; $i -ge 0; $i--) {
        if (Test-Path (Join-Path $allDirs[$i].FullName 'created-values.json')) { $createdDir = $allDirs[$i]; break }
    }
    if ($createdDir) {
        try {
            $created = Get-Content -Path (Join-Path $createdDir.FullName 'created-values.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($cv in @($created)) {
                if (-not $cv) { continue }
                $cur = Get-RegValueText -Path $cv.Path -Name $cv.Name
                if (($null -ne $cur) -and ($cur -eq [string]$cv.Value)) {
                    try {
                        Remove-ItemProperty -Path $cv.Path -Name $cv.Name -ErrorAction Stop
                        Write-Log ('已删除脚本新建的注册表值：' + $cv.Name) 'OK'
                    } catch { }
                }
            }
        } catch {
            Write-Log ('清理新建注册表值记录失败：' + $_.Exception.Message) 'WARN'
        }
    }
    $restoreRegFile = Join-Path $restorePoint.FullName 'HKCU-Software-Kingsoft.reg'
    if (Test-Path $restoreRegFile) {
        $regText = ''
        try { $regText = Get-Content -Path $restoreRegFile -Raw -Encoding Unicode } catch { }
        if ($regText) {
            $candidates = @(
                @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\officespace\flogin'; Name = 'enableForceLoginForFirstInstallDevice' },
                @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\officespace\flogin'; Name = 'enableForceLoginForHasInstallDevice' },
                @{ Path = 'HKCU:\Software\Kingsoft\wpscloud\miniTransferAssistant'; Name = 'globalEnable' },
                @{ Path = 'HKCU:\Software\Kingsoft\wpscloud\miniTransferAssistant'; Name = 'defaultEnable' },
                @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kdocerpage'; Name = 'docerTabShow' },
                @{ Path = 'HKCU:\Software\Kingsoft\Office\6.0\plugins\kstartpage\rightAppInfo'; Name = 'rightPanelWidgetVisibleGlobal' }
            )
            foreach ($cand in $candidates) {
                if ($regText.Contains('"' + $cand.Name + '"')) { continue }
                $cur = Get-RegValueText -Path $cand.Path -Name $cand.Name
                if ($null -ne $cur) {
                    try {
                        Remove-ItemProperty -Path $cand.Path -Name $cand.Name -ErrorAction Stop
                        Write-Log ('已删除脚本新建的注册表值：' + $cand.Name) 'OK'
                    } catch { }
                }
            }
        }
    }
    foreach ($v in $manifest.RunValues) {
        try {
            if (-not (Test-Path $v.Path)) { New-Item -Path $v.Path -Force | Out-Null }
            New-ItemProperty -Path $v.Path -Name $v.Name -Value $v.Value -PropertyType String -Force -ErrorAction Stop | Out-Null
            Write-Log ('已恢复启动项：' + $v.Name) 'OK'
        } catch {
            Write-Log ('恢复启动项失败：' + $v.Name) 'WARN'
        }
    }
    foreach ($l in $manifest.StartupLinks) {
        if (Test-Path $l.Backup) {
            try {
                Move-Item -Path $l.Backup -Destination $l.Original -Force -ErrorAction Stop
                Write-Log ('已恢复启动文件夹快捷方式：' + $l.Original) 'OK'
            } catch {
                Write-Log ('恢复启动快捷方式失败：' + $l.Original) 'WARN'
            }
        }
    }
    foreach ($t in $manifest.Tasks) {
        $xmlPath = Join-Path $restorePoint.FullName $t.XmlFile
        try {
            $existing = Get-ScheduledTask -TaskName $t.Name -TaskPath $t.TaskPath -ErrorAction SilentlyContinue
            $wasDisabled = ([string]$t.State -eq 'Disabled')
            if ($existing) {
                if ($wasDisabled) {
                    Disable-ScheduledTask -TaskName $t.Name -TaskPath $t.TaskPath -ErrorAction Stop | Out-Null
                } else {
                    Enable-ScheduledTask -TaskName $t.Name -TaskPath $t.TaskPath -ErrorAction Stop | Out-Null
                }
                Write-Log ('已恢复计划任务状态：' + $t.Name + '（原始状态=' + $t.State + '）') 'OK'
            } elseif (Test-Path $xmlPath) {
                $xml = Get-Content -Path $xmlPath -Raw -Encoding UTF8
                Register-ScheduledTask -Xml $xml -TaskName $t.Name -TaskPath $t.TaskPath -Force -ErrorAction Stop | Out-Null
                if ($wasDisabled) { Disable-ScheduledTask -TaskName $t.Name -TaskPath $t.TaskPath -ErrorAction SilentlyContinue | Out-Null }
                Write-Log ('已重建计划任务：' + $t.Name + '（原始状态=' + $t.State + '）') 'OK'
            }
        } catch {
            Write-Log ('恢复计划任务失败：' + $t.Name + '（' + $_.Exception.Message + '）') 'WARN'
        }
    }
    if ($manifest.Service) {
        try {
            $svc = Get-Service -Name $manifest.Service.Name -ErrorAction SilentlyContinue
            if ($svc) {
                Set-Service -Name $manifest.Service.Name -StartupType $manifest.Service.StartType -ErrorAction Stop
                if ($manifest.Service.Status -eq 'Running') { Start-Service -Name $manifest.Service.Name -ErrorAction SilentlyContinue }
                Write-Log ('已恢复服务：' + $manifest.Service.Name + '（启动类型=' + $manifest.Service.StartType + '）') 'OK'
            }
        } catch {
            Write-Log ('恢复服务失败：' + $_.Exception.Message) 'WARN'
        }
    }
    Restart-ExplorerSafe
    Write-Log '恢复完成。' 'OK'
    Write-Log '功能定制 13 项的开关保存在注册表 HKCU\Software\Kingsoft 中，已随注册表备份一并恢复。' 'INFO'
    Write-Log '提示：右键菜单等资源管理器集成项的恢复，重启资源管理器或重启电脑后完全生效。' 'INFO'
}

# ----------------------------------------------------------------------------
# 入口
# ----------------------------------------------------------------------------
function Show-MainMenu {
    while ($true) {
        Write-CliLine ''
        Write-CliLine -Text '请选择要执行的操作：' -Color 'Cyan'
        Write-CliLine '  1. 一键优化（关闭推荐/稻壳AI模板弹窗/自启动/WPS网盘/资源管理器入口等）'
        Write-CliLine '  2. 只读检测（查看当前状态，不改动任何东西）'
        Write-CliLine '  3. 一键恢复（从最早备份恢复原始状态）'
        Write-CliLine '  0. 退出'
        Write-CliLine ''
        $choice = $null
        try { $choice = Read-Host '请输入序号后回车' } catch { $choice = $null }
        $choice = ([string]$choice).Trim()
        switch ($choice) {
            '1' { return 'Apply' }
            '2' { return 'Scan' }
            '3' { return 'Revert' }
            '0' { Write-CliLine -Text '已退出。' -Color 'DarkGray'; return 'Exit' }
            '' {
                Write-CliLine -Text '没有收到输入（可能不是交互式控制台）。请带参数运行，例如：WPS一键优化.bat -Mode Apply / -Mode Scan / -Mode Revert' -Color 'Yellow'
                return 'Exit'
            }
            default {
                Write-CliLine -Text '序号无效，请输入 0~3 的数字。' -Color 'Yellow'
                continue
            }
        }
    }
}

function Main {
    Initialize-Environment
    Write-CliLine ''
    Write-CliLine -Text '============================================================' -Color 'Cyan'
    Write-CliLine -Text ('  WPS Office 一键优化（单文件版）  v' + $script:ScriptVersion) -Color 'Cyan'
    Write-CliLine -Text '  优化 / 检测 / 恢复  三合一，一个文件搞定' -Color 'Cyan'
    Write-CliLine -Text '============================================================' -Color 'Cyan'
    Write-Log ('脚本目录：' + $script:ScriptDir) 'INFO'

    if ($Mode -eq 'Menu') {
        while ($true) {
            $picked = Show-MainMenu
            if ($picked -eq 'Exit') { return }
            $Mode = $picked
            Write-Log ('运行模式：' + $Mode) 'INFO'

            if (($Mode -eq 'Apply' -or $Mode -eq 'Revert') -and (-not $Elevated) -and (-not (Test-IsAdmin))) {
                $elevatedOk = Invoke-SelfElevate
                if ($elevatedOk) {
                    $script:SkipFinalPause = $true
                    return
                }
                Write-Log '继续以普通权限运行（部分项可能跳过/失败）。' 'WARN'
            }

            $script:Summary.Clear()
            if ($Mode -eq 'Scan') { Show-CurrentState; Show-Summary }
            elseif ($Mode -eq 'Revert') { Invoke-Revert; Show-Summary }
            else { Invoke-Apply; Show-Summary }
            Write-CliLine ''
            Write-CliLine -Text '按回车键返回主菜单（或直接关闭窗口退出）...' -Color 'DarkGray'
            try { Read-Host | Out-Null } catch { return }
            $Mode = 'Menu'
        }
    }

    Write-Log ('运行模式：' + $Mode) 'INFO'
    if (($Mode -eq 'Apply' -or $Mode -eq 'Revert') -and (-not $Elevated) -and (-not (Test-IsAdmin))) {
        $elevatedOk = Invoke-SelfElevate
        if ($elevatedOk) {
            $script:SkipFinalPause = $true
            return
        }
        Write-Log '继续以普通权限运行（部分项可能跳过/失败）。' 'WARN'
    }

    $script:Summary.Clear()
    if ($Mode -eq 'Scan') { Show-CurrentState; Show-Summary }
    elseif ($Mode -eq 'Revert') { Invoke-Revert; Show-Summary }
    else { Invoke-Apply; Show-Summary }
}

try {
    Main
} catch {
    Write-Log ('脚本执行异常：' + $_.Exception.Message) 'ERR'
    Write-Log ([string]$_.ScriptStackTrace) 'ERR'
    Save-LogBuffer
}
Save-LogBuffer
if ($Pause -and (-not $script:SkipFinalPause)) {
    Write-CliLine ''
    Write-CliLine -Text '按回车键退出...' -Color 'DarkGray'
    try { Read-Host | Out-Null } catch { }
}
