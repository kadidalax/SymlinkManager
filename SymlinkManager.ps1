#requires -Version 5.0
<#
.SYNOPSIS
    Windows 软链接管理器（修复版）
.DESCRIPTION
    Source（源）   = 已经存在的真实文件/文件夹（数据实际所在位置）
    Target（目标） = 要新建的链接路径（不能已存在，最后一段是链接名称）

    修复内容：
    - 明确 Source / Target 的含义与方向，避免和 mklink 的参数顺序混淆
    - 不再静默使用 -Force 覆盖已有路径
    - 增加路径、名称、类型、磁盘格式、网络位置、同步盘校验
    - 删除时只删除重解析点，绝不递归删除目标内容
    - 支持目录符号链接 / Junction / 文件符号链接
    - 支持双击记录打开链接、右键菜单、状态校验、配置文件容错
.PARAMETER SelfTest
    运行内置自检（不显示界面），用于验证创建/删除逻辑。
#>
[CmdletBinding()]
param(
    [switch]$SelfTest
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

try { $script:UiFont = New-Object System.Drawing.Font('Microsoft YaHei UI', 9) }
catch { $script:UiFont = New-Object System.Drawing.Font('Segoe UI', 9) }
$script:BoldFont = New-Object System.Drawing.Font($script:UiFont, [System.Drawing.FontStyle]::Bold)
$script:HelpTitleFont = New-Object System.Drawing.Font($script:UiFont.FontFamily, 14, [System.Drawing.FontStyle]::Bold)
$script:HelpHeadingFont = New-Object System.Drawing.Font($script:UiFont.FontFamily, 11, [System.Drawing.FontStyle]::Bold)

$script:ConfigPath     = $null
$script:LogPath        = $null
$script:DataDir        = $null
$script:IsAdmin        = $false
$script:StartupWarning = $null
$script:MainForm       = $null
$script:DataGridView   = $null
$script:StatusLabel    = $null
$script:SelfTestFailed = 0

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

function Test-Admin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

function Initialize-DataPaths {
    $candidates = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        [void]$candidates.Add($PSScriptRoot)
    }
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        [void]$candidates.Add((Join-Path $env:LOCALAPPDATA 'MLinkTools'))
    }
    if (-not [string]::IsNullOrWhiteSpace($env:TEMP)) {
        [void]$candidates.Add((Join-Path $env:TEMP 'MLinkTools'))
    }

    foreach ($dir in $candidates) {
        try {
            if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            $probe = Join-Path $dir ('.mlink_write_test_' + [guid]::NewGuid().ToString('N'))
            [IO.File]::WriteAllText($probe, 'ok')
            Remove-Item -LiteralPath $probe -Force
            $script:DataDir = $dir
            break
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($script:DataDir)) {
        $script:DataDir = $env:TEMP
    }
    $script:ConfigPath = Join-Path $script:DataDir 'symlinks.json'
    $script:LogPath    = Join-Path $script:DataDir 'symlink-manager.log'
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    if ([string]::IsNullOrWhiteSpace($script:LogPath)) { return }
    try {
        $line = '{0} [{1}] {2}{3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message, [Environment]::NewLine
        [IO.File]::AppendAllText($script:LogPath, $line, (New-Object System.Text.UTF8Encoding($false)))
    } catch { }
}

function Get-FriendlyError {
    param([System.Exception]$Exception)
    $msg = $Exception.Message
    $inner = $Exception.InnerException
    while ($null -ne $inner) {
        if (-not [string]::IsNullOrWhiteSpace($inner.Message)) { $msg = $inner.Message }
        $inner = $inner.InnerException
    }

    if ($msg -match '管理员|Administrator|privilege|elevat|提升|拒绝访问|Access is denied|Access denied') {
        return ("权限不足：创建符号链接通常需要管理员权限（或 Windows 开发者模式）。`n" +
                "也可能是 Target 所在位置不允许写入。`n`n" +
                "文件夹可以改用『目录联接 (Junction)』；文件链接必须提升权限。`n" +
                "或点击界面下方的『以管理员身份重启』。`n`n" +
                "原始错误：$msg")
    }
    if ($msg -match '已存在|already exists') {
        return ("Target 路径已存在。`n" +
                "请换一个链接名称，或先删除旧的链接。`n`n" +
                "原始错误：$msg")
    }
    if ($msg -match '不支持|not supported|重解析|reparse') {
        return ("该位置不支持软链接/目录联接。`n" +
                "请把链接放在本地 NTFS 磁盘上，不要放在网络共享或同步盘。`n`n" +
                "原始错误：$msg")
    }
    return "操作失败：`n$msg"
}

function Show-ErrorMessage {
    param(
        [string]$Title,
        [string]$Message,
        $Owner = $null
    )
    Write-Log "$Title : $Message" 'ERROR'
    if ($null -ne $Owner) {
        [void][System.Windows.Forms.MessageBox]::Show($Owner, $Message, $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        [void][System.Windows.Forms.MessageBox]::Show($Message, $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# ---------------------------------------------------------------------------
# 路径与名称校验
# ---------------------------------------------------------------------------

function Get-InputFullPath {
    param(
        [string]$Path,
        [string]$FieldName = '路径'
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "$FieldName 不能为空。" }
    $p = $Path.Trim()
    if ($p.Length -ge 2 -and $p.StartsWith('"') -and $p.EndsWith('"')) {
        $p = $p.Substring(1, $p.Length - 2).Trim()
    }
    $p = [Environment]::ExpandEnvironmentVariables($p)
    if ($p -notmatch '^(?:[A-Za-z]:[\\/]|[/\\]{2})') {
        throw "$FieldName 必须是完整路径（例如 C:\Folder 或 \\server\share）。"
    }
    try { $p = [IO.Path]::GetFullPath($p) }
    catch { throw "$FieldName 格式无效：$Path" }
    return $p
}

function Get-ExistingItem {
    param(
        [string]$Path,
        [string]$FieldName = '路径'
    )
    $p = Get-InputFullPath -Path $Path -FieldName $FieldName
    try {
        return (Get-Item -LiteralPath $p -Force -ErrorAction Stop)
    } catch {
        throw "$FieldName 不存在或无法访问：`n$p"
    }
}

function Remove-TrailingSeparator {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    $root = [IO.Path]::GetPathRoot($Path)
    if (-not [string]::IsNullOrWhiteSpace($root) -and $Path.Length -le $root.Length) { return $Path }
    return $Path.TrimEnd('\', '/')
}

function Assert-ValidLinkName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { throw '链接名称不能为空。' }
    if ($Name -ne $Name.Trim()) { throw '链接名称前后不能有空格。' }
    if ($Name -match '[\\/:*?"<>|]') { throw '链接名称不能包含这些字符：\ / : * ? " < > |' }
    if ($Name -eq '.' -or $Name -eq '..') { throw '链接名称不能是 . 或 ..' }
    if ($Name.EndsWith('.')) { throw '链接名称不能以句点结尾。' }

    $base = ($Name -split '\.')[0].ToUpperInvariant()
    $reserved = @('CON','PRN','AUX','NUL',
                  'COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
                  'LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')
    if ($reserved -contains $base) { throw "链接名称不能使用系统保留名：$base" }
}

function Test-NetworkPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ($Path.StartsWith('\\')) { return $true }
    try {
        $root = [IO.Path]::GetPathRoot($Path)
        if ([string]::IsNullOrWhiteSpace($root)) { return $false }
        $drive = New-Object System.IO.DriveInfo($root)
        return ($drive.DriveType -eq [IO.DriveType]::Network)
    } catch { return $false }
}

function Get-CloudProvider {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ($env:OneDrive -and $Path.StartsWith($env:OneDrive, [StringComparison]::OrdinalIgnoreCase)) { return 'OneDrive' }
    if ($env:OneDriveCommercial -and $Path.StartsWith($env:OneDriveCommercial, [StringComparison]::OrdinalIgnoreCase)) { return 'OneDrive' }
    if ($env:OneDriveConsumer -and $Path.StartsWith($env:OneDriveConsumer, [StringComparison]::OrdinalIgnoreCase)) { return 'OneDrive' }

    $markers = @(
        @{ Name = 'OneDrive';      Pattern = '\OneDrive' },
        @{ Name = 'Dropbox';       Pattern = '\Dropbox' },
        @{ Name = 'Google Drive';  Pattern = '\Google Drive' },
        @{ Name = 'iCloud';        Pattern = '\iCloudDrive' },
        @{ Name = '坚果云';         Pattern = '\Nutstore' },
        @{ Name = '坚果云';         Pattern = '坚果云' },
        @{ Name = '百度网盘';       Pattern = '\BaiduNetdisk' },
        @{ Name = '百度网盘';       Pattern = '百度网盘' },
        @{ Name = 'MEGA';          Pattern = '\MEGA' }
    )
    foreach ($m in $markers) {
        if ($Path -like ('*' + $m.Pattern + '*')) { return $m.Name }
    }
    return $null
}

function Assert-LinkLocationSupported {
    param([string]$ParentPath)
    if (Test-NetworkPath $ParentPath) {
        throw "Target 不能放在网络共享或映射的网络驱动器上。`n请选择本地 NTFS 文件夹。"
    }

    $format = $null
    try {
        $root = [IO.Path]::GetPathRoot($ParentPath)
        if (-not [string]::IsNullOrWhiteSpace($root)) {
            $drive = New-Object System.IO.DriveInfo($root)
            $format = $drive.DriveFormat
        }
    } catch { }

    if (-not [string]::IsNullOrWhiteSpace($format) -and ($format -notin @('NTFS', 'ReFS'))) {
        throw "Target 所在磁盘是 $format，不支持软链接/目录联接。`n请选择本地 NTFS 磁盘。"
    }
}


# ---------------------------------------------------------------------------
# 链接核心逻辑
# ---------------------------------------------------------------------------

function Test-IsReparsePoint {
    param($Item)
    if ($null -eq $Item) { return $false }
    return [bool]($Item.Attributes -band [IO.FileAttributes]::ReparsePoint)
}

function Get-LinkTargetPath {
    param($Item)
    if ($null -eq $Item) { return $null }
    $targets = @($Item.Target)
    if ($targets.Count -eq 0) { return $null }
    $t = [string]$targets[0]
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }

    # Windows PowerShell 5.1 can expose a UNC target as "UNC\server\share"
    # instead of "\\server\share". Normalize all extended/UNC prefixes before
    # the relative-path fallback, otherwise a UNC target is combined with the
    # link's parent and incorrectly reported as broken.
    if ($t.StartsWith('\\?\UNC\', [StringComparison]::OrdinalIgnoreCase)) {
        $t = '\\' + $t.Substring(8)
    } elseif ($t.StartsWith('\\?\', [StringComparison]::OrdinalIgnoreCase)) {
        $t = $t.Substring(4)
    } elseif ($t.StartsWith('\??\UNC\', [StringComparison]::OrdinalIgnoreCase)) {
        $t = '\\' + $t.Substring(8)
    } elseif ($t.StartsWith('\??\', [StringComparison]::OrdinalIgnoreCase)) {
        $t = $t.Substring(4)
    } elseif ($t.StartsWith('UNC\', [StringComparison]::OrdinalIgnoreCase)) {
        $t = '\\' + $t.Substring(4)
    }

    if (-not [IO.Path]::IsPathRooted($t)) {
        try {
            $parent = Split-Path -Parent $Item.FullName
            $t = [IO.Path]::GetFullPath([IO.Path]::Combine($parent, $t))
        } catch { }
    }
    return $t
}

function Test-TargetChainExists {
    param([string]$Path)
    # Test-Path returns $true for a path that is itself a dangling link, so an
    # outer link to a broken junction/symlink would be reported as healthy.
    # Walk the reparse chain (with a loop limit) instead.
    for ($depth = 0; $depth -lt 16; $depth++) {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        if ($null -eq $item) { return $false }
        if (-not (Test-IsReparsePoint $item)) { return $true }

        $linkType = [string]$item.LinkType
        if ($linkType -notin @('SymbolicLink', 'Junction')) {
            # Cloud placeholders and other non-link reparse points are valid
            # data objects; do not follow them.
            return $true
        }

        $next = Get-LinkTargetPath $item
        if ([string]::IsNullOrWhiteSpace($next)) { return $false }
        $Path = $next
    }
    return $false
}

function Resolve-TargetChain {
    param([string]$Path)
    # Return the final local path of a link chain. Used by New-Link so the
    # loop guard sees the physical location even when the Target parent is a
    # junction/symlink, and so the volume check is done on the real volume.
    for ($depth = 0; $depth -lt 16; $depth++) {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        if ($null -eq $item) { return $Path }
        if (-not (Test-IsReparsePoint $item)) { return $item.FullName }

        $linkType = [string]$item.LinkType
        if ($linkType -notin @('SymbolicLink', 'Junction')) { return $item.FullName }

        $next = Get-LinkTargetPath $item
        if ([string]::IsNullOrWhiteSpace($next)) { return $item.FullName }
        $Path = $next
    }
    return $Path
}

function Get-KindText {
    param($Record)
    switch ([string]$Record.linkKind) {
        'Junction' {
            return '目录联接 (Junction)'
        }
        'SymbolicLink' {
            if ([string]::IsNullOrWhiteSpace([string]$Record.type)) { return '链接（类型未知）' }
            if ([string]$Record.type -eq 'Directory') { return '目录符号链接' }
            return '文件符号链接'
        }
        default {
            if ([string]::IsNullOrWhiteSpace([string]$Record.type)) { return '链接（类型未知）' }
            if ([string]$Record.type -eq 'Directory') { return '目录链接' }
            return '文件链接'
        }
    }
}

function Get-LinkState {
    param($Record)
    $linkPath = [string]$Record.targetPath
    if ([string]::IsNullOrWhiteSpace($linkPath)) {
        return [pscustomobject]@{
            Code = 'Missing'; Text = '记录无效'; Color = [Drawing.Color]::Firebrick
            Target = ''; Message = '记录里没有 Target 路径。'
        }
    }

    $item = Get-Item -LiteralPath $linkPath -Force -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return [pscustomobject]@{
            Code = 'Missing'; Text = '链接不存在'; Color = [Drawing.Color]::Firebrick
            Target = ''; Message = '链接路径已不存在。'
        }
    }
    if (-not (Test-IsReparsePoint $item)) {
        return [pscustomobject]@{
            Code = 'NotLink'; Text = '不是链接'; Color = [Drawing.Color]::Firebrick
            Target = ''; Message = '该路径现在是普通文件/文件夹，不是重解析点。工具不会删除它。'
        }
    }

    $actual = Get-LinkTargetPath $item
    if ([string]::IsNullOrWhiteSpace($actual)) {
        return [pscustomobject]@{
            Code = 'BadLink'; Text = '链接异常'; Color = [Drawing.Color]::DarkOrange
            Target = ''; Message = '无法读取链接目标。'
        }
    }
    if (-not (Test-TargetChainExists $actual)) {
        return [pscustomobject]@{
            Code = 'Broken'; Text = '目标失效'; Color = [Drawing.Color]::DarkOrange
            Target = $actual; Message = "Source（或其链接目标）已不存在：`n$actual"
        }
    }

    $recorded = [string]$Record.sourcePath
    if (-not [string]::IsNullOrWhiteSpace($recorded)) {
        $a = $actual.TrimEnd('\')
        $r = $recorded.TrimEnd('\')
        if (-not [string]::Equals($a, $r, [StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{
                Code = 'Mismatch'; Text = '记录不一致'; Color = [Drawing.Color]::DarkOrange
                Target = $actual
                Message = "链接实际指向：`n$actual`n`n记录中的 Source：`n$recorded`n`n链接仍可使用，但记录已过期。"
            }
        }
    }

    return [pscustomobject]@{
        Code = 'Ok'; Text = '正常'; Color = [Drawing.Color]::ForestGreen
        Target = $actual; Message = ''
    }
}

function New-Link {
    param(
        [string]$SourcePath,
        [string]$LinkPath,
        [ValidateSet('SymbolicLink', 'Junction')][string]$Kind = 'SymbolicLink'
    )

    $srcItem = Get-ExistingItem -Path $SourcePath -FieldName 'Source'
    $srcPath = $srcItem.FullName
    $isDir   = [bool]$srcItem.PSIsContainer

    if ($Kind -eq 'Junction') {
        if (-not $isDir) { throw 'Junction 只能链接文件夹。' }
        if (Test-NetworkPath $srcPath) { throw 'Junction 不能指向网络路径，请改用目录符号链接。' }
    }

    $linkPath = Get-InputFullPath -Path $LinkPath -FieldName 'Target'
    $linkPath = Remove-TrailingSeparator $linkPath
    $linkParent = Split-Path -Parent $linkPath
    if ([string]::IsNullOrWhiteSpace($linkParent)) { throw 'Target 路径无效。' }
    if (-not (Test-Path -LiteralPath $linkParent -PathType Container)) {
        throw "Target 所在的文件夹不存在：`n$linkParent"
    }

    $existing = Get-Item -LiteralPath $linkPath -Force -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        throw "Target 已存在：`n$linkPath`n请换一个链接名称，或先删除旧链接。"
    }

    if ($linkPath.Length -gt 259) {
        throw "Target 路径太长（$($linkPath.Length) 个字符，超过 259）。请缩短名称或换一个父文件夹。"
    }
    $resolvedLinkParent = Resolve-TargetChain $linkParent
    Assert-LinkLocationSupported $resolvedLinkParent

    if ($isDir) {
        $srcFull  = (Resolve-TargetChain $srcPath).TrimEnd('\')
        $linkParentFull = Resolve-TargetChain $linkParent
        $linkFull = [IO.Path]::Combine($linkParentFull, (Split-Path -Leaf $linkPath)).TrimEnd('\')
        if ([string]::Equals($srcFull, $linkFull, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Target 不能和 Source 相同。'
        }
        if ($linkFull.StartsWith($srcFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Target 不能放在 Source 文件夹内部（会造成循环链接）。'
        }
    }

    switch ($Kind) {
        'Junction' {
            New-Item -ItemType Junction -Path $linkPath -Target $srcPath -ErrorAction Stop | Out-Null
        }
        default {
            New-Item -ItemType SymbolicLink -Path $linkPath -Target $srcPath -ErrorAction Stop | Out-Null
        }
    }

    $newItem = Get-Item -LiteralPath $linkPath -Force -ErrorAction Stop
    if (-not (Test-IsReparsePoint $newItem)) {
        throw '创建后校验失败：新路径不是重解析点。'
    }

    return [pscustomobject]@{
        sourcePath = $srcPath
        targetPath = $linkPath
        type       = $(if ($isDir) { 'Directory' } else { 'File' })
        linkKind   = $Kind
        createTime = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

function Remove-LinkSafely {
    param([string]$LinkPath)
    # The config is user-editable; never let a relative path resolve against
    # the process working directory during a delete. Do not expand environment
    # variables here, because a literal folder name can contain "%...%".
    if ([string]::IsNullOrWhiteSpace($LinkPath)) { throw '链接路径为空，已拒绝删除。' }
    if ($LinkPath -notmatch '^(?:[A-Za-z]:[\\/]|[/\\]{2})') {
        throw "链接路径必须是完整路径，已拒绝删除：`n$LinkPath"
    }
    $LinkPath = [IO.Path]::GetFullPath($LinkPath)
    $item = Get-Item -LiteralPath $LinkPath -Force -ErrorAction Stop
    if (-not (Test-IsReparsePoint $item)) {
        throw '安全保护：该路径不是重解析点，已拒绝删除。请手动确认后再处理。'
    }

    $linkType = [string]$item.LinkType
    if ($linkType -notin @('SymbolicLink', 'Junction')) {
        throw "安全保护：该路径不是符号链接/目录联接（LinkType=$linkType），已拒绝删除。可能是云占位符或其他重解析点。"
    }
    if (@($item.Target).Count -eq 0) {
        throw '安全保护：该重解析点没有可读取的链接目标，已拒绝删除。可能是云占位符或其他重解析点。'
    }

    $isDir = [bool]($item.Attributes -band [IO.FileAttributes]::Directory)
    if ($isDir) {
        [IO.Directory]::Delete($LinkPath, $false)
    } else {
        [IO.File]::Delete($LinkPath)
    }

    $left = Get-Item -LiteralPath $LinkPath -Force -ErrorAction SilentlyContinue
    if ($null -ne $left) { throw '删除后链接仍然存在，请检查权限。' }
}

# ---------------------------------------------------------------------------
# 配置文件（兼容旧版 symlinks.json）
# ---------------------------------------------------------------------------

function Read-Records {
    $list = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $script:ConfigPath -PathType Leaf)) { return $list }

    try {
        $raw = [IO.File]::ReadAllText($script:ConfigPath, [Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($raw)) { return $list }
        $cfg = $raw | ConvertFrom-Json -ErrorAction Stop

        $items = @()
        if ($null -ne $cfg) {
            if ($cfg.PSObject.Properties.Name -contains 'symlinks') {
                $items = @($cfg.symlinks)
            } elseif ($cfg -is [array]) {
                $items = @($cfg)
            }
        }

        foreach ($it in $items) {
            if ($null -eq $it) { continue }
            $src = [string]$it.sourcePath
            $lnk = [string]$it.targetPath
            if ([string]::IsNullOrWhiteSpace($lnk)) { continue }

            $kind = [string]$it.linkKind
            $type = [string]$it.type
            $item = Get-Item -LiteralPath $lnk -Force -ErrorAction SilentlyContinue

            if ([string]::IsNullOrWhiteSpace($kind)) {
                if ($null -ne $item -and $item.LinkType -eq 'Junction') { $kind = 'Junction' }
                else { $kind = 'SymbolicLink' }
            }
            if ([string]::IsNullOrWhiteSpace($type) -and $null -ne $item) {
                if ($item.PSIsContainer) { $type = 'Directory' } else { $type = 'File' }
            }

            [void]$list.Add([pscustomobject]@{
                sourcePath = $src
                targetPath = $lnk
                type       = $type
                linkKind   = $kind
                createTime = [string]$it.createTime
            })
        }
    } catch {
        $backup = $script:ConfigPath + '.broken-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
        try {
            Copy-Item -LiteralPath $script:ConfigPath -Destination $backup -Force -ErrorAction Stop
        } catch {
            $backup = '(备份失败)'
        }
        Write-Log "读取配置失败：$($_.Exception.ToString())" 'ERROR'
        $script:StartupWarning = "配置文件读取失败，已备份到：`n$backup`n`n将从空列表开始。"
    }
    return $list
}

function Save-Records {
    param($Records)
    $arr = @()
    foreach ($r in @($Records)) {
        if ($null -eq $r) { continue }
        $arr += [pscustomobject]@{
            sourcePath = [string]$r.sourcePath
            targetPath = [string]$r.targetPath
            type       = [string]$r.type
            linkKind   = [string]$r.linkKind
            createTime = [string]$r.createTime
        }
    }

    $json = [pscustomobject]@{ symlinks = $arr } | ConvertTo-Json -Depth 5
    $tmp  = $script:ConfigPath + '.tmp'
    [IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $tmp -Destination $script:ConfigPath -Force
}


# ---------------------------------------------------------------------------
# 界面辅助
# ---------------------------------------------------------------------------

function Add-RecordRow {
    param($Record)
    $state = Get-LinkState $Record
    $idx = $script:DataGridView.Rows.Add(
        $state.Text,
        $Record.sourcePath,
        $Record.targetPath,
        (Get-KindText $Record),
        $Record.createTime
    )
    $row = $script:DataGridView.Rows[$idx]
    $row.Tag = $Record
    $row.Cells[0].Style.ForeColor = $state.Color
    $row.Cells[0].Style.Font = $script:BoldFont
    $row.Cells[0].ToolTipText = $state.Message
}

function Get-CurrentRecords {
    $list = New-Object System.Collections.ArrayList
    foreach ($row in $script:DataGridView.Rows) {
        if ($null -ne $row.Tag) { [void]$list.Add($row.Tag) }
    }
    return $list
}

function Get-SelectedRecord {
    if ($null -eq $script:DataGridView) { return $null }
    if ($script:DataGridView.SelectedRows.Count -eq 0) { return $null }
    return $script:DataGridView.SelectedRows[0].Tag
}

function Refresh-AllStatuses {
    $total = 0
    $ok    = 0
    $bad   = 0

    foreach ($row in $script:DataGridView.Rows) {
        $rec = $row.Tag
        if ($null -eq $rec) { continue }
        $total++

        $state = Get-LinkState $rec
        $row.Cells[0].Value = $state.Text
        $row.Cells[0].Style.ForeColor = $state.Color
        $row.Cells[0].ToolTipText = $state.Message
        $row.Cells[1].Value = $rec.sourcePath
        $row.Cells[2].Value = $rec.targetPath
        $row.Cells[3].Value = Get-KindText $rec
        $row.Cells[4].Value = $rec.createTime

        if ($state.Code -eq 'Ok') { $ok++ } else { $bad++ }
    }

    $script:StatusLabel.Text = "共 $total 条　|　正常 $ok　|　异常 $bad　|　配置：$($script:ConfigPath)"
}

function Open-LinkRecord {
    param($Record)
    if ($null -eq $Record) {
        [void][System.Windows.Forms.MessageBox]::Show('请先选择一条记录。', '提示',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $state = Get-LinkState $Record
    if ($state.Code -ne 'Ok' -and $state.Code -ne 'Mismatch') {
        [void][System.Windows.Forms.MessageBox]::Show(
            "链接当前不可用：$($state.Text)`n`n$($state.Message)",
            '无法打开',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    try {
        $item = Get-Item -LiteralPath $Record.targetPath -Force -ErrorAction Stop
        if ($item.PSIsContainer) {
            Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $Record.targetPath + '"')
        } else {
            Start-Process -FilePath $Record.targetPath
        }
    } catch {
        Show-ErrorMessage '无法打开' (Get-FriendlyError $_.Exception) $script:MainForm
    }
}

function Copy-ToClipboard {
    param([string]$Text)
    try {
        [System.Windows.Forms.Clipboard]::SetText($Text)
    } catch {
        Show-ErrorMessage '复制失败' $_.Exception.Message $script:MainForm
    }
}

function Add-HelpText {
    param(
        $Box,
        [string]$Text,
        [string]$Style = 'Body',
        [switch]$NoNewline
    )
    if ($null -eq $Box) { return }
    $Box.SelectionStart = $Box.TextLength
    $Box.SelectionLength = 0
    switch ($Style) {
        'Heading' {
            $Box.SelectionFont = $script:HelpHeadingFont
            $Box.SelectionColor = [Drawing.Color]::FromArgb(28, 78, 140)
            $Box.SelectionBackColor = [Drawing.Color]::White
        }
        'Example' {
            $Box.SelectionFont = $script:UiFont
            $Box.SelectionColor = [Drawing.Color]::FromArgb(45, 65, 95)
            $Box.SelectionBackColor = [Drawing.Color]::FromArgb(244, 247, 252)
        }
        'Muted' {
            $Box.SelectionFont = $script:UiFont
            $Box.SelectionColor = [Drawing.Color]::FromArgb(110, 110, 110)
            $Box.SelectionBackColor = [Drawing.Color]::White
        }
        default {
            $Box.SelectionFont = $script:UiFont
            $Box.SelectionColor = [Drawing.Color]::FromArgb(45, 45, 45)
            $Box.SelectionBackColor = [Drawing.Color]::White
        }
    }
    if ($NoNewline) {
        $Box.AppendText($Text)
    } else {
        $Box.AppendText($Text + "`n")
    }
}

function New-HelpTab {
    param([string]$Text)
    $page = New-Object System.Windows.Forms.TabPage
    $page.Text = $Text
    $page.BackColor = [Drawing.Color]::White
    $page.Padding = New-Object System.Windows.Forms.Padding(8)

    $box = New-Object System.Windows.Forms.RichTextBox
    $box.Dock = 'Fill'
    $box.ReadOnly = $true
    $box.BorderStyle = 'None'
    $box.BackColor = [Drawing.Color]::White
    $box.Font = $script:UiFont
    $box.ScrollBars = 'Vertical'
    $box.WordWrap = $true
    $box.DetectUrls = $false
    $box.HideSelection = $false

    $page.Controls.Add($box)
    return [pscustomobject]@{ Page = $page; Box = $box }
}

function Show-HelpDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '使用说明'
    $dlg.ClientSize = New-Object System.Drawing.Size(780, 620)
    $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.BackColor = [Drawing.Color]::White
    $dlg.Font = $script:UiFont

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Top'
    $header.Height = 68
    $header.BackColor = [Drawing.Color]::White

    $title = New-Object System.Windows.Forms.Label
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(18, 10)
    $title.Font = $script:HelpTitleFont
    $title.ForeColor = [Drawing.Color]::FromArgb(28, 78, 140)
    $title.Text = 'Source 和 Target 怎么填'

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.AutoSize = $true
    $subtitle.Location = New-Object System.Drawing.Point(20, 40)
    $subtitle.ForeColor = [Drawing.Color]::FromArgb(110, 110, 110)
    $subtitle.Text = 'Source 是真实数据，Target 是新建的链接'

    $header.Controls.Add($title)
    $header.Controls.Add($subtitle)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $tabs.Font = $script:UiFont
    $tabs.Padding = New-Object System.Drawing.Point(16, 6)
    $tabs.ShowToolTips = $false

    $t1 = New-HelpTab '快速上手'
    Add-HelpText $t1.Box 'Source（源）' 'Heading'
    Add-HelpText $t1.Box '• 已经存在的真实文件/文件夹'
    Add-HelpText $t1.Box '• 数据实际存放的位置'
    Add-HelpText $t1.Box '• 必须是已存在的路径'
    Add-HelpText $t1.Box '• 删除链接不会删除 Source'
    Add-HelpText $t1.Box '例：D:\Data\Project' 'Example'
    Add-HelpText $t1.Box ''

    Add-HelpText $t1.Box 'Target（目标）' 'Heading'
    Add-HelpText $t1.Box '• 要新建的链接路径，必须还不存在'
    Add-HelpText $t1.Box '• 最后一段就是链接名称'
    Add-HelpText $t1.Box '• 不能只选一个已有文件夹'
    Add-HelpText $t1.Box '例：C:\Links\Project（Project 是名称）' 'Example'
    Add-HelpText $t1.Box '方向：D:\Data\Project  →  C:\Links\Project' 'Example'
    Add-HelpText $t1.Box '双击 C:\Links\Project 就会打开 D:\Data\Project。'
    Add-HelpText $t1.Box ''

    Add-HelpText $t1.Box '三步创建' 'Heading'
    Add-HelpText $t1.Box '1. 点“添加链接”'
    Add-HelpText $t1.Box '2. Source 选真实文件/文件夹；Target 选父文件夹并输入链接名称'
    Add-HelpText $t1.Box '3. 类型保持“自动”，点“创建”' -NoNewline

    $t2 = New-HelpTab '创建链接'
    Add-HelpText $t2.Box '创建前会检查' 'Heading'
    Add-HelpText $t2.Box '• Source 必须已经存在'
    Add-HelpText $t2.Box '• Target 的父文件夹必须已经存在'
    Add-HelpText $t2.Box '• Target 本身必须还不存在'
    Add-HelpText $t2.Box '• 名称不能包含 \ / : * ? " < > |，也不能用 CON、PRN 等保留名'
    Add-HelpText $t2.Box ''

    Add-HelpText $t2.Box 'Target 怎么填' 'Heading'
    Add-HelpText $t2.Box '「浏览」只能选择父文件夹，链接名称要自己输入。'
    Add-HelpText $t2.Box "父文件夹：C:\Links`n链接名称：Project`n完整 Target：C:\Links\Project" 'Example'
    Add-HelpText $t2.Box ''

    Add-HelpText $t2.Box '管理员权限' 'Heading'
    Add-HelpText $t2.Box '• Junction：不需要管理员'
    Add-HelpText $t2.Box '• 符号链接：通常需要管理员，或开启 Windows 开发者模式'
    Add-HelpText $t2.Box '• 文件链接只能用符号链接；文件夹优先用 Junction'
    Add-HelpText $t2.Box ''

    Add-HelpText $t2.Box '创建失败时' 'Heading'
    Add-HelpText $t2.Box '弹窗会给出原因；详细错误写入程序目录下的 symlink-manager.log。' 'Muted' -NoNewline

    $t3 = New-HelpTab '链接类型'
    Add-HelpText $t3.Box '目录联接（Junction）' 'Heading'
    Add-HelpText $t3.Box '• 只用于文件夹'
    Add-HelpText $t3.Box '• 不需要管理员权限'
    Add-HelpText $t3.Box '• 兼容性最好，推荐默认使用'
    Add-HelpText $t3.Box '• 只能指向本地文件夹，不能指向网络路径'
    Add-HelpText $t3.Box ''

    Add-HelpText $t3.Box '目录符号链接（Symbolic Link）' 'Heading'
    Add-HelpText $t3.Box '• 用于文件夹'
    Add-HelpText $t3.Box '• 可跨盘，可指向网络路径（UNC）'
    Add-HelpText $t3.Box '• 通常需要管理员权限或开发者模式'
    Add-HelpText $t3.Box ''

    Add-HelpText $t3.Box '文件符号链接（Symbolic Link）' 'Heading'
    Add-HelpText $t3.Box '• 只用于单个文件'
    Add-HelpText $t3.Box '• 通常需要管理员权限或开发者模式'
    Add-HelpText $t3.Box '• 双击时用默认程序打开文件，不会打开文件夹'
    Add-HelpText $t3.Box ''

    Add-HelpText $t3.Box '自动模式' 'Heading'
    Add-HelpText $t3.Box '文件夹 → Junction（网络文件夹除外）'
    Add-HelpText $t3.Box '文件 → 符号链接' -NoNewline

    $t4 = New-HelpTab '注意事项'
    Add-HelpText $t4.Box '重要提醒' 'Heading'
    Add-HelpText $t4.Box '• 不要把链接放在 OneDrive、坚果云、百度网盘等同步目录'
    Add-HelpText $t4.Box '• 不要把链接放在网络共享或映射网络驱动器'
    Add-HelpText $t4.Box '• 删除链接请用本工具的“删除链接”，它只删除链接本身'
    Add-HelpText $t4.Box '• Source 和被链接的数据要保留在原位置'
    Add-HelpText $t4.Box ''

    Add-HelpText $t4.Box '双击链接打不开？' 'Heading'
    Add-HelpText $t4.Box '1. 看“状态”列和鼠标提示'
    Add-HelpText $t4.Box '2. 目标失效：Source 被移动/删除，或磁盘/网络位置未连接'
    Add-HelpText $t4.Box '3. 确认类型：cmd /c dir /AL "链接所在文件夹"'
    Add-HelpText $t4.Box '   目录链接应显示 <SYMLINKD> 或 <JUNCTION>'
    Add-HelpText $t4.Box '4. 确认链接不在同步盘或网络共享上'
    Add-HelpText $t4.Box ''

    Add-HelpText $t4.Box '状态说明' 'Heading'
    Add-HelpText $t4.Box '正常：链接存在，目标可访问'
    Add-HelpText $t4.Box '目标失效：Source 或其链接目标已不存在'
    Add-HelpText $t4.Box '链接不存在：链接路径被手动删除'
    Add-HelpText $t4.Box '不是链接：路径现在是普通文件/文件夹'
    Add-HelpText $t4.Box '记录不一致：链接仍可用，但实际指向与记录不同' -NoNewline

    foreach ($tab in @($t1, $t2, $t3, $t4)) {
        [void]$tabs.TabPages.Add($tab.Page)
        $tab.Box.SelectionStart = 0
    }

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = 48
    $panel.BackColor = [Drawing.Color]::White

    $hint = New-Object System.Windows.Forms.Label
    $hint.AutoSize = $true
    $hint.Location = New-Object System.Drawing.Point(18, 15)
    $hint.ForeColor = [Drawing.Color]::FromArgb(110, 110, 110)
    $hint.Text = '提示：按 Esc 也可以关闭'

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = '关闭'
    $btn.Size = New-Object System.Drawing.Size(100, 32)
    # 窗体固定大小，用固定位置；不要加 Right 锚定（面板初始宽度会让按钮跑到外面）。
    $btn.Location = New-Object System.Drawing.Point(660, 8)

    $panel.Controls.Add($hint)
    $panel.Controls.Add($btn)

    $dlg.Controls.Add($tabs)
    $dlg.Controls.Add($panel)
    $dlg.Controls.Add($header)

    $dlg.CancelButton = $btn
    $btn.Add_Click({ $dlg.Close() })
    $dlg.Add_Shown({
        foreach ($page in $tabs.TabPages) {
            foreach ($ctrl in $page.Controls) {
                if ($ctrl -is [System.Windows.Forms.RichTextBox]) {
                    $ctrl.SelectionStart = 0
                    $ctrl.ScrollToCaret()
                }
            }
        }
    })

    [void]$dlg.ShowDialog($script:MainForm)
}

function Resolve-Kind {
    param(
        [int]$Choice,
        $SourceItem,
        [string]$SourcePath
    )
    $isDir = [bool]$SourceItem.PSIsContainer
    switch ($Choice) {
        0 {
            if ($isDir) {
                if (Test-NetworkPath $SourcePath) { return 'SymbolicLink' }
                return 'Junction'
            }
            return 'SymbolicLink'
        }
        1 {
            if (-not $isDir) { throw '目录符号链接只能用于文件夹。' }
            return 'SymbolicLink'
        }
        2 {
            if (-not $isDir) { throw 'Junction 只能用于文件夹。' }
            return 'Junction'
        }
        3 {
            if ($isDir) { throw '文件符号链接只能用于文件。' }
            return 'SymbolicLink'
        }
        default { return 'SymbolicLink' }
    }
}


# ---------------------------------------------------------------------------
# 添加链接对话框
# ---------------------------------------------------------------------------

function Show-AddDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '添加链接'
    $dlg.ClientSize = New-Object System.Drawing.Size(760, 562)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = $script:UiFont
    $dlg.Tag = $null

    $intro = New-Object System.Windows.Forms.Label
    $intro.Location = New-Object System.Drawing.Point(16, 10)
    $intro.Size = New-Object System.Drawing.Size(728, 24)
    $intro.Text = 'Source（真实数据）  →  Target（新链接）'
    $intro.Font = $script:BoldFont
    $intro.ForeColor = [Drawing.Color]::FromArgb(40, 90, 160)

    # --- Source -------------------------------------------------------------
    $grpSrc = New-Object System.Windows.Forms.GroupBox
    $grpSrc.Location = New-Object System.Drawing.Point(16, 38)
    $grpSrc.Size = New-Object System.Drawing.Size(728, 150)
    $grpSrc.Text = 'Source（源）— 已存在的真实文件/文件夹'

    $lblSrc = New-Object System.Windows.Forms.Label
    $lblSrc.Location = New-Object System.Drawing.Point(14, 22)
    $lblSrc.Size = New-Object System.Drawing.Size(690, 20)
    $lblSrc.Text = '选择你原来的文件或文件夹（数据实际所在位置）：'

    $srcBox = New-Object System.Windows.Forms.TextBox
    $srcBox.Location = New-Object System.Drawing.Point(14, 44)
    $srcBox.Size = New-Object System.Drawing.Size(520, 24)

    $btnSrcFolder = New-Object System.Windows.Forms.Button
    $btnSrcFolder.Location = New-Object System.Drawing.Point(540, 42)
    $btnSrcFolder.Size = New-Object System.Drawing.Size(80, 28)
    $btnSrcFolder.Text = '文件夹...'

    $btnSrcFile = New-Object System.Windows.Forms.Button
    $btnSrcFile.Location = New-Object System.Drawing.Point(626, 42)
    $btnSrcFile.Size = New-Object System.Drawing.Size(76, 28)
    $btnSrcFile.Text = '文件...'

    $srcHint = New-Object System.Windows.Forms.Label
    $srcHint.Location = New-Object System.Drawing.Point(14, 82)
    $srcHint.Size = New-Object System.Drawing.Size(700, 44)
    $srcHint.ForeColor = [Drawing.Color]::DimGray
    $srcHint.Text = "例：D:\Data\Project`nSource 必须已经存在；删除链接不会删除这里的任何内容。"

    $grpSrc.Controls.AddRange(@($lblSrc, $srcBox, $btnSrcFolder, $btnSrcFile, $srcHint))

    # --- Target -------------------------------------------------------------
    $grpTgt = New-Object System.Windows.Forms.GroupBox
    $grpTgt.Location = New-Object System.Drawing.Point(16, 196)
    $grpTgt.Size = New-Object System.Drawing.Size(728, 176)
    $grpTgt.Text = 'Target（目标）— 要新建的链接路径（不能已存在）'

    $lblParent = New-Object System.Windows.Forms.Label
    $lblParent.Location = New-Object System.Drawing.Point(14, 22)
    $lblParent.Size = New-Object System.Drawing.Size(690, 20)
    $lblParent.Text = '链接放在哪个文件夹（必须已存在）：'

    $parentBox = New-Object System.Windows.Forms.TextBox
    $parentBox.Location = New-Object System.Drawing.Point(14, 42)
    $parentBox.Size = New-Object System.Drawing.Size(600, 24)

    $btnParent = New-Object System.Windows.Forms.Button
    $btnParent.Location = New-Object System.Drawing.Point(620, 40)
    $btnParent.Size = New-Object System.Drawing.Size(82, 28)
    $btnParent.Text = '浏览...'

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = New-Object System.Drawing.Point(14, 76)
    $lblName.Size = New-Object System.Drawing.Size(690, 20)
    $lblName.Text = '链接名称（新建，不能与已有文件/文件夹重名）：'

    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Location = New-Object System.Drawing.Point(14, 96)
    $nameBox.Size = New-Object System.Drawing.Size(300, 24)

    $preview = New-Object System.Windows.Forms.Label
    $preview.Location = New-Object System.Drawing.Point(14, 126)
    $preview.Size = New-Object System.Drawing.Size(700, 20)
    $preview.ForeColor = [Drawing.Color]::FromArgb(40, 90, 160)
    $preview.Text = '完整 Target：—'

    $tgtHint = New-Object System.Windows.Forms.Label
    $tgtHint.Location = New-Object System.Drawing.Point(14, 148)
    $tgtHint.Size = New-Object System.Drawing.Size(700, 20)
    $tgtHint.ForeColor = [Drawing.Color]::DimGray
    $tgtHint.Text = 'Browse 只能选父文件夹，链接名请自己输入。Target 不能是已存在的路径。'

    $grpTgt.Controls.AddRange(@($lblParent, $parentBox, $btnParent, $lblName, $nameBox, $preview, $tgtHint))

    # --- 类型 ---------------------------------------------------------------
    $grpType = New-Object System.Windows.Forms.GroupBox
    $grpType.Location = New-Object System.Drawing.Point(16, 380)
    $grpType.Size = New-Object System.Drawing.Size(728, 86)
    $grpType.Text = '链接类型'

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Location = New-Object System.Drawing.Point(14, 24)
    $lblType.Size = New-Object System.Drawing.Size(70, 20)
    $lblType.Text = '类型：'

    $typeCombo = New-Object System.Windows.Forms.ComboBox
    $typeCombo.Location = New-Object System.Drawing.Point(84, 20)
    $typeCombo.Size = New-Object System.Drawing.Size(400, 24)
    $typeCombo.DropDownStyle = 'DropDownList'
    [void]$typeCombo.Items.Add('自动（文件夹→Junction，文件→符号链接）')
    [void]$typeCombo.Items.Add('目录符号链接 (Symbolic Link)')
    [void]$typeCombo.Items.Add('目录联接 (Junction，无需管理员)')
    [void]$typeCombo.Items.Add('文件符号链接 (Symbolic Link)')

    $typeHint = New-Object System.Windows.Forms.Label
    $typeHint.Location = New-Object System.Drawing.Point(14, 52)
    $typeHint.Size = New-Object System.Drawing.Size(700, 24)
    $typeHint.ForeColor = [Drawing.Color]::DimGray

    $grpType.Controls.AddRange(@($lblType, $typeCombo, $typeHint))

    $errLabel = New-Object System.Windows.Forms.Label
    $errLabel.Location = New-Object System.Drawing.Point(16, 474)
    $errLabel.Size = New-Object System.Drawing.Size(505, 38)
    $errLabel.ForeColor = [Drawing.Color]::Firebrick

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Location = New-Object System.Drawing.Point(540, 516)
    $btnCreate.Size = New-Object System.Drawing.Size(100, 32)
    $btnCreate.Text = '创建'

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(648, 516)
    $btnCancel.Size = New-Object System.Drawing.Size(92, 32)
    $btnCancel.Text = '取消'

    $dlg.Controls.AddRange(@($intro, $grpSrc, $grpTgt, $grpType, $errLabel, $btnCreate, $btnCancel))


    # --- 事件 ---------------------------------------------------------------
    $btnSrcFolder.Add_Click({
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = '选择已存在的 Source 文件夹'
        $fb.ShowNewFolderButton = $false
        if ($fb.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) {
            $srcBox.Text = $fb.SelectedPath
        }
    })

    $btnSrcFile.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = '选择已存在的 Source 文件'
        $ofd.CheckFileExists = $true
        $ofd.Multiselect = $false
        if ($ofd.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) {
            $srcBox.Text = $ofd.FileName
        }
    })

    $btnParent.Add_Click({
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = '选择 Target 所在的文件夹（链接会建在这里）'
        $fb.ShowNewFolderButton = $true
        if ($fb.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) {
            $parentBox.Text = $fb.SelectedPath
        }
    })

    $updatePreview = {
        $p = $parentBox.Text.Trim()
        $n = $nameBox.Text.Trim()
        if ($p -and $n) {
            try { $preview.Text = '完整 Target：' + [IO.Path]::Combine($p, $n) }
            catch { $preview.Text = '完整 Target：—' }
        } else {
            $preview.Text = '完整 Target：—'
        }
    }
    $parentBox.Add_TextChanged($updatePreview)
    $nameBox.Add_TextChanged($updatePreview)

    $typeCombo.Add_SelectedIndexChanged({
        switch ($typeCombo.SelectedIndex) {
            0 { $typeHint.Text = '自动：文件夹用 Junction（无需管理员），文件用符号链接。' }
            1 { $typeHint.Text = '目录符号链接：可指向文件夹，通常需要管理员权限或开发者模式。' }
            2 { $typeHint.Text = 'Junction：不需要管理员，兼容性最好，只能链接本地文件夹。' }
            3 { $typeHint.Text = '文件符号链接：只能链接文件，通常需要管理员权限或开发者模式。' }
        }
    })
    $typeCombo.SelectedIndex = 0

    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.AcceptButton = $btnCreate
    $dlg.CancelButton = $btnCancel

    $btnCreate.Add_Click({
        $finished = $false
        $btnCreate.Enabled = $false
        $dlg.UseWaitCursor = $true
        $errLabel.Text = ''

        try {
            $srcItem = Get-ExistingItem -Path $srcBox.Text -FieldName 'Source'
            $srcPath = $srcItem.FullName

            $parentItem = Get-ExistingItem -Path $parentBox.Text -FieldName 'Target 的父文件夹'
            if (-not $parentItem.PSIsContainer) {
                throw 'Target 的父路径必须是文件夹。'
            }
            $parentPath = $parentItem.FullName

            $name = $nameBox.Text.Trim()
            Assert-ValidLinkName $name

            $linkPath = [IO.Path]::Combine($parentPath, $name)
            if ($linkPath.Length -gt 259) {
                throw "Target 路径太长（$($linkPath.Length) 个字符，超过 259）。请缩短名称或换一个父文件夹。"
            }

            $existing = Get-Item -LiteralPath $linkPath -Force -ErrorAction SilentlyContinue
            if ($null -ne $existing) {
                throw "Target 已存在：`n$linkPath`n请换一个链接名称，或先删除旧链接。"
            }

            $kind = Resolve-Kind -Choice $typeCombo.SelectedIndex -SourceItem $srcItem -SourcePath $srcPath
            Assert-LinkLocationSupported $parentPath

            if ($kind -eq 'SymbolicLink' -and -not $script:IsAdmin) {
                $ans = [System.Windows.Forms.MessageBox]::Show(
                    $dlg,
                    "当前没有管理员权限，创建符号链接可能失败（除非已开启 Windows 开发者模式）。`n`n" +
                    "文件夹可以改用 Junction（不需要管理员）；文件链接则必须以管理员身份运行。`n`n仍要尝试吗？",
                    '权限提示',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }

            $cloud = Get-CloudProvider $srcPath
            if ([string]::IsNullOrWhiteSpace($cloud)) { $cloud = Get-CloudProvider $linkPath }
            if (-not [string]::IsNullOrWhiteSpace($cloud)) {
                $ans = [System.Windows.Forms.MessageBox]::Show(
                    $dlg,
                    "检测到路径位于 $cloud 同步目录。`n`n" +
                    "同步盘可能不保留软链接，导致链接失效或变成普通文件。`n" +
                    "建议把 Source 和 Target 都放在本地非同步目录。`n`n仍要继续吗？",
                    '同步目录警告',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }

            if (Test-NetworkPath $srcPath) {
                $ans = [System.Windows.Forms.MessageBox]::Show(
                    $dlg,
                    "Source 位于网络路径。`n`n" +
                    "资源管理器对指向网络路径的符号链接支持有限，可能无法双击打开。`n" +
                    "建议直接使用 UNC 路径（\\server\share），或把 Source 放到本地。`n`n仍要继续吗？",
                    '网络路径警告',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }

            $record = New-Link -SourcePath $srcPath -LinkPath $linkPath -Kind $kind
            $dlg.Tag = $record
            $finished = $true
            $dlg.Close()
        } catch {
            $friendly = Get-FriendlyError $_.Exception
            $errLabel.Text = '创建失败，详情见弹窗。'
            [void][System.Windows.Forms.MessageBox]::Show(
                $dlg, $friendly, '创建失败',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "创建链接失败：$($_.Exception.ToString())" 'ERROR'
        } finally {
            if (-not $finished) {
                try {
                    $btnCreate.Enabled = $true
                    $dlg.UseWaitCursor = $false
                } catch { }
            }
        }
    })

    [void]$dlg.ShowDialog($script:MainForm)
    return $dlg.Tag
}


# ---------------------------------------------------------------------------
# 删除记录
# ---------------------------------------------------------------------------

function Remove-SelectedRecord {
    $rec = Get-SelectedRecord
    if ($null -eq $rec) {
        [void][System.Windows.Forms.MessageBox]::Show('请先选择一条记录。', '提示',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $row = $script:DataGridView.SelectedRows[0]
    $item = Get-Item -LiteralPath $rec.targetPath -Force -ErrorAction SilentlyContinue

    if ($null -eq $item) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            $script:MainForm,
            "链接已经不存在：`n$($rec.targetPath)`n`n是否只删除这条记录？",
            '链接不存在',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    } else {
        if (-not (Test-IsReparsePoint $item)) {
            Show-ErrorMessage '拒绝删除' ("该路径现在不是链接（不是重解析点），工具不会删除它：`n$($rec.targetPath)`n`n" +
                '如果这是真实文件夹，请手动确认后再处理。') $script:MainForm
            return
        }
        $linkType = [string]$item.LinkType
        if ($linkType -notin @('SymbolicLink', 'Junction')) {
            Show-ErrorMessage '拒绝删除' ("该路径不是符号链接/目录联接（LinkType=$linkType），工具不会删除它：`n$($rec.targetPath)`n`n" +
                '可能是云占位符或其他重解析点，请手动确认后再处理。') $script:MainForm
            return
        }

        $state = Get-LinkState $rec
        $actualLine = ''
        if (-not [string]::IsNullOrWhiteSpace([string]$state.Target)) {
            $actual = ([string]$state.Target).TrimEnd('\')
            $recorded = ([string]$rec.sourcePath).TrimEnd('\')
            if (-not [string]::Equals($actual, $recorded, [StringComparison]::OrdinalIgnoreCase)) {
                $actualLine = "`n`n链接实际指向：`n$($state.Target)"
            }
        }

        $ans = [System.Windows.Forms.MessageBox]::Show(
            $script:MainForm,
            "确定删除这个链接吗？`n`nTarget：$($rec.targetPath)`nSource：$($rec.sourcePath)$actualLine`n`n" +
            "只会删除链接本身，不会删除 Source 里的任何内容。",
            '确认删除',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        try {
            Remove-LinkSafely $rec.targetPath
        } catch {
            Show-ErrorMessage '删除失败' (Get-FriendlyError $_.Exception) $script:MainForm
            return
        }
    }

    $script:DataGridView.Rows.Remove($row)
    try {
        Save-Records (Get-CurrentRecords)
    } catch {
        Show-ErrorMessage '保存记录失败' $_.Exception.Message $script:MainForm
    }
    Refresh-AllStatuses
}

# ---------------------------------------------------------------------------
# 主界面
# ---------------------------------------------------------------------------

function Start-Gui {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '软链接管理器 — Symlink Manager'
    $form.ClientSize = New-Object System.Drawing.Size(980, 640)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)
    $form.StartPosition = 'CenterScreen'
    $form.Font = $script:UiFont
    $form.BackColor = [Drawing.Color]::FromArgb(246, 248, 251)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Top'
    $header.Height = 82
    $header.Width = $form.ClientSize.Width
    $header.BackColor = [Drawing.Color]::White

    $title = New-Object System.Windows.Forms.Label
    $title.Location = New-Object System.Drawing.Point(16, 8)
    $title.Size = New-Object System.Drawing.Size(330, 30)
    $title.Text = '软链接管理器'
    $title.Font = New-Object System.Drawing.Font($script:UiFont.FontFamily, 15, [System.Drawing.FontStyle]::Bold)

    $adminLabel = New-Object System.Windows.Forms.Label
    $adminLabel.Location = New-Object System.Drawing.Point(524, 13)
    $adminLabel.Size = New-Object System.Drawing.Size(440, 20)
    $adminLabel.TextAlign = 'MiddleRight'
    $adminLabel.Anchor = 'Top,Right'

    $btnElevate = New-Object System.Windows.Forms.Button
    $btnElevate.Location = New-Object System.Drawing.Point(586, 12)
    $btnElevate.Size = New-Object System.Drawing.Size(170, 32)
    $btnElevate.Text = '以管理员身份重启'
    $btnElevate.Visible = $false

    if ($script:IsAdmin) {
        $adminLabel.Text = '● 已以管理员身份运行'
        $adminLabel.ForeColor = [Drawing.Color]::ForestGreen
    } else {
        $adminLabel.Text = '● 未以管理员运行：Junction 可用，符号链接可能需要提升'
        $adminLabel.ForeColor = [Drawing.Color]::DarkOrange
        $btnElevate.Visible = $true
    }

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Location = New-Object System.Drawing.Point(18, 46)
    $subtitle.Size = New-Object System.Drawing.Size(944, 26)
    $subtitle.Anchor = 'Top,Left,Right'
    $subtitle.ForeColor = [Drawing.Color]::DimGray
    $subtitle.Text = 'Source（源）= 已存在的真实文件/文件夹　→　Target（目标）= 要新建的链接路径（不能已存在）'

    $header.Controls.AddRange(@($title, $adminLabel, $subtitle))

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.BackgroundColor = [Drawing.Color]::White
    $grid.BorderStyle = 'None'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.MultiSelect = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'None'
    $grid.ColumnHeadersHeight = 36
    $grid.RowTemplate.Height = 28
    $grid.ShowCellToolTips = $true
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(243, 246, 250)
    $grid.ColumnHeadersDefaultCellStyle.Font = $script:BoldFont
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(250, 251, 253)
    $grid.GridColor = [Drawing.Color]::FromArgb(228, 232, 238)
    $grid.AllowUserToOrderColumns = $false

    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.Name = 'status'
    $colStatus.HeaderText = '状态'
    $colStatus.Width = 90
    $colStatus.AutoSizeMode = 'None'
    $colStatus.SortMode = 'NotSortable'

    $colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSource.Name = 'source'
    $colSource.HeaderText = 'Source（源路径 / 真实数据）'
    $colSource.AutoSizeMode = 'Fill'
    $colSource.FillWeight = 42
    $colSource.MinimumWidth = 180
    $colSource.SortMode = 'NotSortable'

    $colTarget = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTarget.Name = 'target'
    $colTarget.HeaderText = 'Target（链接路径 / 新建）'
    $colTarget.AutoSizeMode = 'Fill'
    $colTarget.FillWeight = 42
    $colTarget.MinimumWidth = 180
    $colTarget.SortMode = 'NotSortable'

    $colKind = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colKind.Name = 'kind'
    $colKind.HeaderText = '类型'
    $colKind.Width = 140
    $colKind.AutoSizeMode = 'None'
    $colKind.SortMode = 'NotSortable'

    $colTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTime.Name = 'time'
    $colTime.HeaderText = '创建时间'
    $colTime.Width = 150
    $colTime.AutoSizeMode = 'None'
    $colTime.SortMode = 'NotSortable'

    [void]$grid.Columns.Add($colStatus)
    [void]$grid.Columns.Add($colSource)
    [void]$grid.Columns.Add($colTarget)
    [void]$grid.Columns.Add($colKind)
    [void]$grid.Columns.Add($colTime)

    $bottom = New-Object System.Windows.Forms.Panel
    $bottom.Dock = 'Bottom'
    $bottom.Height = 96
    $bottom.Width = $form.ClientSize.Width
    $bottom.BackColor = [Drawing.Color]::White

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Location = New-Object System.Drawing.Point(16, 12)
    $btnAdd.Size = New-Object System.Drawing.Size(110, 32)
    $btnAdd.Text = '添加链接'

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Location = New-Object System.Drawing.Point(136, 12)
    $btnOpen.Size = New-Object System.Drawing.Size(100, 32)
    $btnOpen.Text = '打开链接'

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Location = New-Object System.Drawing.Point(246, 12)
    $btnRefresh.Size = New-Object System.Drawing.Size(110, 32)
    $btnRefresh.Text = '校验 / 刷新'

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Location = New-Object System.Drawing.Point(366, 12)
    $btnDelete.Size = New-Object System.Drawing.Size(100, 32)
    $btnDelete.Text = '删除链接'

    $btnHelp = New-Object System.Windows.Forms.Button
    $btnHelp.Location = New-Object System.Drawing.Point(476, 12)
    $btnHelp.Size = New-Object System.Drawing.Size(100, 32)
    $btnHelp.Text = '使用说明'

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(16, 56)
    $statusLabel.Size = New-Object System.Drawing.Size(944, 26)
    $statusLabel.Anchor = 'Top,Left,Right'
    $statusLabel.ForeColor = [Drawing.Color]::DimGray

    $bottom.Controls.AddRange(@($btnAdd, $btnOpen, $btnRefresh, $btnDelete, $btnHelp, $btnElevate, $statusLabel))


    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $miOpen = $menu.Items.Add('打开链接')
    $miCopyTarget = $menu.Items.Add('复制 Target（链接路径）')
    $miCopySource = $menu.Items.Add('复制 Source（源路径）')
    $miLocate = $menu.Items.Add('在资源管理器中显示链接')
    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miVerify = $menu.Items.Add('校验 / 刷新')
    $miDelete = $menu.Items.Add('删除链接')
    $grid.ContextMenuStrip = $menu

    $miOpen.Add_Click({ Open-LinkRecord (Get-SelectedRecord) })
    $miCopyTarget.Add_Click({
        $r = Get-SelectedRecord
        if ($null -ne $r) { Copy-ToClipboard $r.targetPath }
    })
    $miCopySource.Add_Click({
        $r = Get-SelectedRecord
        if ($null -ne $r) { Copy-ToClipboard $r.sourcePath }
    })
    $miLocate.Add_Click({
        $r = Get-SelectedRecord
        if ($null -ne $r) {
            try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $r.targetPath + '"') }
            catch { Show-ErrorMessage '无法定位' (Get-FriendlyError $_.Exception) $script:MainForm }
        }
    })
    $miVerify.Add_Click({ Refresh-AllStatuses })
    $miDelete.Add_Click({ Remove-SelectedRecord })

    $grid.Add_CellMouseDown({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
            $grid.ClearSelection()
            $grid.Rows[$e.RowIndex].Selected = $true
        }
    })

    $grid.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) { Open-LinkRecord (Get-SelectedRecord) }
    })

    $grid.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
            Remove-SelectedRecord
            $e.Handled = $true
        } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            Open-LinkRecord (Get-SelectedRecord)
            $e.Handled = $true
        }
    })

    $btnAdd.Add_Click({
        $rec = Show-AddDialog
        if ($null -eq $rec) { return }

        Add-RecordRow $rec
        try {
            Save-Records (Get-CurrentRecords)
        } catch {
            Show-ErrorMessage '链接已创建，但保存记录失败' (
                "链接已经创建成功：`n$($rec.targetPath)`n`n" +
                "但写入配置文件失败：`n$($_.Exception.Message)`n`n" +
                "本次记录会在程序关闭后丢失。") $script:MainForm
        }
        Refresh-AllStatuses
    })

    $btnOpen.Add_Click({ Open-LinkRecord (Get-SelectedRecord) })
    $btnDelete.Add_Click({ Remove-SelectedRecord })
    $btnRefresh.Add_Click({ Refresh-AllStatuses })
    $btnHelp.Add_Click({ Show-HelpDialog })

    $btnElevate.Add_Click({
        try {
            $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            if (-not (Test-Path -LiteralPath $exe)) {
                $exe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
            }
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $exe
            $psi.Verb = 'runas'
            $psi.UseShellExecute = $true
            $psi.Arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "' + $PSCommandPath + '"'
            [System.Diagnostics.Process]::Start($psi) | Out-Null
            $form.Close()
        } catch {
            Show-ErrorMessage '无法提升权限' (Get-FriendlyError $_.Exception) $form
        }
    })

    $form.Controls.Add($grid)
    $form.Controls.Add($bottom)
    $form.Controls.Add($header)

    $script:MainForm = $form
    $script:DataGridView = $grid
    $script:StatusLabel = $statusLabel

    $records = Read-Records
    foreach ($rec in $records) { Add-RecordRow $rec }

    $form.Add_Shown({
        Refresh-AllStatuses
        if (-not [string]::IsNullOrWhiteSpace($script:StartupWarning)) {
            [void][System.Windows.Forms.MessageBox]::Show($script:MainForm, $script:StartupWarning, '配置提示',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            $script:StartupWarning = $null
        }
    })

    [void]$form.ShowDialog()
}


# ---------------------------------------------------------------------------
# 自检
# ---------------------------------------------------------------------------

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Name
    )
    if ($Condition) {
        Write-Host "  PASS  $Name" -ForegroundColor Green
    } else {
        Write-Host "  FAIL  $Name" -ForegroundColor Red
        $script:SelfTestFailed++
    }
}

function Invoke-SelfTest {
    Write-Host ''
    Write-Host '=== 软链接管理器自检 ===' -ForegroundColor Cyan
    Write-Host ("管理员: {0}" -f $script:IsAdmin)
    Write-Host ''

    $root = Join-Path $env:TEMP ('mlink-selftest-' + [guid]::NewGuid().ToString('N'))
    $real = Join-Path $root 'real'
    $links = Join-Path $root 'links'
    $file = Join-Path $real 'a.txt'

    $oldConfig = $script:ConfigPath
    $created = New-Object System.Collections.ArrayList

    try {
        New-Item -ItemType Directory -Path $real -Force | Out-Null
        New-Item -ItemType Directory -Path $links -Force | Out-Null
        'hello' | Set-Content -LiteralPath $file -Encoding ASCII

        $dirLink  = Join-Path $links 'dirlink'
        $juncLink = Join-Path $links 'junc'
        $fileLink = Join-Path $links 'filelink'

        $r1 = New-Link -SourcePath $real -LinkPath $dirLink -Kind SymbolicLink
        [void]$created.Add($dirLink)
        Assert-True (Test-Path -LiteralPath (Join-Path $dirLink 'a.txt')) '目录符号链接可访问目标内容'
        Assert-True ((Get-LinkState $r1).Code -eq 'Ok') '目录符号链接状态为正常'

        $r2 = New-Link -SourcePath $real -LinkPath $juncLink -Kind Junction
        [void]$created.Add($juncLink)
        Assert-True (Test-Path -LiteralPath (Join-Path $juncLink 'a.txt')) 'Junction 可访问目标内容'
        Assert-True ((Get-LinkState $r2).Code -eq 'Ok') 'Junction 状态为正常'

        $r3 = New-Link -SourcePath $file -LinkPath $fileLink -Kind SymbolicLink
        [void]$created.Add($fileLink)
        Assert-True ((Get-Content -LiteralPath $fileLink -Raw).Trim() -eq 'hello') '文件符号链接内容正确'
        Assert-True ((Get-LinkState $r3).Code -eq 'Ok') '文件符号链接状态为正常'

        Remove-Item -LiteralPath $file -Force
        Assert-True ((Get-LinkState $r3).Code -eq 'Broken') '目标删除后状态为“目标失效”'
        'hello' | Set-Content -LiteralPath $file -Encoding ASCII
        Assert-True ((Get-LinkState $r3).Code -eq 'Ok') '目标恢复后状态回到正常'

        $script:ConfigPath = Join-Path $root 'test-symlinks.json'
        Save-Records @($r1, $r2, $r3)
        Assert-True (Test-Path -LiteralPath $script:ConfigPath) '配置文件写入成功'
        $loaded = Read-Records
        Assert-True ($loaded.Count -eq 3) '配置文件读回 3 条记录'
        $script:ConfigPath = $oldConfig

        $nameOk = $true
        try { Assert-ValidLinkName 'good-name' } catch { $nameOk = $false }
        Assert-True $nameOk '合法链接名称通过校验'
        $nameBad = $false
        try { Assert-ValidLinkName 'bad:name' } catch { $nameBad = $true }
        Assert-True $nameBad '非法链接名称被拒绝'

        $rootRelative = $false
        try { Get-InputFullPath -Path '\Links' -FieldName 'Test' | Out-Null } catch { $rootRelative = $true }
        Assert-True $rootRelative '拒绝根相对路径（\Links）'

        $driveRelative = $false
        try { Get-InputFullPath -Path 'C:Links' -FieldName 'Test' | Out-Null } catch { $driveRelative = $true }
        Assert-True $driveRelative '拒绝盘符相对路径（C:Links）'

        $uncPathOk = $false
        try { $uncPathOk = ((Get-InputFullPath -Path '\\server\share' -FieldName 'Test') -eq '\\server\share') } catch { }
        Assert-True $uncPathOk 'UNC 完整路径通过校验'

        $mockUnc = [pscustomobject]@{ FullName = 'C:\Links\unclink'; Target = @('UNC\server\share') }
        Assert-True ((Get-LinkTargetPath $mockUnc) -eq '\\server\share') 'UNC 链接目标解析正确'
        $mockUncExtended = [pscustomobject]@{ FullName = 'C:\Links\unclink2'; Target = @('\??\UNC\server\share') }
        Assert-True ((Get-LinkTargetPath $mockUncExtended) -eq '\\server\share') '扩展 UNC 链接目标解析正确'

        $unknownKind = Get-KindText ([pscustomobject]@{ linkKind = 'SymbolicLink'; type = '' })
        Assert-True ($unknownKind -eq '链接（类型未知）') '未知类型不会显示为文件符号链接'

        $nestedReal  = Join-Path $root 'nested-real'
        $nestedInner = Join-Path $root 'nested-inner'
        $nestedOuter = Join-Path $root 'nested-outer'
        New-Item -ItemType Directory -Path $nestedReal -Force | Out-Null
        $rInner = New-Link -SourcePath $nestedReal -LinkPath $nestedInner -Kind Junction
        [void]$created.Add($nestedInner)
        $rOuter = New-Link -SourcePath $nestedInner -LinkPath $nestedOuter -Kind SymbolicLink
        [void]$created.Add($nestedOuter)
        Assert-True ((Get-LinkState $rOuter).Code -eq 'Ok') '嵌套链接：内层目标存在时状态正常'
        Remove-Item -LiteralPath $nestedReal -Recurse -Force
        Assert-True ((Get-LinkState $rOuter).Code -eq 'Broken') '嵌套链接：内层目标删除后状态为目标失效'

        $loopSource = Join-Path $root 'loop-source'
        $loopAlias  = Join-Path $root 'loop-alias'
        New-Item -ItemType Directory -Path $loopSource -Force | Out-Null
        $rAlias = New-Link -SourcePath $loopSource -LinkPath $loopAlias -Kind Junction
        [void]$created.Add($loopAlias)
        $loopBlocked = $false
        $loopMessage = ''
        try {
            New-Link -SourcePath $loopSource -LinkPath (Join-Path $loopAlias 'inside') -Kind Junction | Out-Null
        } catch {
            $loopBlocked = $true
            $loopMessage = $_.Exception.Message
        }
        Assert-True ($loopBlocked -and $loopMessage -like '*循环链接*') '拒绝通过 Junction 父路径创建循环链接'

        $plain = Join-Path $root 'plain'
        New-Item -ItemType Directory -Path $plain -Force | Out-Null
        $refused = $false
        try { Remove-LinkSafely $plain } catch { $refused = $true }
        Assert-True ($refused -and (Test-Path -LiteralPath $plain)) '安全删除会拒绝普通文件夹'

        foreach ($p in $created) {
            if (Test-Path -LiteralPath $p) { Remove-LinkSafely $p }
        }
        Assert-True (Test-Path -LiteralPath (Join-Path $real 'a.txt')) '删除链接后 Source 内容仍然存在'
        Assert-True (-not (Test-Path -LiteralPath $dirLink)) '目录符号链接已删除'
        Assert-True (-not (Test-Path -LiteralPath $juncLink)) 'Junction 已删除'
        Assert-True (-not (Test-Path -LiteralPath $fileLink)) '文件符号链接已删除'
    } catch {
        Assert-True $false ("自检异常：" + $_.Exception.Message)
        Write-Host $_.Exception.ToString() -ForegroundColor DarkRed
    } finally {
        $script:ConfigPath = $oldConfig
        foreach ($p in $created) {
            $i = Get-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
            if ($null -ne $i -and (Test-IsReparsePoint $i)) {
                try { Remove-LinkSafely $p } catch { }
            }
        }
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host ''
    if ($script:SelfTestFailed -eq 0) {
        Write-Host '结果：全部通过' -ForegroundColor Green
        return $true
    }
    Write-Host ("结果：$($script:SelfTestFailed) 项失败") -ForegroundColor Red
    return $false
}

# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

Initialize-DataPaths
$script:IsAdmin = Test-Admin
Write-Log ("启动：PS={0}, Admin={1}, Config={2}" -f $PSVersionTable.PSVersion, $script:IsAdmin, $script:ConfigPath)

if ($SelfTest) {
    if (Invoke-SelfTest) { exit 0 } else { exit 1 }
}

try {
    Start-Gui
} catch {
    Write-Log "FATAL: $($_.Exception.ToString())" 'FATAL'
    Write-Host '程序启动失败：' -ForegroundColor Red
    Write-Host $_.Exception.ToString() -ForegroundColor Red
    exit 1
}
