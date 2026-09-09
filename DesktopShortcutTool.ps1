<#
.SYNOPSIS
    桌面快捷方式备份与恢复工具
.DESCRIPTION
    - 备份: 扫描"用户桌面"和"公共桌面"的所有快捷方式(.lnk/.url),
      导出完整元数据(JSON+CSV)并复制快捷方式文件本身。
    - 恢复: 一键将快捷方式还原到桌面, 支持新旧用户名路径重映射,
      恢复后自动检测目标程序是否缺失并生成报告。
    - 检查: 检测当前桌面上哪些快捷方式的目标已失效。
.NOTES
    直接双击"启动工具.bat"打开图形界面;
    也可命令行调用:
      DesktopShortcutTool.ps1 -Action backup
      DesktopShortcutTool.ps1 -Action restore -BackupDir <目录> [-OldUser a -NewUser b]
      DesktopShortcutTool.ps1 -Action check
#>
param(
    [ValidateSet('gui','backup','restore','check')]
    [string]$Action = 'gui',
    [string]$BackupDir = '',
    [string]$TargetDir = '',
    [string]$OldUser = '',
    [string]$NewUser = '',
    [switch]$SaveLayout,
    [switch]$RestoreLayout,
    [switch]$AutoRestore
)

$ErrorActionPreference = 'Stop'
$Script:ToolName = '桌面快捷方式备份与恢复工具'
$Script:ThisPath = $PSCommandPath
if (-not $Script:ThisPath) { $Script:ThisPath = $MyInvocation.MyCommand.Path }
$script:LastLayoutRestored = $false
# 基准目录: 直接运行脚本时用 $PSScriptRoot; ps2exe 打包成 exe 时 $PSScriptRoot 为空, 改用 exe 所在目录。
$Script:BaseDir = $PSScriptRoot
if (-not $Script:BaseDir) {
    try { $Script:BaseDir = [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch { }
}
if (-not $Script:BaseDir) { $Script:BaseDir = (Get-Location).Path }
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# ==================== 基础函数 ====================

function Get-DesktopDirs {
    $dirs = @()
    $userDesktop = [Environment]::GetFolderPath('Desktop')
    if ($userDesktop -and (Test-Path -LiteralPath $userDesktop)) {
        $dirs += [pscustomobject]@{ Kind = '用户桌面'; Path = $userDesktop }
    }
    $common = [Environment]::GetFolderPath('CommonDesktopDirectory')
    if ($common -and (Test-Path -LiteralPath $common) -and $common -ne $userDesktop) {
        $dirs += [pscustomobject]@{ Kind = '公共桌面'; Path = $common }
    }
    return $dirs
}

function Find-BackupRoot {
    param([string]$Dir)
    if (-not $Dir) {
        $cands = @(Get-ChildItem -LiteralPath $Script:BaseDir -Directory -Filter '快捷方式备份_*' -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
        if ($cands.Count -gt 0) { return $cands[0].FullName }
        return ''
    }
    if (Test-Path -LiteralPath (Join-Path $Dir 'backup_info.json')) { return $Dir }
    $cands = @(Get-ChildItem -LiteralPath $Dir -Directory -Filter '快捷方式备份_*' -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    if ($cands.Count -gt 0) { return $cands[0].FullName }
    $found = Get-ChildItem -LiteralPath $Dir -Recurse -Filter 'backup_info.json' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { return $found.DirectoryName }
    return ''
}

function Convert-Path {
    param([string]$Path, [string]$Old, [string]$New)
    if (-not $Path) { return $Path }
    if ($Old -and $New -and ($Old -ne $New)) { return $Path.Replace($Old, $New) }
    return $Path
}

function Convert-Icon {
    param([string]$Icon, [string]$Old, [string]$New)
    if (-not $Icon) { return $Icon }
    if ($Icon -match '^(.+),(\s*\d+)\s*$') {
        return (Convert-Path -Path $Matches[1] -Old $Old -New $New) + ',' + $Matches[2]
    }
    return (Convert-Path -Path $Icon -Old $Old -New $New)
}

# ==================== 备份 ====================

function Invoke-Backup {
    param(
        [string]$Dir = '',
        [scriptblock]$Log = { param($m) Write-Host $m },
        [switch]$SaveLayout
    )
    $Dir = ($Dir -replace '^True\s+', '').Trim()
    if (-not $Dir) {
        $Dir = Join-Path $Script:BaseDir ('快捷方式备份_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    }
    if (-not (Test-Path -LiteralPath $Dir)) {
        New-Item -ItemType Directory -Force -Path $Dir | Out-Null
    }
    & $Log ('备份保存到: ' + $Dir)

    $sh = New-Object -ComObject WScript.Shell
    $items = @()
    $copied = 0
    foreach ($d in (Get-DesktopDirs)) {
        & $Log ('正在扫描 ' + $d.Kind + ': ' + $d.Path)
        $subDir = Join-Path $Dir ('shortcuts\' + $d.Kind)
        $files = @(Get-ChildItem -LiteralPath $d.Path -File -ErrorAction SilentlyContinue)
        foreach ($f in $files) {
            $base = [ordered]@{}
            $base['Name'] = $f.Name
            $base['SourceDir'] = $d.Kind
            $base['FullPath'] = $f.FullName
            if ($f.Extension -ieq '.lnk') {
                $base['Type'] = 'lnk'
                try {
                    $s = $sh.CreateShortcut($f.FullName)
                    $base['TargetPath'] = [string]$s.TargetPath
                    $base['Arguments'] = [string]$s.Arguments
                    $base['WorkingDirectory'] = [string]$s.WorkingDirectory
                    $base['IconLocation'] = [string]$s.IconLocation
                    $base['Hotkey'] = [string]$s.Hotkey
                    $base['WindowStyle'] = [int]$s.WindowStyle
                    $base['Description'] = [string]$s.Description
                } catch {
                    $base['TargetPath'] = ''
                    $base['Note'] = '读取失败: ' + $_.Exception.Message
                }
                if (-not (Test-Path -LiteralPath $subDir)) { New-Item -ItemType Directory -Force -Path $subDir | Out-Null }
                Copy-Item -LiteralPath $f.FullName -Destination $subDir -Force
                $copied++
                & $Log ('  [备份] ' + $f.Name + '  ->  ' + $base['TargetPath'])
            }
            elseif ($f.Extension -ieq '.url') {
                $base['Type'] = 'url'
                $url = ''
                foreach ($line in @(Get-Content -LiteralPath $f.FullName -ErrorAction SilentlyContinue)) {
                    if ($line -match '^URL\s*=\s*(.+?)\s*$') { $url = $Matches[1]; break }
                }
                $base['Url'] = $url
                if (-not (Test-Path -LiteralPath $subDir)) { New-Item -ItemType Directory -Force -Path $subDir | Out-Null }
                Copy-Item -LiteralPath $f.FullName -Destination $subDir -Force
                $copied++
                & $Log ('  [备份] ' + $f.Name + '  ->  ' + $url)
            }
            else {
                $base['Type'] = 'other'
            }
            $items += [pscustomobject]$base
        }
    }
    $sh = $null

    $shortcutItems = @($items | Where-Object { $_.Type -ne 'other' })
    $otherCount = @($items).Count - $shortcutItems.Count

    $meta = [ordered]@{
        Tool        = $Script:ToolName
        Machine     = $env:COMPUTERNAME
        UserName    = $env:USERNAME
        BackupTime  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        DesktopDirs = @(Get-DesktopDirs | ForEach-Object { $_.Path })
        ShortcutCount = $shortcutItems.Count
        Items       = @($items)
    }
    $json = ConvertTo-Json -InputObject $meta -Depth 6
    [System.IO.File]::WriteAllText((Join-Path $Dir 'backup_info.json'), $json, (New-Object System.Text.UTF8Encoding $true))

    $shortcutItems |
        Select-Object Name, SourceDir, Type, TargetPath, Arguments, WorkingDirectory, IconLocation, Hotkey, Url |
        Export-Csv -Path (Join-Path $Dir '备份清单.csv') -NoTypeInformation -Encoding UTF8

    if ($SaveLayout) { [void](Save-DesktopLayout -Dir $Dir -Log $Log) }

    & $Log ('----------------------------------------')
    & $Log ('备份完成! 共备份 ' + $copied + ' 个快捷方式' + $(if ($otherCount -gt 0) { '(另有 ' + $otherCount + ' 个桌面普通文件已记入清单, 不参与恢复)' } else { '' }))
    & $Log ('重要提醒: 请把整个工具文件夹(含备份)复制到 U盘/移动硬盘/网盘, 重装系统后再复制回来使用。')
    return [string]$Dir
}

# ==================== 恢复 ====================

function Invoke-Restore {
    param(
        [string]$Dir = '',
        [string]$Target = '',
        [string]$Old = '',
        [string]$New = '',
        [scriptblock]$Log = { param($m) Write-Host $m },
        [switch]$RestoreLayout
    )
    $Dir = ($Dir -replace '^True\s+', '').Trim()
    $root = Find-BackupRoot -Dir $Dir
    if (-not $root) {
        & $Log '错误: 未找到备份数据(backup_info.json), 请确认备份文件夹位置。'
        return
    }
    & $Log ('使用备份: ' + $root)
    $jsonText = [System.IO.File]::ReadAllText((Join-Path $root 'backup_info.json'))
    $meta = ConvertFrom-Json -InputObject $jsonText
    $items = @($meta.Items)
    if ($items.Count -eq 0) {
        & $Log '错误: 备份内容为空。'
        return
    }
    & $Log ('备份时间: ' + $meta.BackupTime + '  (来自 ' + $meta.Machine + ' / 用户 ' + $meta.UserName + ')')

    $remap = ($Old -and $New -and ($Old -ne $New))
    if ($remap) { & $Log ('已启用路径重映射: "' + $Old + '" -> "' + $New + '"') }

    $sh = New-Object -ComObject WScript.Shell
    $restored = 0
    $failed = @()
    $broken = @()

    foreach ($it in $items) {
        if ($it.Type -eq 'other') { continue }
        if ($Target) {
            $destDir = $Target
        }
        elseif ($it.SourceDir -eq '公共桌面') {
            $destDir = [Environment]::GetFolderPath('CommonDesktopDirectory')
        }
        else {
            $destDir = [Environment]::GetFolderPath('Desktop')
        }
        if (-not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Force -Path $destDir | Out-Null
        }
        $srcFile = Join-Path $root ('shortcuts\' + $it.SourceDir + '\' + $it.Name)
        $destFile = Join-Path $destDir $it.Name
        try {
            if ($it.Type -eq 'url') {
                if (Test-Path -LiteralPath $srcFile) {
                    Copy-Item -LiteralPath $srcFile -Destination $destFile -Force
                }
                elseif ($it.Url) {
                    $content = "[InternetShortcut]`r`nURL=" + $it.Url + "`r`n"
                    [System.IO.File]::WriteAllText($destFile, $content, (New-Object System.Text.UTF8Encoding $false))
                }
                else { continue }
                $restored++
                & $Log ('  [恢复] ' + $it.Name + '  ->  ' + $it.Url)
            }
            else {
                $needRebuild = $false
                if ($remap -and $it.TargetPath) { $needRebuild = $true }
                if (-not $needRebuild -and -not (Test-Path -LiteralPath $srcFile)) { $needRebuild = $true }

                if ($needRebuild -and $it.TargetPath) {
                    $s = $sh.CreateShortcut($destFile)
                    $s.TargetPath = Convert-Path -Path ([string]$it.TargetPath) -Old $Old -New $New
                    if ($it.Arguments) { $s.Arguments = [string]$it.Arguments }
                    if ($it.WorkingDirectory) { $s.WorkingDirectory = Convert-Path -Path ([string]$it.WorkingDirectory) -Old $Old -New $New }
                    if ($it.IconLocation) { $s.IconLocation = Convert-Icon -Icon ([string]$it.IconLocation) -Old $Old -New $New }
                    if ($it.Hotkey) { $s.Hotkey = [string]$it.Hotkey }
                    if ($it.WindowStyle) { $s.WindowStyle = [int]$it.WindowStyle }
                    if ($it.Description) { $s.Description = [string]$it.Description }
                    $s.Save()
                }
                elseif (Test-Path -LiteralPath $srcFile) {
                    Copy-Item -LiteralPath $srcFile -Destination $destFile -Force
                }
                else {
                    & $Log ('  [跳过] ' + $it.Name + ' (备份文件与元数据均缺失)')
                    continue
                }
                $restored++
                & $Log ('  [恢复] ' + $it.Name + '  ->  ' + $it.TargetPath)
            }
        }
        catch {
            $failed += $it.Name
            & $Log ('  [失败] ' + $it.Name + ' - ' + $_.Exception.Message)
            continue
        }

        if (($it.Type -eq 'lnk') -and $it.TargetPath) {
            $chk = Convert-Path -Path ([string]$it.TargetPath) -Old $Old -New $New
            if ($chk -and -not (Test-Path -LiteralPath $chk)) {
                $broken += [pscustomobject]@{ Name = $it.Name; Target = $chk }
            }
        }
    }
    $sh = $null

    # 生成恢复报告
    $report = @()
    $report += '恢复报告 - ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $report += '成功恢复: ' + $restored + ' 个'
    if ($failed.Count -gt 0) { $report += '恢复失败: ' + $failed.Count + ' 个 (' + ($failed -join ', ') + ')' }
    if ($broken.Count -gt 0) {
        $report += ''
        $report += '以下快捷方式的目标程序/文件目前不存在(程序可能尚未重装):'
        foreach ($b in $broken) { $report += '  ' + $b.Name + '  ->  ' + $b.Target }
    }
    else { $report += '所有快捷方式的目标均存在。' }
    $reportPath = Join-Path $root ('恢复报告_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.txt')
    [System.IO.File]::WriteAllLines($reportPath, $report, (New-Object System.Text.UTF8Encoding $true))

    & $Log ('----------------------------------------')
    & $Log ('恢复完成! 成功 ' + $restored + ' 个' + $(if ($failed.Count -gt 0) { ', 失败 ' + $failed.Count + ' 个' } else { '' }))
    if ($broken.Count -gt 0) {
        & $Log ('注意: 有 ' + $broken.Count + ' 个快捷方式的目标程序不存在(可能还没重装), 详见恢复报告。')
        foreach ($b in $broken) { & $Log ('  [目标缺失] ' + $b.Name + '  ->  ' + $b.Target) }
    }
    else {
        & $Log '所有快捷方式的目标均存在, 全部有效!'
    }
    & $Log ('恢复报告已保存: ' + $reportPath)
    if ($RestoreLayout) { [void](Restore-DesktopLayout -Dir $root -Log $Log) }
    return [string]$reportPath
}

# ==================== 检查 ====================

function Invoke-Check {
    param([scriptblock]$Log = { param($m) Write-Host $m })
    $sh = New-Object -ComObject WScript.Shell
    $total = 0
    $broken = @()
    foreach ($d in (Get-DesktopDirs)) {
        foreach ($f in @(Get-ChildItem -LiteralPath $d.Path -File -Filter '*.lnk' -ErrorAction SilentlyContinue)) {
            $total++
            try {
                $s = $sh.CreateShortcut($f.FullName)
                $t = [string]$s.TargetPath
                if ($t -and -not (Test-Path -LiteralPath $t)) {
                    $broken += [pscustomobject]@{ Name = $f.Name; Target = $t }
                    & $Log ('  [失效] ' + $f.Name + '  ->  ' + $t)
                }
            }
            catch {
                & $Log ('  [无法读取] ' + $f.Name)
            }
        }
    }
    $sh = $null
    & $Log ('检查完成: 共 ' + $total + ' 个快捷方式, ' + @($broken).Count + ' 个目标不存在。')
    return @($broken)
}

# ==================== 桌面图标位置 ====================

function Find-DesktopBagKey {
    $k = Get-ChildItem 'HKCU:\Software\Microsoft\Windows\Shell\Bags' -ErrorAction SilentlyContinue |
        Where-Object { Test-Path "$($_.PSPath)\Desktop" } | Select-Object -First 1
    if ($k) { return $k }
    if (Test-Path 'HKCU:\Software\Microsoft\Windows\Shell\Bags\1\Desktop') { return Get-Item 'HKCU:\Software\Microsoft\Windows\Shell\Bags\1\Desktop' }
    return $null
}

function Save-DesktopLayout {
    param(
        [string]$Dir,
        [scriptblock]$Log = { param($m) Write-Host $m }
    )
    $key = Find-DesktopBagKey
    if (-not $key) {
        & $Log '  [位置] 未找到桌面图标布局注册表项, 跳过位置备份。'
        return $false
    }
    $regPath = $key.Name.Replace('HKEY_CURRENT_USER', 'HKCU')
    $out = Join-Path $Dir '桌面图标布局.reg'
    & $Log '  [位置] 正在导出桌面图标布局...'
    try {
        reg export $regPath $out /y | Out-Null
    } catch {
        & $Log ('  [位置] 导出失败: ' + $_.Exception.Message)
        return $false
    }
    if (Test-Path -LiteralPath $out) {
        & $Log ('  [位置] 桌面图标布局已保存: ' + $out)
        return $true
    }
    & $Log '  [位置] 导出失败, 未生成布局文件。'
    return $false
}

function Restore-DesktopLayout {
    param(
        [string]$Dir,
        [scriptblock]$Log = { param($m) Write-Host $m }
    )
    $script:LastLayoutRestored = $false
    $reg = Join-Path $Dir '桌面图标布局.reg'
    if (-not (Test-Path -LiteralPath $reg)) {
        & $Log '  [位置] 备份中没有桌面图标布局文件, 跳过位置恢复。'
        return $false
    }
    & $Log '  [位置] 正在导入桌面图标布局...'
    try {
        reg import $reg | Out-Null
    } catch {
        & $Log ('  [位置] 导入失败: ' + $_.Exception.Message)
        return $false
    }
    & $Log '  [位置] 桌面图标布局已导入, 重启资源管理器后生效。'
    $script:LastLayoutRestored = $true
    return $true
}

# ==================== 图形界面 ====================

function Show-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $fontMain = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $fontBold = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
    $fontBtn = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
    $fontTitle = New-Object System.Drawing.Font('Microsoft YaHei UI', 13.5, [System.Drawing.FontStyle]::Bold)
    $fontSub = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5)
    $fontBadge = New-Object System.Drawing.Font('Microsoft YaHei UI', 8, [System.Drawing.FontStyle]::Bold)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Script:ToolName
    $form.Size = New-Object System.Drawing.Size(800, 770)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.Font = $fontMain
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0D1117')
    $form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F0F6FC')
    $icoPath = Join-Path $Script:BaseDir 'app.ico'
    if (Test-Path -LiteralPath $icoPath) {
        try { $form.Icon = New-Object System.Drawing.Icon($icoPath) } catch { }
    }

    $drawCardBorder = {
        param($sender, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml('#30363D'), 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $sender.Width - 1, $sender.Height - 1)
        $pen.Dispose()
    }

    # ---- 顶部横幅 (Header) ----
    $pnlHead = New-Object System.Windows.Forms.Panel
    $pnlHead.Location = New-Object System.Drawing.Point(0, 0)
    $pnlHead.Size = New-Object System.Drawing.Size(800, 68)
    $pnlHead.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#161B22')
    $pnlHead.Add_Paint({
        param($sender, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml('#21262D'), 1)
        $e.Graphics.DrawLine($pen, 0, $sender.Height - 1, $sender.Width, $sender.Height - 1)
        $pen.Dispose()
    })
    $form.Controls.Add($pnlHead)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = '桌面快捷方式备份与恢复工具'
    $lblTitle.Font = $fontTitle
    $lblTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F0F6FC')
    $lblTitle.AutoSize = $true
    $lblTitle.Location = New-Object System.Drawing.Point(18, 12)
    $pnlHead.Controls.Add($lblTitle)

    $lblBadge = New-Object System.Windows.Forms.Label
    $lblBadge.Text = ' 开源便携版 '
    $lblBadge.Font = $fontBadge
    $lblBadge.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#58A6FF')
    $lblBadge.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#1F2937')
    $lblBadge.AutoSize = $true
    $lblBadge.Location = New-Object System.Drawing.Point(295, 16)
    $lblBadge.Padding = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)
    $pnlHead.Controls.Add($lblBadge)

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = '专为系统重装设计 · 快捷方式完整提取 · 智能路径纠正 · 图标位置保存'
    $lblSub.Font = $fontSub
    $lblSub.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#8B949E')
    $lblSub.AutoSize = $true
    $lblSub.Location = New-Object System.Drawing.Point(19, 41)
    $pnlHead.Controls.Add($lblSub)

    $lblStatusPill = New-Object System.Windows.Forms.Label
    $lblStatusPill.Text = '● 系统已就绪'
    $lblStatusPill.Font = $fontSub
    $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#3FB950')
    $lblStatusPill.AutoSize = $true
    $lblStatusPill.Location = New-Object System.Drawing.Point(680, 24)
    $pnlHead.Controls.Add($lblStatusPill)

    # ---- 卡片 1: 备份目录与历史记录 ----
    $pnlCard1 = New-Object System.Windows.Forms.Panel
    $pnlCard1.Location = New-Object System.Drawing.Point(16, 78)
    $pnlCard1.Size = New-Object System.Drawing.Size(752, 98)
    $pnlCard1.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#161B22')
    $pnlCard1.Add_Paint($drawCardBorder)
    $form.Controls.Add($pnlCard1)

    $lblDir = New-Object System.Windows.Forms.Label
    $lblDir.Text = '📁 备份存储目录 (留空则在工具同级目录自动创建):'
    $lblDir.Location = New-Object System.Drawing.Point(14, 10)
    $lblDir.AutoSize = $true
    $lblDir.Font = $fontBold
    $lblDir.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#C9D1D9')
    $pnlCard1.Controls.Add($lblDir)

    $txtDir = New-Object System.Windows.Forms.TextBox
    $txtDir.Location = New-Object System.Drawing.Point(14, 32)
    $txtDir.Size = New-Object System.Drawing.Size(622, 24)
    $txtDir.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0D1117')
    $txtDir.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F0F6FC')
    $txtDir.BorderStyle = 'FixedSingle'
    $pnlCard1.Controls.Add($txtDir)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = '浏览...'
    $btnBrowse.Location = New-Object System.Drawing.Point(644, 30)
    $btnBrowse.Size = New-Object System.Drawing.Size(94, 27)
    $btnBrowse.FlatStyle = 'Flat'
    $btnBrowse.FlatAppearance.BorderSize = 1
    $btnBrowse.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnBrowse.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnBrowse.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#21262D')
    $btnBrowse.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#C9D1D9')
    $btnBrowse.UseVisualStyleBackColor = $false
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = '选择备份文件夹'
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtDir.Text = $dlg.SelectedPath
        }
    })
    $pnlCard1.Controls.Add($btnBrowse)

    $lblPick = New-Object System.Windows.Forms.Label
    $lblPick.Text = '选择历史备份:'
    $lblPick.Location = New-Object System.Drawing.Point(14, 65)
    $lblPick.AutoSize = $true
    $lblPick.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#8B949E')
    $pnlCard1.Controls.Add($lblPick)

    $cmbBackup = New-Object System.Windows.Forms.ComboBox
    $cmbBackup.Location = New-Object System.Drawing.Point(105, 63)
    $cmbBackup.Size = New-Object System.Drawing.Size(531, 24)
    $cmbBackup.DropDownStyle = 'DropDownList'
    $cmbBackup.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0D1117')
    $cmbBackup.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F0F6FC')
    $cmbBackup.FlatStyle = 'Flat'
    $cmbBackup.Add_SelectedIndexChanged({
        if ($cmbBackup.SelectedIndex -ge 0) { $txtDir.Text = [string]$cmbBackup.SelectedItem }
    })
    $pnlCard1.Controls.Add($cmbBackup)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = '刷新'
    $btnRefresh.Location = New-Object System.Drawing.Point(644, 62)
    $btnRefresh.Size = New-Object System.Drawing.Size(94, 25)
    $btnRefresh.FlatStyle = 'Flat'
    $btnRefresh.FlatAppearance.BorderSize = 1
    $btnRefresh.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnRefresh.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnRefresh.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#21262D')
    $btnRefresh.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#8B949E')
    $btnRefresh.UseVisualStyleBackColor = $false
    $pnlCard1.Controls.Add($btnRefresh)

    $script:RefreshBackupList = {
        $cmbBackup.Items.Clear()
        $backups = @(Get-ChildItem -LiteralPath $Script:BaseDir -Directory -Filter '快捷方式备份_*' -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
        foreach ($b in $backups) { [void]$cmbBackup.Items.Add($b.FullName) }
        if ($backups.Count -gt 0) {
            $cur = ($txtDir.Text.Trim() -replace '^True\s+', '').Trim()
            if (-not $cur -or $cur -match '^True') { 
                $cur = $backups[0].FullName
                $txtDir.Text = $cur 
            }
            $idx = $cmbBackup.Items.IndexOf($cur)
            if ($idx -ge 0) { $cmbBackup.SelectedIndex = $idx }
        }
    }
    $btnRefresh.Add_Click({ & $script:RefreshBackupList })
    & $script:RefreshBackupList

    # ---- 核心操作按钮 (Action Center) ----
    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Text = '💾  1. 备份当前桌面'
    $btnBackup.Font = $fontBtn
    $btnBackup.Location = New-Object System.Drawing.Point(16, 186)
    $btnBackup.Size = New-Object System.Drawing.Size(240, 48)
    $btnBackup.FlatStyle = 'Flat'
    $btnBackup.FlatAppearance.BorderSize = 0
    $btnBackup.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#3B82F6')
    $btnBackup.FlatAppearance.MouseDownBackColor = [System.Drawing.ColorTranslator]::FromHtml('#1D4ED8')
    $btnBackup.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
    $btnBackup.ForeColor = [System.Drawing.Color]::White
    $btnBackup.UseVisualStyleBackColor = $false
    $btnBackup.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Text = '🚀  2. 一键还原桌面'
    $btnRestore.Font = $fontBtn
    $btnRestore.Location = New-Object System.Drawing.Point(272, 186)
    $btnRestore.Size = New-Object System.Drawing.Size(240, 48)
    $btnRestore.FlatStyle = 'Flat'
    $btnRestore.FlatAppearance.BorderSize = 0
    $btnRestore.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#10B981')
    $btnRestore.FlatAppearance.MouseDownBackColor = [System.Drawing.ColorTranslator]::FromHtml('#047857')
    $btnRestore.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#059669')
    $btnRestore.ForeColor = [System.Drawing.Color]::White
    $btnRestore.UseVisualStyleBackColor = $false
    $btnRestore.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($btnRestore)

    $btnCheck = New-Object System.Windows.Forms.Button
    $btnCheck.Text = '🔍  检查失效快捷方式'
    $btnCheck.Font = $fontBtn
    $btnCheck.Location = New-Object System.Drawing.Point(528, 186)
    $btnCheck.Size = New-Object System.Drawing.Size(240, 48)
    $btnCheck.FlatStyle = 'Flat'
    $btnCheck.FlatAppearance.BorderSize = 1
    $btnCheck.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnCheck.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnCheck.FlatAppearance.MouseDownBackColor = [System.Drawing.ColorTranslator]::FromHtml('#161B22')
    $btnCheck.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#21262D')
    $btnCheck.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F0F6FC')
    $btnCheck.UseVisualStyleBackColor = $false
    $btnCheck.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($btnCheck)

    # ---- 卡片 3: 选项与重映射配置 ----
    $pnlCard3 = New-Object System.Windows.Forms.Panel
    $pnlCard3.Location = New-Object System.Drawing.Point(16, 244)
    $pnlCard3.Size = New-Object System.Drawing.Size(752, 94)
    $pnlCard3.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#161B22')
    $pnlCard3.Add_Paint($drawCardBorder)
    $form.Controls.Add($pnlCard3)

    $chkLayout = New-Object System.Windows.Forms.CheckBox
    $chkLayout.Text = '同时备份 / 恢复桌面图标摆放位置 (导入注册表, 恢复后需重启资源管理器生效)'
    $chkLayout.Location = New-Object System.Drawing.Point(14, 10)
    $chkLayout.Size = New-Object System.Drawing.Size(600, 22)
    $chkLayout.Checked = $true
    $chkLayout.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#C9D1D9')
    $pnlCard3.Controls.Add($chkLayout)

    $chkRemap = New-Object System.Windows.Forms.CheckBox
    $chkRemap.Text = '启用路径 / 盘符重映射 (重装系统后用户名或软件安装盘符发生变化时使用)'
    $chkRemap.Location = New-Object System.Drawing.Point(14, 36)
    $chkRemap.Size = New-Object System.Drawing.Size(600, 22)
    $chkRemap.Checked = $false
    $chkRemap.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#C9D1D9')
    $pnlCard3.Controls.Add($chkRemap)

    $lblOld = New-Object System.Windows.Forms.Label
    $lblOld.Text = '旧路径 / 旧名称:'
    $lblOld.Location = New-Object System.Drawing.Point(34, 64)
    $lblOld.AutoSize = $true
    $lblOld.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#8B949E')
    $pnlCard3.Controls.Add($lblOld)

    $txtOld = New-Object System.Windows.Forms.TextBox
    $txtOld.Location = New-Object System.Drawing.Point(135, 61)
    $txtOld.Size = New-Object System.Drawing.Size(240, 23)
    $txtOld.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0D1117')
    $txtOld.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F0F6FC')
    $txtOld.BorderStyle = 'FixedSingle'
    $pnlCard3.Controls.Add($txtOld)

    $lblArrow = New-Object System.Windows.Forms.Label
    $lblArrow.Text = '➔'
    $lblArrow.Location = New-Object System.Drawing.Point(385, 62)
    $lblArrow.AutoSize = $true
    $lblArrow.Font = $fontBold
    $lblArrow.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#58A6FF')
    $pnlCard3.Controls.Add($lblArrow)

    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = '新路径 / 新名称:'
    $lblNew.Location = New-Object System.Drawing.Point(415, 64)
    $lblNew.AutoSize = $true
    $lblNew.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#8B949E')
    $pnlCard3.Controls.Add($lblNew)

    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(515, 61)
    $txtNew.Size = New-Object System.Drawing.Size(223, 23)
    $txtNew.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0D1117')
    $txtNew.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F0F6FC')
    $txtNew.BorderStyle = 'FixedSingle'
    $pnlCard3.Controls.Add($txtNew)

    # ---- 日志区域 ----
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = '📋 运行日志'
    $lblLog.Location = New-Object System.Drawing.Point(18, 348)
    $lblLog.AutoSize = $true
    $lblLog.Font = $fontBold
    $lblLog.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#C9D1D9')
    $form.Controls.Add($lblLog)

    $btnClearLog = New-Object System.Windows.Forms.Button
    $btnClearLog.Text = '清空日志'
    $btnClearLog.Location = New-Object System.Drawing.Point(688, 346)
    $btnClearLog.Size = New-Object System.Drawing.Size(80, 22)
    $btnClearLog.FlatStyle = 'Flat'
    $btnClearLog.FlatAppearance.BorderSize = 1
    $btnClearLog.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnClearLog.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#30363D')
    $btnClearLog.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#21262D')
    $btnClearLog.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#8B949E')
    $btnClearLog.Font = $fontSub
    $btnClearLog.UseVisualStyleBackColor = $false
    $btnClearLog.Add_Click({ if ($script:LogBox) { $script:LogBox.Clear() } })
    $form.Controls.Add($btnClearLog)

    $script:LogBox = New-Object System.Windows.Forms.TextBox
    $script:LogBox.Multiline = $true
    $script:LogBox.ReadOnly = $true
    $script:LogBox.ScrollBars = 'Vertical'
    $script:LogBox.WordWrap = $false
    $logFont = New-Object System.Drawing.Font('Consolas', 9.5)
    try {
        $testFont = New-Object System.Drawing.Font('Cascadia Code', 9.5)
        if ($testFont.Name -eq 'Cascadia Code') { $logFont = $testFont }
    } catch { }
    $script:LogBox.Font = $logFont
    $script:LogBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0D1117')
    $script:LogBox.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#A5D6FF')
    $script:LogBox.BorderStyle = 'FixedSingle'
    $script:LogBox.Location = New-Object System.Drawing.Point(16, 372)
    $script:LogBox.Size = New-Object System.Drawing.Size(752, 345)
    $form.Controls.Add($script:LogBox)

    if ($BackupDir) { $txtDir.Text = $BackupDir }
    if ($OldUser -and $NewUser) { $txtOld.Text = $OldUser; $txtNew.Text = $NewUser; $chkRemap.Checked = $true }

    $script:GLog = {
        param($m)
        if ($script:LogBox) {
            $ts = Get-Date -Format 'HH:mm:ss'
            $script:LogBox.AppendText("[$ts] " + ([string]$m) + "`r`n")
            $script:LogBox.SelectionStart = $script:LogBox.Text.Length
            $script:LogBox.ScrollToCaret()
            $script:LogBox.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $script:LogBox.AppendText('========================================================================' + "`r`n")
    $script:LogBox.AppendText('  桌面快捷方式备份与恢复工具 (开源便携版)' + "`r`n")
    $script:LogBox.AppendText('  - 重装系统前: 点击「1. 备份当前桌面」, 将生成的备份文件夹连同工具拷贝至U盘。' + "`r`n")
    $script:LogBox.AppendText('  - 重装系统后: 在新系统打开工具, 点击「2. 一键还原桌面」即可无损复原。' + "`r`n")
    $script:LogBox.AppendText('========================================================================' + "`r`n`r`n")

    $btnBackup.Add_Click({
        $btnBackup.Enabled = $false; $btnRestore.Enabled = $false; $btnCheck.Enabled = $false
        $lblStatusPill.Text = '● 正在备份...'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E3B341')
        try {
            $cleanDir = ($txtDir.Text.Trim() -replace '^True\s+', '').Trim()
            $dir = (Invoke-Backup -Dir $cleanDir -Log $script:GLog -SaveLayout:$chkLayout.Checked | Select-Object -Last 1)
            $txtDir.Text = [string]$dir
            & $script:RefreshBackupList
            $lblStatusPill.Text = '● 备份完成'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#3FB950')
            [System.Windows.Forms.MessageBox]::Show($form, "备份已成功完成!`n`n已为您保存快捷方式元数据与图标位置。`n重要提示: 请务必将备份文件夹连同本工具复制到 U盘/移动硬盘，以免重装系统格式化丢失!", '备份完成', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            $lblStatusPill.Text = '● 备份出错'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F85149')
            & $script:GLog ('出错了: ' + $_.Exception.Message)
            [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        finally { $btnBackup.Enabled = $true; $btnRestore.Enabled = $true; $btnCheck.Enabled = $true }
    })

    $script:doRestore = {
        $btnBackup.Enabled = $false; $btnRestore.Enabled = $false; $btnCheck.Enabled = $false
        $lblStatusPill.Text = '● 正在恢复...'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E3B341')
        try {
            # 需要管理员权限时, 询问一次并以管理员身份重启工具后自动恢复
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            if (-not $isAdmin) {
                $ans = [System.Windows.Forms.MessageBox]::Show($form, "恢复「公共桌面」等受保护位置需要管理员权限。`n是否立即以管理员身份启动并继续恢复? (将弹出一次系统 UAC 确认)", '需要管理员权限', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { 
                    $lblStatusPill.Text = '● 操作取消'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#8B949E')
                    return 
                }
                $currentExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
                $isExe = ($currentExe -match '\.exe$' -and $currentExe -notmatch 'powershell')
                $d = ($txtDir.Text.Trim() -replace '^True\s+', '').Trim(); $o = $txtOld.Text.Trim(); $n = $txtNew.Text.Trim()
                $extraArgs = "-Action gui -AutoRestore"
                if ($d) { $extraArgs += " -BackupDir `"$d`"" }
                if ($chkRemap.Checked -and $o -and $n) { $extraArgs += " -OldUser `"$o`" -NewUser `"$n`"" }
                try {
                    if ($isExe) {
                        Start-Process -FilePath $currentExe -ArgumentList $extraArgs -Verb RunAs | Out-Null
                    } else {
                        if (-not $Script:ThisPath) { & $script:GLog '错误: 无法确定脚本路径, 无法提升权限。'; return }
                        $elevArgs = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$Script:ThisPath`" $extraArgs"
                        Start-Process -FilePath 'powershell.exe' -ArgumentList $elevArgs -Verb RunAs | Out-Null
                    }
                } catch {
                    & $script:GLog ('提升权限失败或已取消: ' + $_.Exception.Message)
                    [System.Windows.Forms.MessageBox]::Show($form, '提升权限失败或已取消, 未能以管理员身份恢复。', '提示', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                }
                $form.Close()
                return
            }
            $old = ''; $new = ''
            if ($chkRemap.Checked) { $old = $txtOld.Text.Trim(); $new = $txtNew.Text.Trim() }
            $cleanDir = ($txtDir.Text.Trim() -replace '^True\s+', '').Trim()
            Invoke-Restore -Dir $cleanDir -Old $old -New $new -Log $script:GLog -RestoreLayout:$chkLayout.Checked | Out-Null
            $lblStatusPill.Text = '● 恢复完成'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#3FB950')
            if ($chkLayout.Checked -and $script:LastLayoutRestored) {
                $ans = [System.Windows.Forms.MessageBox]::Show($form, "桌面图标位置已成功导入注册表!`n是否立即重启「资源管理器」让图标排列即刻生效?`n(会短暂刷新桌面，不影响已打开的文件)", '重启资源管理器确认', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
                    taskkill /F /IM explorer.exe | Out-Null
                    Start-Sleep -Seconds 2
                    Start-Process explorer.exe
                    & $script:GLog '  [位置] 已重启资源管理器, 桌面图标位置已生效。'
                }
            }
            [System.Windows.Forms.MessageBox]::Show($form, '桌面快捷方式已全部恢复完成! 详细情况请查看运行日志或同级目录的恢复报告。', '恢复成功', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            $lblStatusPill.Text = '● 恢复出错'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F85149')
            & $script:GLog ('出错了: ' + $_.Exception.Message)
            [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        finally { $btnBackup.Enabled = $true; $btnRestore.Enabled = $true; $btnCheck.Enabled = $true }
    }
    $btnRestore.Add_Click($script:doRestore)
    if ($AutoRestore) { $form.Add_Shown({ & $script:doRestore }) }

    $btnCheck.Add_Click({
        $btnBackup.Enabled = $false; $btnRestore.Enabled = $false; $btnCheck.Enabled = $false
        $lblStatusPill.Text = '● 正在检查...'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E3B341')
        try {
            & $script:GLog '开始检查桌面快捷方式有效性...'
            $broken = Invoke-Check -Log $script:GLog
            if (@($broken).Count -gt 0) {
                $lblStatusPill.Text = '● 发现失效项'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F85149')
            } else {
                $lblStatusPill.Text = '● 全部有效'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#3FB950')
            }
        }
        catch {
            $lblStatusPill.Text = '● 检查出错'; $lblStatusPill.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F85149')
            & $script:GLog ('出错了: ' + $_.Exception.Message)
        }
        finally { $btnBackup.Enabled = $true; $btnRestore.Enabled = $true; $btnCheck.Enabled = $true }
    })

    [void]$form.ShowDialog()
}

# ==================== 入口 ====================

switch ($Action) {
    'gui'     { Show-Gui }
    'backup'  { Invoke-Backup -Dir $BackupDir -SaveLayout:$SaveLayout }
    'restore' { Invoke-Restore -Dir $BackupDir -Target $TargetDir -Old $OldUser -New $NewUser -RestoreLayout:$RestoreLayout }
    'check'   { Invoke-Check }
}
