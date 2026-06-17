# TOOLS.md - Local Notes

Skills define *how* tools work. This file is for *your* specifics — the stuff that's unique to your setup.

## What Goes Here

Things like:
- Camera names and locations
- SSH hosts and aliases  
- Preferred voices for TTS
- Speaker/room names
- Device nicknames
- Anything environment-specific

## Examples

```markdown
### Cameras
- living-room → Main area, 180° wide angle
- front-door → Entrance, motion-triggered

### SSH
- home-server → 192.168.1.100, user: admin

### TTS
- Preferred voice: "Nova" (warm, slightly British)
- Default speaker: Kitchen HomePod
```

## Why Separate?

Skills are shared. Your setup is yours. Keeping them apart means you can update skills without losing your notes, and share skills without leaking your infrastructure.

---

## PowerShell Helpers

### `Open-OfficeFile` — 智能打开 Office 文件

**解决的问题**：`Start-Process` 打开大 `.xlsm`（含宏 + ActiveX）会闪退（1.4MB 文件 + 7MB sheet + 6MB drawing，进程存活 2.5 秒后退出）。手动双击能成，但调用 `ShellExecute` 的 `Start-Process` 走的是另一条 Shell 路径，反而不行。

**核心策略**：
- 大文件 + 宏/ActiveX → **PowerShell COM 方式**（先稳起 Excel，再 `Workbooks.Open`）
- 普通文件 → `Start-Process`（最快）
- 失败 → 返回提示，让用户手动双击

**使用方法**：把下面的函数粘到 PowerShell profile 里（`$PROFILE`），以后所有打开 Office 文件都走 `Open-OfficeFile`。

```powershell
# ============================================================
# Open-OfficeFile - 智能打开 Office 文件
# ============================================================
# 用法:
#   Open-OfficeFile "C:\path\to\file.xlsx"            # 自动选方式
#   Open-OfficeFile "C:\path\to\file.xlsx" -ForceCom  # 强制 COM
#   Open-OfficeFile "C:\path\to\file.xlsx" -ForceShell # 强制 Start-Process
# ============================================================
function Open-OfficeFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path,

        [switch]$ForceCom,
        [switch]$ForceShell
    )

    if (-not (Test-Path $Path)) {
        Write-Error "❌ 文件不存在: $Path"
        return
    }

    $ext = [System.IO.Path]::GetExtension($Path).ToLower()
    $resolvedPath = (Resolve-Path $Path).Path
    $fileSize = (Get-Item $Path).Length

    # 触发 COM 方式的条件: Office 类文件 + (大文件 OR 宏格式)
    $macroExts = @('.xlsm', '.xlsb', '.xltm', '.dotm', '.pptm', '.potm')
    $officeExts = @('.xlsx', '.xlsm', '.xlsb', '.xls', '.csv', '.docx', '.doc', '.dotx', '.pptx', '.ppt') + $macroExts
    $isLarge = $fileSize -gt 1MB
    $isMacroFmt = $ext -in $macroExts
    $useCom = $ForceCom -or ($ext -in $officeExts -and ($isLarge -or $isMacroFmt))
    if ($ForceShell) { $useCom = $false }

    Write-Host "📄 $Path"
    Write-Host "   格式: $ext | 大小: $([math]::Round($fileSize/1MB, 2)) MB"

    if ($useCom) {
        Write-Host "   方式: COM（智能检测: 大文件/宏格式）"
        try {
            # 优先复用已运行的 Excel/Word/PPT 进程
            $appMap = @{
                '.xlsx' = 'Excel.Application'; '.xlsm' = 'Excel.Application'; '.xlsb' = 'Excel.Application'
                '.xls'  = 'Excel.Application'; '.csv'  = 'Excel.Application'
                '.docx' = 'Word.Application';  '.doc'  = 'Word.Application';  '.dotx' = 'Word.Application'
                '.pptx' = 'PowerPoint.Application'; '.ppt' = 'PowerPoint.Application'
            }
            $progId = $appMap[$ext]
            $app = $null
            if ($progId) {
                try {
                    $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
                    Write-Host "   复用已运行的 $progId 进程"
                } catch {
                    $app = New-Object -ComObject $progId
                    Write-Host "   新建 $progId 进程"
                }
                $app.Visible = $true
                if ($ext -like '.xl*' -or $ext -eq '.csv') {
                    $app.Workbooks.Open($resolvedPath) | Out-Null
                } elseif ($ext -like '.do*') {
                    $app.Documents.Open($resolvedPath) | Out-Null
                } elseif ($ext -like '.pp*') {
                    $app.Presentations.Open($resolvedPath) | Out-Null
                }
                Write-Host "✅ COM 打开成功" -ForegroundColor Green
                return
            }
        } catch {
            Write-Warning "⚠️ COM 方式失败: $_"
            Write-Host "   回落到 Start-Process..."
        }
    }

    if (-not $ForceCom) {
        Write-Host "   方式: Start-Process"
        try {
            Start-Process $resolvedPath
            Write-Host "✅ Start-Process 启动成功" -ForegroundColor Green
        } catch {
            Write-Warning "⚠️ Start-Process 失败: $_"
            Write-Host ""
            Write-Host "   👉 请手动在资源管理器里双击打开:" -ForegroundColor Yellow
            Write-Host "   $resolvedPath" -ForegroundColor Yellow
        }
    } else {
        Write-Host ""
        Write-Host "   👉 请手动在资源管理器里双击打开:" -ForegroundColor Yellow
        Write-Host "   $resolvedPath" -ForegroundColor Yellow
    }
}
```

**安装位置**：PowerShell profile (`$PROFILE`)，一般是 `C:\Users\wuxin\Documents\PowerShell\Microsoft.PowerShell_profile.ps1`，新窗口生效。

**常见使用**：

```powershell
# 打开 D0409 详设 (普通 xlsx，走 Start-Process 即可)
Open-OfficeFile "D:\work\收保.NETFramework升级\详细设计\詳細設計\1_機能詳細設計\1_画面設計\詳細設計_画面設計書_D0409_保険料等収納返還管理_G3_V02.04.xlsx"

# 打开进度表（大 xlsm + 宏 + ActiveX，自动走 COM）
Open-OfficeFile "D:\work\收保.NETFramework升级\进度表\【収入保険.NETFramework】武漢側開発スケジュール.xlsm"
```

**踩坑记录**：2026-06-15/16，`D:\work\收保.NETFramework升级\进度表\【収入保険.NETFramework】武漢側開発スケジュール.xlsm` 用 `Start-Process` 闪退，改用 COM 方式稳定。

---

Add whatever helps you do your job. This is your cheat sheet.
