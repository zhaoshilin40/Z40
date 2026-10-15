#Requires -Version 5.1
<#
.SYNOPSIS
    把本仓库里的 skill 安装（联接）到 DSH 技能目录。

.DESCRIPTION
    对 skills\ 下的每个子目录，在 %USERPROFILE%\.dsh\skills\ 建一个目录联接
    （junction），指向仓库里的真实目录。文件不复制，因此改一处两边同步。

    已存在且指向本仓库的联接会跳过；指向别处的会先备份成 <名字>.bak 再重建。

.PARAMETER DshSkillsDir
    目标技能目录。默认 %USERPROFILE%\.dsh\skills

.PARAMETER Force
    即使联接已存在且正确，也删掉重建。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1

.EXAMPLE
    .\install.ps1 -Force
#>
[CmdletBinding()]
param(
    [string] $DshSkillsDir = (Join-Path $env:USERPROFILE '.dsh\skills'),
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

$repoRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Join-Path $repoRoot 'skills'

function Write-Step($msg) { Write-Host "  $msg" }
function Write-Ok($msg)   { Write-Host "  [OK]   $msg" -ForegroundColor Green }
function Write-Skip($msg) { Write-Host "  [SKIP] $msg" -ForegroundColor DarkGray }
function Write-Warn2($msg){ Write-Host "  [WARN] $msg" -ForegroundColor Yellow }

# 删除联接本体。Windows PowerShell 5.1 的 Remove-Item 对 junction 会抛
# NullReferenceException，用 .NET 的 Directory.Delete 才不会误删目标内容。
function Remove-Link([string] $path) {
    [System.IO.Directory]::Delete($path, $false)
}

# 删除一棵真实的目录树（用于清理 .bak 备份）。
function Remove-Tree([string] $path) {
    [System.IO.Directory]::Delete($path, $true)
}

Write-Host ''
Write-Host 'Z40 skill installer' -ForegroundColor Cyan
Write-Host '-------------------'
Write-Host "  repo   : $repoRoot"
Write-Host "  target : $DshSkillsDir"
Write-Host ''

if (-not (Test-Path -LiteralPath $sourceDir)) {
    throw "找不到 skills 目录：$sourceDir（请从仓库根目录运行本脚本）"
}

# ---- 确认目标目录存在 ----
if (-not (Test-Path -LiteralPath $DshSkillsDir)) {
    New-Item -ItemType Directory -Path $DshSkillsDir -Force | Out-Null
    Write-Ok "已创建技能目录 $DshSkillsDir"
}

# ---- 逐个子目录安装 ----
$skills = Get-ChildItem -LiteralPath $sourceDir -Directory |
          Where-Object { $_.Name -notlike '.*' } | Sort-Object Name

if (-not $skills) {
    Write-Warn2 'skills 目录下没有找到任何 skill'
    exit 0
}

$installed = 0
$skipped   = 0

foreach ($skill in $skills) {
    $name   = $skill.Name
    $target = $skill.FullName
    $link   = Join-Path $DshSkillsDir $name

    # 没有 SKILL.md 的目录不是 skill
    if (-not (Test-Path -LiteralPath (Join-Path $target 'SKILL.md'))) {
        Write-Warn2 "$name 缺少 SKILL.md，跳过"
        continue
    }

    $existing = Get-Item -LiteralPath $link -Force -ErrorAction SilentlyContinue

    if ($existing) {
        $isLink = $existing.LinkType -eq 'Junction' -or $existing.LinkType -eq 'SymbolicLink'
        $sameTarget = $isLink -and ($existing.Target -contains $target -or $existing.Target -eq $target)

        if ($sameTarget -and -not $Force) {
            Write-Skip "$name 已安装（联接指向本仓库）"
            $skipped++
            continue
        }

        # 备份并移除旧条目
        $backup = "$link.bak"
        if (Test-Path -LiteralPath $backup) { Remove-Tree $backup }
        if ($isLink) {
            # 只删联接本体，绝不递归（递归会删掉目标目录里的真实文件）。
            # PowerShell 5.1 的 Remove-Item 删 junction 会抛 NullReferenceException，
            # 所以走 .NET。
            Remove-Link $link
        }
        else {
            # 实体目录：改名备份，绝不直接删除用户文件
            Rename-Item -LiteralPath $link -NewName (Split-Path -Leaf $backup)
            Write-Warn2 "$name 原有实体目录已备份为 $(Split-Path -Leaf $backup)"
        }
    }

    New-Item -ItemType Junction -Path $link -Target $target | Out-Null
    Write-Ok "$name -> $target"
    $installed++
}

# ---- 校验 ----
Write-Host ''
Write-Host '安装结果：' -ForegroundColor Cyan
Get-ChildItem -LiteralPath $DshSkillsDir -Force | Sort-Object Name | ForEach-Object {
    $t = if ($_.LinkType) { "$($_.LinkType) -> $($_.Target)" }
         elseif ($_.PSIsContainer) { '实体目录' }
         else { '文件' }
    Write-Host ("  {0,-20} {1}" -f $_.Name, $t)
}

Write-Host ''
Write-Host "新装 $installed 个，跳过 $skipped 个。" -ForegroundColor Cyan
Write-Host '下一步：重启 DSH Desktop，让技能目录重新扫描。' -ForegroundColor Yellow
Write-Host ''
