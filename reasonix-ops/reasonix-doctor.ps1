<#
  Reasonix 工作区体检 / Merge-Back 收尾医师   (rx-doctor v2)
  --------------------------------------------------------------------
  设计原则（2026-09-22 纠偏，与 Reasonix 原生一致）：
    只排除『可再生的依赖 / 缓存 / 运行时产物』；你的数据全部纳入版本控制。
    理由：worktree 隔离的副本来自 git 索引；数据不提交 = 新会话在独立工作区里看不到这些文件。
    「排除掉未提交项」只能买到 porcelain=0 的假象，会破坏隔离副本的完整性。
  用法：
    .\reasonix-doctor.ps1                               # 只诊断，零写入
    .\reasonix-doctor.ps1 -Heal                         # 恢复原生排除规则（安全、幂等）
    .\reasonix-doctor.ps1 -Heal -CommitBaseline         # 再提交数据基线（让 checkout 干净 + 副本完整）
    .\reasonix-doctor.ps1 -Heal -IncludeTaskLogs        # 额外清理 worktree 任务日志（备份后删除）
    .\reasonix-doctor.ps1 -Heal -ExcludeAll             # 旧行为（排除一切）—— 仅在明确只要 porcelain=0 时用
  硬约束：不 push / 不改 config.toml / 不重启应用 / 不删用户数据文件；写前先备份。
#>
[CmdletBinding()]
param(
  [switch]$Heal,
  [switch]$CommitBaseline,
  [switch]$ExcludeAll,
  [switch]$IncludeTaskLogs,
  [string]$Workspace  = '',
  [string]$BackupRoot = ''
)
$ErrorActionPreference = 'Continue'
if(-not $Workspace){ $Workspace = Join-Path $env:APPDATA 'reasonix\global-workspace' }
if(-not $BackupRoot){
  $here = $PSScriptRoot
  if(-not $here){ $here = Split-Path -Parent $MyInvocation.MyCommand.Path }
  if(-not $here){ $here = (Get-Location).Path }
  $BackupRoot = Join-Path $here 'backup'
}
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$bak   = Join-Path $BackupRoot $stamp

function Head($t){ Write-Host ''; Write-Host ('-' * 66) -ForegroundColor DarkGray; Write-Host "  $t" -ForegroundColor Cyan }
function Ok($t)  { Write-Host "  [ OK ]  $t" -ForegroundColor Green }
function Warn($t){ Write-Host "  [WARN]  $t" -ForegroundColor Yellow }
function Bad($t) { Write-Host "  [BLOCK] $t" -ForegroundColor Red }
function Info($t){ Write-Host "          $t" -ForegroundColor Gray }
function GPorcelain([string]$p){ @(git -c core.quotepath=false -C $p status --porcelain 2>$null) }
function Git1([string]$p,[string[]]$a){ ((& git -C $p @a 2>$null) | Out-String).Trim() }

Write-Host ''
Write-Host '  Reasonix 工作区体检 / Merge-Back 收尾医师' -ForegroundColor White
Write-Host ("  时间 {0}   模式 {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if($Heal){'HEAL'}else{'DIAGNOSE(零写入)'})) -ForegroundColor White

Head '0. 环境'
if(-not (Test-Path $Workspace)){ Bad "工作区不存在: $Workspace"; exit 2 }
Info "工作区 : $Workspace"
$procs = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match 'reasonix' })
Info ("进程   : " + (($procs | Group-Object ProcessName | ForEach-Object { "{0}x{1}" -f $_.Name,$_.Count }) -join '  '))
$leaseDir = Join-Path $env:LOCALAPPDATA 'reasonix\workspace-leases'
if(Test-Path $leaseDir){
  $lf = @(Get-ChildItem $leaseDir -File -ErrorAction SilentlyContinue)
  $oldest = ($lf | Sort-Object LastWriteTime | Select-Object -First 1).LastWriteTime
  Info ("锁目录 : {0}  文件 {1} 个（最早 {2}）" -f $leaseDir, $lf.Count, $oldest)
}
if((Git1 $Workspace @('rev-parse','--is-inside-work-tree')) -ne 'true'){
  Warn '该目录不是 git 仓库 -> Reasonix 走「共享工作区」模式：无 worktree 隔离、无 Merge-Back'
  exit 0
}
Info 'git    : 是 git 仓库（Reasonix 会启用 worktree 隔离 + Merge-Back）'
$srcBranch = Git1 $Workspace @('rev-parse','--abbrev-ref','HEAD')

Head '1. worktree 全景'
$blocks = ((git -C $Workspace worktree list --porcelain 2>$null) -join "`n") -split "`n`n"
$wts = @()
foreach($b in $blocks){
  if(-not $b.Trim()){ continue }
  $lines = $b -split "`n"
  $p  = ($lines | Where-Object { $_ -like 'worktree *' }) -replace '^worktree\s+',''
  $br = ($lines | Where-Object { $_ -like 'branch *' })   -replace '^branch\s+refs/heads/',''
  if(-not $br){ $br = '(detached)' }
  $wts += [pscustomobject]@{ Path = $p; Branch = $br; IsSource = ((($p -replace '/','\').TrimEnd('\')) -ieq (($Workspace -replace '/','\').TrimEnd('\'))) }
}
$srcDirty = (GPorcelain $Workspace).Count
Info ("source : {0}" -f $Workspace)
Info ("         分支 {0}  未提交 {1} 条" -f $srcBranch, $srcDirty)
foreach($wt in ($wts | Where-Object { -not $_.IsSource })){
  if(-not (Test-Path $wt.Path)){ Warn "worktree 路径不存在: $($wt.Path)"; continue }
  $d  = (GPorcelain $wt.Path).Count
  $ab = (Git1 $Workspace @('rev-list','--left-right','--count',($wt.Branch + '...' + $srcBranch)))
  Info ("wt     : {0}" -f ($wt.Path -replace '/','\'))
  Info ("         分支 {0}  未提交 {1} 条  领先/落后 {2}" -f $wt.Branch, $d, $ab)
}

Head '2. Merge-Back 阻断项（与界面文案一一对应）'
$wtDirty = 0; $lockCount = 0; $logFiles = 0
foreach($wt in ($wts | Where-Object { -not $_.IsSource })){
  if(-not (Test-Path $wt.Path)){ continue }
  $wtDirty += (GPorcelain $wt.Path).Count
  $tasks = Join-Path $wt.Path '.reasonix\tasks'
  if(Test-Path $tasks){
    $lockCount += @(Get-ChildItem $tasks -Recurse -Filter task.lock -ErrorAction SilentlyContinue).Count
    $logFiles  += @(Get-ChildItem $tasks -Recurse -File      -ErrorAction SilentlyContinue).Count
  }
}
if($wtDirty -gt 0){ Bad ("worktree has uncommitted changes -> {0} 条" -f $wtDirty) } else { Ok 'worktree has uncommitted changes -> 0 条' }
if($srcDirty -gt 0){ Bad ("the recorded source checkout has uncommitted changes -> {0} 条" -f $srcDirty) } else { Ok 'the recorded source checkout has uncommitted changes -> 0 条' }
if($lockCount -gt 0){ Warn ("a source or worktree runtime still has active or waiting work -> task.lock x {0}（会话在跑时属正常）" -f $lockCount) } else { Ok 'active or waiting work -> 未发现 task.lock' }
if($logFiles -gt 0){ Warn ("files would be preserved -> worktree 任务日志 {0} 个（ignored 也算阻止清理）" -f $logFiles) } else { Ok 'files would be preserved -> 无' }

Head '3. 结论与处置'
if((($wtDirty + $srcDirty) -eq 0) -and $logFiles -eq 0){ Ok '未发现会阻止 Merge-Back 收尾的项；若界面仍报错，请把原文文案贴出比对。' }
else {
  if(($wtDirty + $srcDirty) -gt 0){ Info '未提交项 -> 正确解法是【提交】: -Heal -CommitBaseline（checkout 干净 + worktree 副本完整）' }
  Info '[!] 不要用「排除掉未提交项」去换 porcelain=0：worktree 副本来自 git 索引，排除 = 新会话在隔离工作区里看不到这些文件。'
  if($logFiles -gt 0){ Info '任务日志 -> -Heal -IncludeTaskLogs（备份后删除）' }
  Info '界面上的「自动提交工作区未提交修改」与 baseline snapshot 是同一机制，方向正确。'
  Info '第三条(active work)只能靠结束会话解决：合并会切换/删除 checkout，会话在跑时本就该被拦。'
}

Head '4. 修复'
if(-not $Heal){ Info '当前为诊断模式，未做任何写入。加 -Heal 执行可逆修复。'; Write-Host ''; exit 0 }
New-Item -ItemType Directory -Force -Path $bak | Out-Null
Info "备份目录: $bak"
Copy-Item (Join-Path $Workspace '.git\index')        (Join-Path $bak 'index')   -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $Workspace '.git\info\exclude') (Join-Path $bak 'exclude') -Force -ErrorAction SilentlyContinue
$ex  = Join-Path $Workspace '.git\info\exclude'
if($ExcludeAll){
  Warn '使用 -ExcludeAll（旧行为）：会破坏 worktree 副本完整性，仅在明确只要 porcelain=0 时使用'
  $all = Get-ChildItem $Workspace -Force | Where-Object { $_.Name -ne '.git' }
  $rules = @('', ("# --- rx-doctor ExcludeAll " + $stamp + " ---"), '.reasonix/', '*.conflict.*.bak', '*.bak-*', 'hostapd.conf', 'reasonix.toml')
  foreach($i in $all){ if($i.PSIsContainer){ $rules += ($i.Name + '/') } else { $rules += $i.Name } }
} else {
  $rules = @('',
    "# Reasonix 自动提交 —— 本地排除规则（rx-doctor 重建 " + $stamp + "）",
    '# 原则：只排除『可再生的依赖 / 缓存 / 运行时产物』；你的数据全部纳入版本控制。',
    '# 这样『创建隔离工作区』得到的副本才是完整的，才能多会话并行写入。',
    '# 想让某条重新被跟踪：删掉对应行，或 git add -f <path>',
    '',
    '.git/',
    '.git.disabled-*/',
    '.mitm-venv/',
    '.venv/',
    'venv/',
    'env/',
    'node_modules/',
    '__pycache__/',
    '*.pyc',
    '*.pyo',
    '.pytest_cache/',
    '.mypy_cache/',
    '.ruff_cache/',
    '.reasonix/tasks/')
}
Set-Content -Path $ex -Value ($rules -join "`r`n") -Encoding UTF8
Info ("排除规则已写入 {0} 条" -f $rules.Count)
git -C $Workspace reset -q 2>$null
Info 'git reset -q（取消全部暂存）'
if($CommitBaseline){
  git -C $Workspace add -A 2>$null
  $n = @(git -C $Workspace diff --cached --name-only 2>$null).Count
  if($n -gt 0){
    git -C $Workspace -c user.name="Reasonix Baseline" -c user.email="baseline@reasonix.local" commit -q -m ("chore: Reasonix baseline snapshot (" + (Get-Date -Format 'yyyy-MM-dd') + ")") 2>$null
    Info ("已提交基线：{0} 个文件" -f $n)
  } else { Info '无待提交内容，跳过基线提交' }
}
if($IncludeTaskLogs){
  $tlog = Join-Path $bak 'task-logs'
  New-Item -ItemType Directory -Force -Path $tlog | Out-Null
  foreach($wt in ($wts | Where-Object { -not $_.IsSource })){
    $t = Join-Path $wt.Path '.reasonix\tasks'
    if(-not (Test-Path $t)){ continue }
    Copy-Item (Join-Path $t '*') $tlog -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem $t -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '--bash-\d+$' } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Info ("已备份并清理任务日志: {0}" -f $t)
  }
}
$afterSrc = (GPorcelain $Workspace).Count
$afterWt  = 0
foreach($wt in ($wts | Where-Object { -not $_.IsSource })){ if(Test-Path $wt.Path){ $afterWt += (GPorcelain $wt.Path).Count } }
Write-Host ''
if($afterSrc -eq 0 -and $afterWt -eq 0){ Ok ("修复后：source {0} 条 / worktree {1} 条 —— 两条未提交阻断项已归零" -f $afterSrc, $afterWt) }
else { Warn ("修复后仍残留：source {0} 条 / worktree {1} 条" -f $afterSrc, $afterWt) }
Info '回滚：把备份目录里的 index / exclude 覆盖回 .git\ 即可。'
Write-Host ''
