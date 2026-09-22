<#
  Reasonix / dsh / zcode / trae  →  GitHub 每日备份   (v2)
  --------------------------------------------------------------------
  仓库  : git@github.com:TianQue6916/reasonix.git   （public → 强制脱敏）
  结构  : reasonix/  dsh/  zcode-appdata/  zcode-home/  trae-cn/  trae-solo/
  用法  :
    .\backup-to-github.ps1            同步 + 脱敏 + 提交 + 推送
    .\backup-to-github.ps1 -DryRun    只同步+脱敏+报告
    .\backup-to-github.ps1 -NoPush    提交但不推送
    .\backup-to-github.ps1 -Full      全量重扫（默认只处理有变更的文件）
  硬约束:
    * 文件清单取自 index（git add -A 后），避免 porcelain 的目录折叠漏文件
    * 提交前二次全量扫描，发现残留敏感即中止且不提交
    * 永不 force push / 永不改写远端历史
  v2 修正（2026-09-23）: ①文件清单改从 index 取 ②二进制探测（前 4KB 含 NUL 即跳过）
                        ③UTF-8 输出编码（中文文件名）④路径拼接用 [IO.Path]::Combine
#>
[CmdletBinding()]
param(
  [switch]$DryRun,
  [switch]$NoPush,
  [switch]$Full
)
$ErrorActionPreference = 'Continue'
try{
  [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
  $OutputEncoding            = [Text.UTF8Encoding]::new($false)
}catch{}

$RepoDir = Join-Path -Path $PSScriptRoot -ChildPath 'repo'
$RepoUrl = 'git@github.com:TianQue6916/reasonix.git'
$Branch  = 'master'
$LogFile = Join-Path -Path $env:APPDATA -ChildPath 'reasonix\github-backup.log'

function Log([string]$m){
  $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
  Write-Host $line
  try{
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    $fi = Get-Item -LiteralPath $LogFile -ErrorAction SilentlyContinue
    if($fi -and $fi.Length -gt 1MB){ Set-Content -Path $LogFile -Value (Get-Content $LogFile -Tail 800) -Encoding UTF8 }
  }catch{}
}

$AppData = $env:APPDATA
$User    = $env:USERPROFILE
$Targets = @(
  @{ Name='reasonix';      Src="$AppData\reasonix"
     XD=@('global-workspace','projects','desktop-sessions-v5','cache','archive','repair','mcp-state',
          'crash-fatal','crash-pending','transactions','state','remote','diagnostics',
          '文件翻译','文件翻译2','.git')
     XF=@('*.zip','*.lock','*.log') },

  @{ Name='dsh';           Src="$User\.dsh"
     XD=@('gate','_backup-20260920-upgrade','node_modules','.git')
     XF=@('*.lock','*.log','credentials.json','*.sqlite','*.db') },

  @{ Name='zcode-appdata'; Src="$AppData\ZCode"
     XD=@('session','rum-electron-store','.git')
     XF=@() },

  @{ Name='zcode-home';    Src="$User\.zcode"
     XD=@('cli','workspace','plugin-workspace','.git','certs','checkpoints','crash','logs')
     XF=@('*.lock','*.log','credentials.json','*.sqlite','*.db') },

  @{ Name='trae-cn';       Src="$AppData\Trae CN"
     XD=@('logs','Cache','CachedData','CachedProfilesData','CachedExtensionVSIXs','CachedConfigurations',
          'Code Cache','GPUCache','DawnGraphiteCache','DawnWebGPUCache','Crashpad','ModularData','Partitions',
          'Local Storage','Session Storage','WebStorage','IndexedDB','Service Worker','blob_storage','Network',
          'Backups','monitor','.git','workspaceStorage','globalStorage','History')
     XF=@('*.log','*.lock','DIPS*','*.vscdb','*.vscdb.*','credentials.json','*.sqlite','*.db') },

  @{ Name='trae-solo';     Src="$AppData\TRAE SOLO CN"
     XD=@('logs','Cache','CachedData','CachedProfilesData','Code Cache','GPUCache','DawnGraphiteCache',
          'DawnWebGPUCache','Crashpad','ModularData','Partitions','Local Storage','Session Storage',
          'WebStorage','IndexedDB','Service Worker','blob_storage','Network','Backups','monitor',
          'solo-lite','solo-lite-default-workspace','.git','workspaceStorage','globalStorage','History')
     XF=@('*.log','*.lock','DIPS*') }
)

$Scrub = @(
  @{ rx='sk-[A-Za-z0-9_\-]{16,}';                                                    to='sk-REDACTED' },
  @{ rx='(?i)sk-ant-[A-Za-z0-9_\-]{16,}';                                            to='sk-ant-REDACTED' },
  @{ rx='ghp_[A-Za-z0-9]{20,}';                                                      to='ghp_REDACTED' },
  @{ rx='github_pat_[A-Za-z0-9_]{20,}';                                              to='github_pat_REDACTED' },
  @{ rx='gho_[A-Za-z0-9]{20,}';                                                      to='gho_REDACTED' },
  @{ rx='AKIA[0-9A-Z]{16}';                                                          to='AKIA_REDACTED' },
  @{ rx='xox[baprs]-[A-Za-z0-9\-]{10,}';                                             to='xox-REDACTED' },
  @{ rx='(?s)-----BEGIN [A-Z ]*PRIVATE KEY-----.*?-----END [A-Z ]*PRIVATE KEY-----'; to='[PRIVATE-KEY-REDACTED]' },
  @{ rx='(?i)(bearer\s+)[A-Za-z0-9_\-\.]{20,}';                                      to='$1REDACTED' },
  @{ rx='(?im)^(\s*wpa_passphrase\s*=\s*).+$';                                       to='$1REDACTED' }
)
$Detect = @(
  'sk-[A-Za-z0-9_\-]{16,}','ghp_[A-Za-z0-9]{20,}','github_pat_[A-Za-z0-9_]{20,}','gho_[A-Za-z0-9]{20,}',
  'AKIA[0-9A-Z]{16}','xox[baprs]-[A-Za-z0-9\-]{10,}','-----BEGIN [A-Z ]*PRIVATE KEY-----','sk-ant-[A-Za-z0-9_\-]{16,}'
)
$TextExt = @('.json','.jsonl','.md','.txt','.yaml','.yml','.toml','.cfg','.ini','.conf','.properties',
             '.ps1','.psm1','.bat','.cmd','.sh','.py','.js','.mjs','.cjs','.ts','.xml','.csv','.bak','.key','.pem','.env')

function Test-IsTextFile([string]$path){
  try{
    $fi = Get-Item -LiteralPath $path -ErrorAction Stop
    if($fi.Length -gt 40MB){ return $false }
    $ext = ([IO.Path]::GetExtension($fi.Name)).ToLower()
    $nm  = $fi.Name.ToLower()
    $ok  = ($TextExt -contains $ext) -or ($nm -like '.env*') -or ($nm -like '*.jsonl')
    if(-not $ok){ return $false }
    $fs = [IO.File]::OpenRead($path)
    try{
      $n   = [Math]::Min(4096, [int]$fi.Length)
      $buf = New-Object byte[] $n
      $null = $fs.Read($buf, 0, $n)
      foreach($b in $buf){ if($b -eq 0){ return $false } }
    } finally { $fs.Close() }
    return $true
  }catch{ return $false }
}

Log '=== backup-to-github 开始 ==='

# 1. 准备仓库
if(-not (Test-Path (Join-Path -Path $RepoDir -ChildPath '.git'))){
  Log "首次克隆仓库 → $RepoDir"
  git clone $RepoUrl $RepoDir 2>&1 | Out-Null
  if(-not (Test-Path (Join-Path -Path $RepoDir -ChildPath '.git'))){ Log '克隆失败'; exit 2 }
} else {
  git -C $RepoDir pull --ff-only 2>&1 | Out-Null
}

# 2. 同步（robocopy /E，不删除目标多余文件）
foreach($t in $Targets){
  if(-not (Test-Path -LiteralPath $t.Src)){ Log ("跳过（源不存在）: {0}" -f $t.Name); continue }
  $dst = [IO.Path]::Combine($RepoDir, $t.Name)
  New-Item -ItemType Directory -Force -Path $dst | Out-Null
  $ro = @($t.Src, $dst, '/E', '/NFL', '/NDL', '/NJH', '/NJS', '/R:1', '/W:1')
  if($t.XD.Count -gt 0){ $ro += '/XD'; $ro += $t.XD }
  if($t.XF.Count -gt 0){ $ro += '/XF'; $ro += $t.XF }
  & robocopy @ro | Out-Null
  $rc = $LASTEXITCODE
  if($rc -ge 8){ Log ("robocopy 失败 rc={0}: {1}" -f $rc, $t.Name) } else { Log ("同步 {0} (rc={1})" -f $t.Name, $rc) }
}

# 3. 文件清单：先 add 再从 index 取（可穿透未跟踪目录）
git -C $RepoDir add -A 2>$null
$files = @(git -c core.quotepath=false -C $RepoDir diff --cached --name-only 2>$null)
if($Full){
  $files = @(Get-ChildItem -LiteralPath $RepoDir -Recurse -File -Force -ErrorAction SilentlyContinue |
             Where-Object { $_.FullName -notmatch '\\\.git\\' } |
             ForEach-Object { $_.FullName.Substring($RepoDir.Length + 1) })
  Log ("-Full: 全量 {0} 个文件" -f $files.Count)
} else {
  Log ("待处理文件 {0} 个" -f $files.Count)
}

# 4. 脱敏
$scanned = 0; $scrubbed = 0
foreach($rel in $files){
  $fullPath = [IO.Path]::Combine($RepoDir, $rel)
  if(-not (Test-Path -LiteralPath $fullPath -PathType Leaf)){ continue }
  if(-not (Test-IsTextFile $fullPath)){ continue }
  $scanned++
  try{
    $txt  = [IO.File]::ReadAllText($fullPath)
    $orig = $txt
    foreach($s in $Scrub){ $txt = [regex]::Replace($txt, $s.rx, $s.to) }
    if($txt -ne $orig){ [IO.File]::WriteAllText($fullPath, $txt, [Text.UTF8Encoding]::new($false)); $scrubbed++ }
  }catch{ Log ("  脱敏失败: " + $rel) }
}
Log ("脱敏：扫描 {0} 个文本文件，改写 {1} 个" -f $scanned, $scrubbed)

# 5. 提交前兜底校验
git -C $RepoDir add -A 2>$null
$staged = @(git -c core.quotepath=false -C $RepoDir diff --cached --name-only 2>$null)
$bad = @()
foreach($rel in $staged){
  $fullPath = [IO.Path]::Combine($RepoDir, $rel)
  if(-not (Test-Path -LiteralPath $fullPath -PathType Leaf)){ continue }
  if(-not (Test-IsTextFile $fullPath)){ continue }
  try{
    $txt = [IO.File]::ReadAllText($fullPath)
    $probe = $txt.Replace("REDACTED-BY-BACKUP","").Replace("sk-REDACTED","").Replace("sk-ant-REDACTED","")
    foreach($d in $Detect){ if($probe -match $d){ $bad += ("$rel  [匹配 " + $d + "]"); break } }
  }catch{}
}
if($bad.Count -gt 0){
  Log ("!! {0} 个文件仍含疑似密钥 → 中止提交，不推送" -f $bad.Count)
  $bad | Select-Object -First 20 | ForEach-Object { Log ("   SUSPECT: " + $_) }
  exit 3
}

# 5b. 字节级全扫（不限扩展名，抓白名单外的二进制/未知类型）
$bad2 = @()
foreach($rel in $staged){
  $fp = [IO.Path]::Combine($RepoDir, $rel)
  if(-not (Test-Path -LiteralPath $fp -PathType Leaf)){ continue }
  $fi2 = Get-Item -LiteralPath $fp -ErrorAction SilentlyContinue
  if(-not $fi2 -or $fi2.Length -gt 40MB){ continue }
  try{
    $s2 = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($fp))
    $p2 = $s2.Replace("REDACTED-BY-BACKUP","").Replace("sk-REDACTED","").Replace("sk-ant-REDACTED","")
    foreach($d in $Detect){ if($p2 -match $d){ $bad2 += $rel; break } }
  }catch{}
  $s2 = $null; $p2 = $null
}
if($bad2.Count -gt 0){
  Log ("!! 字节级全扫发现 {0} 个文件含密钥（白名单外类型）→ 中止，不推送" -f $bad2.Count)
  $bad2 | Select-Object -First 20 | ForEach-Object { Log ("   SUSPECT-BIN: " + $_) }
  exit 3
}
Log "字节级全扫通过"
Log '兜底校验通过：无残留密钥'
if($staged.Count -eq 0){ Log '无变更，结束'; exit 0 }
Log ("将提交 {0} 个文件" -f $staged.Count)
if($DryRun){ Log '-DryRun：未提交、未推送'; exit 0 }

# 6. 提交 + 推送
$msg = 'backup: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm')
git -C $RepoDir -c user.name='Reasonix Backup' -c user.email='backup@reasonix.local' commit -q -m $msg 2>&1 | Out-Null
if($LASTEXITCODE -ne 0){ Log 'commit 失败'; exit 4 }
Log ("已提交: " + $msg)
if($NoPush){ Log '-NoPush：跳过推送'; exit 0 }
$pr = git -C $RepoDir push origin $Branch 2>&1
if($LASTEXITCODE -ne 0){ Log ("push 失败: " + ($pr -join ' ')); exit 5 }
Log '已推送到 GitHub  === 完成 ==='
