<#
  Reasonix workspace autocommit
  ------------------------------------------------------------------
  Keeps git HEAD in sync with the working tree, so that Reasonix
  "Create isolated worktree" (App.CreateIsolatedWorktree) always
  produces a COMPLETE copy -> multiple sessions can write in parallel.

  Workspaces are AUTO-DISCOVERED: every direct child folder of the
  Reasonix home that contains .git  (plus global-workspace).

  Safe by design:
    - single instance (lock file, 10 min stale guard)
    - skips while a merge / rebase / cherry-pick is in progress
    - no-op (fast exit) when the tree is clean
    - never deletes, resets or checks out anything

  Log:   %APPDATA%\reasonix\autocommit.log
  Undo:  schtasks /Delete /TN "ReasonixWorkspaceAutocommit" /F
#>
$ErrorActionPreference = 'SilentlyContinue'

$Home2 = 'C:\Users\27063\AppData\Roaming\reasonix'
$Log   = Join-Path $Home2 'autocommit.log'
$Lock  = Join-Path $Home2 'autocommit.lock'
$AuthorName  = 'Reasonix Autosave'
$AuthorEmail = 'autosave@reasonix.local'

# ---- discover workspaces (a workspace = folder containing .git) ----
$Workspaces = @()
$cands = @( (Join-Path $Home2 'global-workspace') )
foreach ($d in (Get-ChildItem -Path $Home2 -Directory -Force -ErrorAction SilentlyContinue)) {
  $cands += $d.FullName
}
foreach ($c in $cands) {
  if (Test-Path (Join-Path $c '.git')) { $Workspaces += $c }
}
$Workspaces = @($Workspaces | Sort-Object -Unique)

function Write-Log([string]$m) {
  $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
  Add-Content -Path $Log -Value $line -Encoding UTF8
  $fi = Get-Item $Log -ErrorAction SilentlyContinue
  if ($fi -and $fi.Length -gt 512KB) {
    $tail = Get-Content $Log -Tail 400
    Set-Content -Path $Log -Value $tail -Encoding UTF8
  }
}

# ---- single instance ----
if (Test-Path $Lock) {
  $age = (Get-Date) - (Get-Item $Lock).LastWriteTime
  if ($age.TotalMinutes -lt 10) { exit 0 }
}
New-Item -ItemType File -Force -Path $Lock | Out-Null

try {
  if ($Workspaces.Count -eq 0) { Write-Log 'no git workspace found' }

  foreach ($Ws in $Workspaces) {
    $gitdir = (& git -C $Ws rev-parse --absolute-git-dir 2>$null)
    if (-not $gitdir) { Write-Log "skip : not a git repo  [$Ws]"; continue }

    $blocked = $false
    foreach ($f in @('MERGE_HEAD','REBASE_HEAD','CHERRY_PICK_HEAD','rebase-merge','rebase-apply')) {
      if (Test-Path (Join-Path $gitdir $f)) { $blocked = $true; break }
    }
    if ($blocked) { Write-Log "skip : merge/rebase running  [$Ws]"; continue }

    $dirty = (& git -C $Ws status --porcelain 2>$null)
    if (-not $dirty) { continue }

    & git -C $Ws add -A 2>$null
    $staged = (& git -C $Ws diff --cached --name-only 2>$null)
    if (-not $staged) { continue }

    $n = @($staged).Count
    $msg = 'autosave: {0} ({1} files)' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $n
    & git -c "user.name=$AuthorName" -c "user.email=$AuthorEmail" -C $Ws commit -q -m $msg 2>$null
    if ($LASTEXITCODE -eq 0) { Write-Log "committed $n files  [$Ws]" }
    else { Write-Log "commit FAILED (exit $LASTEXITCODE)  [$Ws]" }
  }
}
finally {
  Remove-Item $Lock -Force -ErrorAction SilentlyContinue
}
exit 0
