# reasonix-ops — Reasonix 工作区体检与 Merge-Back 收尾

> 建于 2026-09-21。把「工作区正在被写入 / 合并至主分支报错」的排查与修复固化成可重复执行的流程。

## 什么时候跑

出现以下任一症状：

- 界面提示 **当前工作区正在被写入** / **另一个 Reasonix 窗口当前正在写入这个工作区**
- **合并至主分支** 面板报：`worktree has uncommitted changes`、`the recorded source checkout has uncommitted changes`、
  `a source or worktree runtime still has active or waiting work`、`以下内容会阻止清理并保留工作区`
- 工具调用直接失败：`acquire workspace write lease: acquire file lock: context canceled`
- 日志出现：`workspace mutation conflicts with persisted state`

## 用法

```powershell
cd $env:USERPROFILE\Desktop\reasonix-ops

# 1) 体检（默认，零写入）
powershell -NoProfile -ExecutionPolicy Bypass -File .\reasonix-doctor.ps1

# 2) 可逆修复：让 source + worktree 的未提交项归零（改 .git\info\exclude + git reset）
powershell -NoProfile -ExecutionPolicy Bypass -File .\reasonix-doctor.ps1 -Heal

# 3) 额外清理 worktree 内的任务日志（备份后才删）
powershell -NoProfile -ExecutionPolicy Bypass -File .\reasonix-doctor.ps1 -Heal -IncludeTaskLogs

# 指定别的工作区
powershell -NoProfile -ExecutionPolicy Bypass -File .\reasonix-doctor.ps1 -Workspace "D:\some\repo"
```

## 四条阻断项对照（界面文案 → 实质 → 处置）

| 界面文案 | 实质 | 处置 |
|---|---|---|
| `worktree has uncommitted changes` | worktree 里 `.reasonix/` 等未提交/未忽略文件 | `-Heal` |
| `the recorded source checkout has uncommitted changes` | 源 checkout 的 `git status` 非空 | `-Heal` |
| `a source or worktree runtime still has active or waiting work` | 有会话正在该 worktree 里跑（`task.lock` 存在） | **结束会话后重试**，脚本无法代做 |
| `tracked, untracked, or ignored files would be preserved` | worktree 内 `.reasonix\tasks\**` 任务日志（**ignored 也算**） | `-Heal -IncludeTaskLogs` |

## 为什么会有这些"错误"

Reasonix 的工作区写隔离粒度是 **workspaceRoot**（不是文件）：同工作区同一时刻只有一个写者持有整区写租约，租约在首次变更时获取、
并贯穿参与的 subagent 与后台任务（官方 PR #6542 / `docs/SESSION_RECOVERY_AND_PARALLELISM.zh-CN.md`）。
而 worktree 一旦启用，收尾（Merge-Back）会做一套预检：**源 checkout 干净、无进行中 Git 操作、无活动任务、无待保留文件**。
`.reasonix/` 是 Reasonix 的运行时目录、天天产生新文件——只要它在跟踪范围内，源目录就**永远不可能"干净"**，这就是同类报错反复出现的原因。

## 硬约束（脚本内部也不会越线）

1. 不 `commit` / 不 `push` / 不改 `config.toml` / 不重启应用 / 不删用户数据文件
2. 所有写操作前先备份到 `backup\<时间戳>\`（index、info\exclude、任务日志、被移动的 .gitignore）
3. 排除规则写进 **`.git\info\exclude`**，不写工作区 `.gitignore`——后者自身作为未跟踪文件会让状态又不干净
4. 中文路径规则必须用 `Get-ChildItem` 的**真实 Unicode 名**；git 输出会经 `core.quotepath` 转义成 `\346\226\271`，直接采纳会匹配失败
5. 界面上的「自动提交工作区未提交修改」**不要勾**：它会把运行时文件提交进分支

## 回滚

```powershell
# 恢复某次修复前的 git 元数据
Copy-Item .\backup\<时间戳>\index   "$env:APPDATA\reasonix\global-workspace\.git\index" -Force
Copy-Item .\backup\<时间戳>\exclude "$env:APPDATA\reasonix\global-workspace\.git\info\exclude" -Force
```

恢复某个文件被 git 跟踪：删掉 `.git\info\exclude` 里对应行，或 `git add -f <路径>`。

## 注意

`-Heal` 让工作区"对 git 干净"的手法是**把未提交项排除在版本控制之外**（磁盘文件一个不动）。
代价是这些文件不再被 git 跟踪；若要真正版本化它们，应另建干净仓库，而不是在这个杂物工作区上做。
