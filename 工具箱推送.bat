@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

:: ============================================
::   工具箱推送 - 一键 Git 推送脚本
::   双击即可运行，自动初始化并推送到 GitHub
::   使用 SSH 方式，无需令牌
:: ============================================

echo.
echo ============================================
echo    工具箱推送
echo ============================================
echo.

:: --------------------------------------------
:: 配置区域
:: --------------------------------------------
set "REPO_NAME=tool-box"
set "BRANCH=main"
set "REMOTE_URL=git@github.com:TianQue6916/tool-box.git"

:: --------------------------------------------
:: 自动检测脚本所在目录作为项目根目录
:: --------------------------------------------
set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"

echo [信息] 项目目录: %PROJECT_DIR%
echo.

cd /d "%PROJECT_DIR%"

:: --------------------------------------------
:: 初始化 Git 仓库（如果尚未初始化）
:: --------------------------------------------
if not exist ".git" (
    echo [步骤1] 初始化 Git 仓库...
    git init
    if !errorlevel! neq 0 (
        echo [错误] Git 初始化失败！
        pause
        exit /b 1
    )
    git checkout -b %BRANCH% 2>nul
    echo [成功] Git 仓库初始化完成！
    echo.
) else (
    echo [步骤1] Git 仓库已存在，跳过初始化。
    echo.
)

:: --------------------------------------------
:: 检查远程仓库是否已配置
:: --------------------------------------------
git remote get-url origin >nul 2>&1
if !errorlevel! neq 0 (
    echo [步骤2] 配置远程仓库...
    git remote add origin %REMOTE_URL%
    if !errorlevel! neq 0 (
        echo [错误] 添加远程仓库失败！
        pause
        exit /b 1
    )
    echo [成功] 远程仓库配置完成: %REMOTE_URL%
    echo.
) else (
    echo [步骤2] 远程仓库已配置，跳过。
    echo.
)

:: --------------------------------------------
:: 确保 .gitignore 存在
:: --------------------------------------------
if not exist ".gitignore" (
    echo [步骤3] 创建 .gitignore 文件...
    (
        echo # Python
        echo __pycache__/
        echo *.py[cod]
        echo *$py.class
        echo *.so
        echo *.egg-info/
        echo dist/
        echo build/
        echo.
        echo # Virtual Environment
        echo venv/
        echo .env/
        echo.
        echo # IDE
        echo .vscode/
        echo .idea/
        echo *.swp
        echo *.swo
        echo.
        echo # OS
        echo .DS_Store
        echo Thumbs.db
        echo desktop.ini
        echo.
        echo # Output / temp files
        echo *.tmp
    ) > .gitignore
    echo [成功] .gitignore 创建完成！
    echo.
) else (
    echo [步骤3] .gitignore 已存在，跳过。
    echo.
)

:: --------------------------------------------
:: 添加所有文件到暂存区
:: --------------------------------------------
echo [步骤4] 添加文件到暂存区...
git add -A
if !errorlevel! neq 0 (
    echo [错误] git add 失败！
    pause
    exit /b 1
)
echo [成功] 文件添加完成！
echo.

:: --------------------------------------------
:: 检查是否有变更需要提交
:: --------------------------------------------
git diff --cached --quiet >nul 2>&1
if !errorlevel! equ 0 (
    echo [信息] 没有新的变更需要提交。
    echo.
    goto :push_step
)

:: --------------------------------------------
:: 生成提交信息（自动包含日期时间）
:: --------------------------------------------
for /f "tokens=1-6 delims=/ " %%a in ("%date%") do set "DATE_STR=%%a-%%b-%%c"
for /f "tokens=1-3 delims=:." %%a in ("%time%") do set "TIME_STR=%%a-%%b-%%c"
set "COMMIT_MSG=自动推送 %DATE_STR% %TIME_STR%"

echo [步骤5] 提交变更...
git commit -m "%COMMIT_MSG%"
if !errorlevel! neq 0 (
    echo [错误] git commit 失败！
    pause
    exit /b 1
)
echo [成功] 提交完成: %COMMIT_MSG%
echo.

:push_step
:: --------------------------------------------
:: 推送到 GitHub
:: --------------------------------------------
echo [步骤6] 推送到 GitHub...
git push -u origin %BRANCH%
if !errorlevel! neq 0 (
    echo [信息] 首次推送可能失败，尝试设置上游并重试...
    git branch -M %BRANCH%
    git push -u origin %BRANCH%
    if !errorlevel! neq 0 (
        echo [错误] 推送失败！请检查网络连接和SSH密钥配置。
        pause
        exit /b 1
    )
)
echo.
echo ============================================
echo    推送成功！
echo ============================================
echo.
echo 仓库地址: https://github.com/TianQue6916/%REPO_NAME%
echo.

:: 自动打开浏览器查看仓库
start https://github.com/TianQue6916/%REPO_NAME%

pause
