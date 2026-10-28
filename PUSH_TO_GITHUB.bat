@echo off
setlocal
cd /d "%~dp0"

echo ====================================================================
echo        Push YASKAWA Accounting Automation to GitHub
echo ====================================================================
echo Target: https://github.com/GYaswanth030604/yaskawa-accounting-automation
echo.

set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;%PATH%"

where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Git was not found in PATH or at %LOCALAPPDATA%\Programs\Git\cmd.
    pause
    exit /b 1
)

set REPO_URL=https://github.com/GYaswanth030604/yaskawa-accounting-automation.git

git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%
echo Remote origin set to %REPO_URL%
echo.
echo Pushing branch 'main' to GitHub...
echo (If prompted, enter your GitHub Username and Personal Access Token or sign in)
echo.

git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ====================================================================
    echo [SUCCESS] Project successfully pushed to GitHub!
    echo.
    echo Next Steps - How to run on GitHub:
    echo 1. Open: https://github.com/GYaswanth030604/yaskawa-accounting-automation/actions
    echo 2. Click "Verify YASKAWA Accounting Automation" -^> "Run workflow".
    echo 3. Download the generated Excel workbooks under "Artifacts".
    echo ====================================================================
) else (
    echo.
    echo ====================================================================
    echo [NOTICE] Push needs authentication!
    echo.
    echo GitHub requires a Personal Access Token (PAT) for password authentication:
    echo 1. Generate token here: https://github.com/settings/tokens/new
    echo    - Expiration: 30 days (or no expiration)
    echo    - Select scope: 'repo' (Full control of private/public repositories)
    echo 2. When prompted for Password, paste the generated token (starts with ghp_).
    echo ====================================================================
)

echo.
pause
