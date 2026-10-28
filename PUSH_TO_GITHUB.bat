@echo off
setlocal
cd /d "%~dp0"

echo ====================================================================
echo        Push YASKAWA Accounting Automation to GitHub
echo ====================================================================
echo.

set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;%PATH%"

where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Git is not found in PATH or at %LOCALAPPDATA%\Programs\Git\cmd.
    pause
    exit /b 1
)

echo Step 1: Create a new repository on GitHub:
echo         - Open https://github.com/new in your browser
echo         - Choose a repository name (e.g. yaskawa-accounting-automation)
echo         - DO NOT initialize with README, .gitignore, or license
echo         - Click "Create repository"
echo.
echo Step 2: Copy the repository HTTPS URL (e.g. https://github.com/username/repo.git)
echo.
set /p REPO_URL="Enter your GitHub Repository URL: "

if "%REPO_URL%"=="" (
    echo [ERROR] No URL provided. Aborted.
    pause
    exit /b 1
)

git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%
echo.
echo Added remote origin: %REPO_URL%
echo Pushing branch 'main' to GitHub...
echo.

git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ====================================================================
    echo [SUCCESS] Project pushed to GitHub successfully!
    echo.
    echo To run this project on GitHub:
    echo 1. Go to your repository on GitHub.
    echo 2. Click the "Actions" tab.
    echo 3. Click "Verify YASKAWA Accounting Automation" -^> "Run workflow".
    echo 4. After it finishes, download the generated Excel deliverables under "Artifacts".
    echo ====================================================================
) else (
    echo.
    echo [NOTICE] Push failed or authentication required.
    echo If prompted for password, use a GitHub Personal Access Token (PAT):
    echo https://github.com/settings/tokens
)

echo.
pause
