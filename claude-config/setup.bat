@echo off
chcp 65001 > nul
echo ========================================
echo  Claude Code セットアップ
echo ========================================
echo.

REM --- パス設定（このPCに合わせて変更） ---
set CLAUDE_DIR=%USERPROFILE%\.claude
set SCRIPT_DIR=%~dp0

echo [1/3] スキルファイルをコピーします...
if not exist "%CLAUDE_DIR%\skills" mkdir "%CLAUDE_DIR%\skills"
xcopy /Y /Q "%SCRIPT_DIR%skills\*" "%CLAUDE_DIR%\skills\"
echo     完了: %CLAUDE_DIR%\skills\

echo.
echo [2/3] settings.json を確認してください...
echo     テンプレート: %SCRIPT_DIR%settings-template.json
echo     配置先:       %CLAUDE_DIR%\settings.json
echo.
echo     ※ フックのパス（${SHIRYO_ROOT}, ${BUMON_AGENT}）を
echo       このPCの実際のパスに書き換えてから手動でコピーしてください。
echo.

echo [3/3] CLAUDE.md の配置を確認してください...
echo     shiryo/ 直下の CLAUDE.md は git clone で自動配置されます。
echo     各テーマフォルダの CLAUDE.md も同様です。
echo.

echo ========================================
echo  セットアップ案内が完了しました。
echo  settings.json のパスを手動で書き換えてください。
echo ========================================
pause
