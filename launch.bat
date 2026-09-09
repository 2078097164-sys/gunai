@echo off
rem Desktop Shortcut Tool launcher. Runs the tool with a hidden console approach.
title Desktop Shortcut Backup and Restore Tool
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0DesktopShortcutTool.ps1"
if errorlevel 1 (
  echo.
  echo Tool exited with an error. Message above.
  pause
)
