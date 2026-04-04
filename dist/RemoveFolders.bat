@echo off
chcp 65001 >nul
title RemoveFolders
color 0A

:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ============================================
    echo   Please run as Administrator!
    echo ============================================
    pause
    exit /b
)

:MENU
cls
echo ============================================
echo        This PC - Folder Manager
echo ============================================
echo.
echo   [1] Remove 3D Objects (3D duixiang)
echo   [2] Remove Videos (shipin)
echo   [3] Remove Music (yinyue)
echo   [4] Remove Pictures (tupian)
echo   [5] Remove Documents (wendang)
echo   [6] Remove Downloads (xiazai)
echo   [7] Remove Desktop (zhuomian)
echo.
echo   [A] Remove ALL 7 folders
echo   [B] Keep Downloads+Documents+Desktop only
echo   [R] Restore ALL 7 folders
echo   [Q] Quit
echo.
echo ============================================

set /p choice="Enter option: "

if /i "%choice%"=="1" goto DEL_3D
if /i "%choice%"=="2" goto DEL_VIDEO
if /i "%choice%"=="3" goto DEL_MUSIC
if /i "%choice%"=="4" goto DEL_PICTURE
if /i "%choice%"=="5" goto DEL_DOCUMENT
if /i "%choice%"=="6" goto DEL_DOWNLOAD
if /i "%choice%"=="7" goto DEL_DESKTOP
if /i "%choice%"=="A" goto DEL_ALL
if /i "%choice%"=="B" goto DEL_RECOMMEND
if /i "%choice%"=="R" goto RESTORE_ALL
if /i "%choice%"=="Q" exit /b

echo Invalid option!
timeout /t 2 >nul
goto MENU

:DEL_3D
call :RemoveFolder "{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" "3D Objects"
goto DONE

:DEL_VIDEO
call :RemoveFolder "{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" "Videos"
goto DONE

:DEL_MUSIC
call :RemoveFolder "{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}" "Music"
goto DONE

:DEL_PICTURE
call :RemoveFolder "{24ad3ad4-a569-4530-98e1-ab02f9417aa8}" "Pictures"
goto DONE

:DEL_DOCUMENT
call :RemoveFolder "{d3162b92-9365-467a-956b-92703aca08af}" "Documents"
goto DONE

:DEL_DOWNLOAD
call :RemoveFolder "{088e3905-0323-4b02-9826-5d99428e115f}" "Downloads"
goto DONE

:DEL_DESKTOP
call :RemoveFolder "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}" "Desktop"
goto DONE

:DEL_ALL
call :RemoveFolder "{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" "3D Objects"
call :RemoveFolder "{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" "Videos"
call :RemoveFolder "{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}" "Music"
call :RemoveFolder "{24ad3ad4-a569-4530-98e1-ab02f9417aa8}" "Pictures"
call :RemoveFolder "{d3162b92-9365-467a-956b-92703aca08af}" "Documents"
call :RemoveFolder "{088e3905-0323-4b02-9826-5d99428e115f}" "Downloads"
call :RemoveFolder "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}" "Desktop"
goto DONE

:DEL_RECOMMEND
call :RemoveFolder "{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" "3D Objects"
call :RemoveFolder "{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" "Videos"
call :RemoveFolder "{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}" "Music"
call :RemoveFolder "{24ad3ad4-a569-4530-98e1-ab02f9417aa8}" "Pictures"
echo.
echo   Kept: Downloads, Documents, Desktop
goto DONE

:RESTORE_ALL
call :RestoreFolder "{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" "3D Objects"
call :RestoreFolder "{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" "Videos"
call :RestoreFolder "{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}" "Music"
call :RestoreFolder "{24ad3ad4-a569-4530-98e1-ab02f9417aa8}" "Pictures"
call :RestoreFolder "{d3162b92-9365-467a-956b-92703aca08af}" "Documents"
call :RestoreFolder "{088e3905-0323-4b02-9826-5d99428e115f}" "Downloads"
call :RestoreFolder "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}" "Desktop"
echo.
echo   All folders restored!
goto DONE

:RemoveFolder
set "GUID=%~1"
set "NAME=%~2"
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\%GUID%" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\%GUID%" /f >nul 2>&1
echo   [OK] Removed: %NAME%
goto :eof

:RestoreFolder
set "GUID=%~1"
set "NAME=%~2"
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\%GUID%" /f >nul 2>&1
reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\%GUID%" /f >nul 2>&1
echo   [OK] Restored: %NAME%
goto :eof

:DONE
echo.
echo ============================================
echo   Done! Restarting Explorer...
echo ============================================
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 >nul
start explorer.exe
echo   Explorer restarted. Check result now.
echo.
pause
goto MENU