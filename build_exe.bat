@echo off
setlocal

echo ============================================================
echo  WINS Wafer Loading Automation -- Build Script
echo ============================================================
echo.

REM A running WINS.exe holds files inside dist\WINS\ open (it loads
REM Assets\qcells_logo.png into memory for as long as it's running,
REM for instance) -- PyInstaller can't delete or replace those files
REM while something still has them locked, which fails confusingly
REM deep inside the build rather than with a clear reason up front.
tasklist /FI "IMAGENAME eq WINS.exe" 2>nul | find /I "WINS.exe" >nul
if not errorlevel 1 (
    echo [ERROR] WINS.exe is currently running -- close it completely
    echo         before rebuilding. PyInstaller can't replace files that
    echo         a running process still has open, which is exactly what
    echo         just failed if you're seeing this after a previous
    echo         build attempt.
    echo.
    echo         Close WINS.exe fully, check Task Manager if you're not
    echo         sure it's really gone, then run this script again.
    echo.
    pause
    exit /b 1
)

where pyinstaller >nul 2>nul
if errorlevel 1 (
    echo [ERROR] pyinstaller is not installed or not on PATH.
    echo         Run this first:  pip install -r requirements-build.txt
    echo.
    pause
    exit /b 1
)

echo [INFO] Checking that WINS's own dependencies are available in this
echo        same Python environment...
python -c "import customtkinter, win32com.client, pywinauto, tkcalendar, openpyxl, pyperclip, dotenv" 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] One or more of WINS's own dependencies are missing from
    echo         this Python environment -- customtkinter, pywin32,
    echo         pywinauto, tkcalendar, openpyxl, pyperclip, or
    echo         python-dotenv.
    echo.
    echo         PyInstaller has to run in the SAME environment as WINS
    echo         itself, since it needs to find and bundle everything WINS
    echo         imports. Having pyinstaller installed by itself is not
    echo         enough if these aren't here too.
    echo.
    echo         Fix:  pip install -r requirements.txt
    echo         Then run this script again.
    echo.
    pause
    exit /b 1
)
echo [INFO] Dependencies OK.
echo.

REM config.py fails fast with a clear message if .env is missing or
REM incomplete (see config.py itself for why) -- catching that HERE,
REM before spending time on a full PyInstaller build, is much better
REM than finding out only after opening the freshly-built .exe and
REM watching it immediately crash.
echo [INFO] Checking that .env is set up correctly...
python -c "import config" 2>build_env_check_error.tmp
if errorlevel 1 (
    echo.
    echo [ERROR] config.py failed to load -- .env is likely missing or
    echo         incomplete. Python's exact error:
    echo.
    type build_env_check_error.tmp
    echo.
    echo         Fix:  copy .env.example to .env, then fill in your own
    echo         SAP credentials and plant/storage codes. Then run this
    echo         script again.
    echo.
    del build_env_check_error.tmp
    pause
    exit /b 1
)
del build_env_check_error.tmp 2>nul
echo [INFO] .env OK.
echo.

if not exist "wins_icon.ico" (
    echo [INFO] wins_icon.ico not found in this folder -- the build will
    echo        still work, just without a custom app icon. Move it here
    echo        alongside WINS.spec if you want one.
    echo.
)

echo [INFO] Cleaning previous build folders...
if exist build rmdir /s /q build
if exist dist_previous rmdir /s /q dist_previous

REM Keep the previous dist\WINS around under a temp name rather than
REM deleting it outright -- that's where Data\ (your live Excel file)
REM lives, and we want to preserve it across rebuilds, not just avoid
REM overwriting it during the copy step below.
if exist dist (
    ren dist dist_previous
)

echo [INFO] Running PyInstaller...
echo.
pyinstaller WINS.spec

if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller build failed -- see the output above for why.
    echo         Your previous build is still in dist_previous\WINS if you need it.
    pause
    exit /b 1
)

echo.
echo [INFO] Setting up Data, Assets, and Logs alongside the new .exe...

REM Data\ holds your live, edited Excel file -- carry it over from the
REM previous build if one exists, otherwise seed it fresh from the
REM project source. Either way, never overwrite an existing copy.
REM Written with goto instead of nested if/else blocks -- deeply
REM nested parentheses in batch have subtle parsing edge cases that
REM are hard to verify without a live Windows shell to test against,
REM and goto sidesteps that risk entirely.
if exist "dist_previous\WINS\Data" goto data_from_previous
if exist "Data" goto data_from_source
echo [WARN] No Data folder found anywhere -- you'll need to add one manually.
goto data_done

:data_from_previous
xcopy /E /I /Y "dist_previous\WINS\Data" "dist\WINS\Data" >nul
echo [INFO] Carried over your existing Data folder from the previous build.
goto data_done

:data_from_source
xcopy /E /I /Y "Data" "dist\WINS\Data" >nul
echo [INFO] Seeded Data folder from the project source (first build).

:data_done

REM Assets is static branding, not user data -- always safe to refresh
REM from source so it stays in sync with the project.
if not exist "Assets" goto assets_done
xcopy /E /I /Y "Assets" "dist\WINS\Assets" >nul
echo [INFO] Refreshed Assets folder.
:assets_done

REM .env holds real SAP credentials -- the .exe needs it sitting right
REM alongside it to actually run (config.py looks for it relative to
REM its own location, not the current working directory). The
REM pre-flight check earlier already confirmed .env exists and is
REM complete, so this is just placing it where the packaged app can
REM find it. dist\WINS\ is gitignored the same as Data\ and Logs\, so
REM this doesn't introduce any new exposure -- it's a local build
REM output, not something that gets committed.
if not exist ".env" goto env_done
copy /Y ".env" "dist\WINS\.env" >nul
echo [INFO] Copied .env alongside the .exe.
:env_done

REM Carry over existing logs/history too, same reasoning as Data.
if exist "dist_previous\WINS\Logs" goto logs_from_previous
if not exist "dist\WINS\Logs" mkdir "dist\WINS\Logs"
goto logs_done

:logs_from_previous
xcopy /E /I /Y "dist_previous\WINS\Logs" "dist\WINS\Logs" >nul

:logs_done

if not exist "dist_previous\WINS\doi_history.json" goto skip_doi_history
copy /Y "dist_previous\WINS\doi_history.json" "dist\WINS\doi_history.json" >nul
:skip_doi_history

if exist dist_previous rmdir /s /q dist_previous

echo.
echo ============================================================
echo  BUILD COMPLETE
echo.
echo  Your app is in:      dist\WINS\
echo  Run it by opening:   dist\WINS\WINS.exe
echo.
echo  The WHOLE dist\WINS\ folder is what you copy to another
echo  machine -- not just the .exe by itself, since customtkinter
echo  needs its data files sitting right alongside it.
echo.
echo  dist\WINS\.env now holds your real SAP credentials, copied
echo  there so the .exe can actually log in. Treat this folder the
echo  same way you'd treat the credentials themselves -- fine to
echo  copy to another machine you control, not fine to upload
echo  anywhere public.
echo ============================================================
echo.

endlocal
pause
