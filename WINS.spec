# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for WINS Wafer Loading Automation.
#
# Run this ON WINDOWS (PyInstaller builds are platform-specific -- a
# build run anywhere else won't produce a working .exe):
#
#     pyinstaller WINS.spec
#
# --onedir (not --onefile) is REQUIRED here. customtkinter ships its
# own data files (theme .json, font .otf) that PyInstaller's --onefile
# mode does not pack correctly -- this is a confirmed, documented
# limitation of customtkinter + PyInstaller together, not something
# specific to WINS. See:
# https://customtkinter.tomschimansky.com/documentation/packaging/
#
# The output is a FOLDER (dist/WINS/), not a single file. WINS.exe
# lives at the top of that folder -- Data/, Assets/, and Logs/ should
# sit right alongside it (build_exe.bat handles copying them there
# after this spec finishes).

import os
import customtkinter

block_cipher = None

# customtkinter's own data files aren't picked up by PyInstaller's
# automatic analysis -- point at wherever it's actually installed on
# THIS build machine, whatever that path happens to be.
customtkinter_path = os.path.dirname(customtkinter.__file__)

# Cosmetic, not functional -- if the icon file didn't make it into the
# project folder for some reason, fall back to no custom icon rather
# than failing the entire build over it.
icon_path = 'wins_icon.ico' if os.path.exists('wins_icon.ico') else None
if icon_path is None:
    print("[WARN] wins_icon.ico not found in the project folder -- "
          "building without a custom icon. Move it there and rebuild "
          "if you want the app icon set.")

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        (customtkinter_path, 'customtkinter'),
    ],
    hiddenimports=[
        # win32com / pywin32 -- covers the late-binding GetObject() /
        # Dispatch() calls this project uses (not gencache.EnsureDispatch,
        # which has its own separate, well-known PyInstaller caching
        # issue that doesn't apply here, but keeping this list explicit
        # is cheap insurance either way).
        'win32com',
        'win32com.client',
        'win32timezone',
        'pythoncom',
        'pywintypes',
        # pywinauto's UI Automation backend (used in sap_launcher.py)
        # generates these comtypes modules at runtime the first time
        # they're needed -- PyInstaller's static analysis can't see a
        # dependency that doesn't exist as a file yet, so these have to
        # be listed explicitly or the frozen .exe fails on first launch
        # rather than at build time, which is a worse place to find out.
        'comtypes.gen',
        'comtypes.gen.UIAutomationClient',
        'comtypes.patcher',
        'comtypes.GUID',
        'pywinauto.application',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='WINS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,        # GUI app -- no console window behind it
    icon=icon_path,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WINS',
)
