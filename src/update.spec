# -*- mode: python ; coding: utf-8 -*-
"""update.exe 打包配置 —— onefile、无窗口（出错只弹 MessageBox，见 update.py 的 alert）。

更新器自身也可能被更新包更新：update.py 启动时会把自己复制到
%TEMP%\deyufen-update\ 再执行，所以 onefile 没问题。
"""

a = Analysis(
    ['update.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas', 'openpyxl', 'matplotlib', 'scipy', 'sympy', 'IPython',
        'notebook', 'jupyter', 'pytest', 'sphinx', 'numba', 'llvmlite',
        'pyarrow', 'tkinter',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='update',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
