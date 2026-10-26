# -*- mode: python ; coding: utf-8 -*-
"""德育分自动化系统 · 自写安装器打包配置（方案 3.4 节 B 路线）。

前置：先跑 tools/make_payload.py 生成 _build_pkg\payload.zip（内嵌进安装器）。
产物：德育分自动化系统-安装程序-<版本>.exe（onefile，双击即向导）。
"""

import os

REPO_ROOT = os.path.dirname(os.path.abspath(SPECPATH))
PAYLOAD_ZIP = os.path.join(REPO_ROOT, "_build_pkg", "payload.zip")

a = Analysis(
    ['installer.py'],
    pathex=[],
    binaries=[],
    datas=[(PAYLOAD_ZIP, '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas', 'openpyxl', 'matplotlib', 'scipy', 'sympy', 'IPython',
        'notebook', 'jupyter', 'pytest', 'sphinx', 'numba', 'llvmlite',
        'pyarrow',
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
    name='德育分自动化系统-安装程序-2.1.1.0',
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
