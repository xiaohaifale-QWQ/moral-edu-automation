# -*- mode: python ; coding: utf-8 -*-
"""mcp_server.exe 打包配置 —— onefile、console 模式（MCP 走 stdio，必须有控制台句柄）。

只依赖 openpyxl，体积约 10 MB。用户机器无需安装 Python。
BASE_DIR 已改为按 sys.executable 定位（见 mcp_server.py 头部），
所以 exe 放在 <安装目录>\AI对接\ 下即可找到 名单模板.xlsx 与 _ai_bridge\。
"""

a = Analysis(
    ['mcp_server.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas', 'matplotlib', 'scipy', 'sympy', 'IPython', 'notebook',
        'jupyter', 'pytest', 'sphinx', 'numba', 'llvmlite', 'pyarrow',
        'paddle', 'paddlepaddle', 'paddleocr', 'whisper', 'sounddevice',
        'torch', 'torchvision', 'torchaudio', 'cv2', 'tkinter',
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
    name='mcp_server',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
