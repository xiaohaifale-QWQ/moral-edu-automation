# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— onedir 窗口模式。

识别/语音/OCR 模块已经移除，所以程序不再依赖
torch / paddle / cv2 / whisper / sounddevice，
这里显式排除掉，产物体积能小一大截。
"""

from PyInstaller.utils.hooks import collect_data_files

# tkinterdnd2：拖拽支持（名单模板拖进窗口导入），必须带上它自带的 tkdnd 二进制
TKDND_DATAS = collect_data_files('tkinterdnd2')


a = Analysis(
    ['德育分自动化系统.py'],
    pathex=[],
    binaries=[],
    datas=[('doro.ico', '.')] + TKDND_DATAS,
    hiddenimports=['pandas', 'openpyxl', 'tkwebui', 'tkinterdnd2'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 已移除功能带来的重依赖
        'paddle', 'paddlepaddle', 'paddleocr', 'whisper', 'sounddevice',
        'torch', 'torchvision', 'torchaudio', 'cv2',
        # pandas 的可选 JIT 加速路径（运行时按需 import，缺失会自动回退）
        # 依赖链 pandas.core._numba -> numba -> llvmlite，其中 llvmlite.dll 单个就有 115MB
        'numba', 'llvmlite',
        # 用不到的科学计算 / 交互环境 / 数据格式引擎
        'matplotlib', 'scipy', 'sympy', 'IPython', 'notebook', 'jupyter',
        'pytest', 'sphinx', 'pypinyin', 'pyarrow', 'tables',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    icon='doro.ico',
    name='德育分自动化系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='德育分自动化系统',
)
