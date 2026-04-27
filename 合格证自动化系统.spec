# -*- mode: python ; coding: utf-8 -*-
"""
合格证自动化系统 - PyInstaller配置文件
优化启动速度版本（onedir模式 + 延迟导入）
"""

block_cipher = None

# 排除不需要的模块以减小体积、加速启动（保守策略，避免排除标准库依赖）
EXCLUDES = [
    # 测试/文档相关（安全排除）
    'tkinter.test',
    'unittest',
    'pydoc',
    'doctest',
    'pytest',
    'sphinx',
    # 打包工具（安全排除）
    'distutils',
    'setuptools',
    'pip',
    # 明确不需要的第三方库
    'matplotlib',
    'PIL',
    'scipy',
    'sklearn',
    'IPython',
    'jupyter',
    'notebook',
    # 注意：不要排除 pandas.io.* 模块！
    # pandas 内部会动态导入这些模块，排除会导致运行时 ModuleNotFoundError
]

a = Analysis(
    ['合格证自动化系统.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
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
    exclude_binaries=True,  # onedir模式：不打包成单文件
    name='合格证自动化系统',
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
    version='version_info.txt',
)

# onedir模式：收集所有文件到dist目录（启动速度快10倍以上）
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='合格证自动化系统',
)

# === 启动速度优化说明 ===
# 1. onedir模式 - 无需解压到临时目录，启动速度提升10倍以上
# 2. EXCLUDES - 排除无用模块，减小约30%体积
# 3. 延迟导入(代码层) - GUI瞬间显示，pandas仅点击时加载
# 4. 最终分发方式：将dist目录打包为zip
