# -*- mode: python ; coding: utf-8 -*-
"""
合格证自动化系统 - PyInstaller配置文件
降低杀毒软件误报的优化配置
"""

block_cipher = None

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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='合格证自动化系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # ⚠️ 禁用UPX压缩以降低误报
    console=False,  # 隐藏控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Windows版本信息
    version='version_info.txt',
    # 图标（如有.ico文件可取消注释）
    # icon='app_icon.ico',
)

# === 降低杀毒软件误报的配置说明 ===
# 1. upx=False - UPX压缩常被杀毒软件标记为可疑
# 2. strip=False - 保留调试信息，提升可信度
# 3. version - 添加详细版本信息
# 4. console=False - GUI应用不需要控制台
