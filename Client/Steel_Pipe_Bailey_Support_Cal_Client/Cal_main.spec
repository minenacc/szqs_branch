# -*- mode: python ; coding: utf-8 -*-

# 请根据你的项目根目录调整下面这个变量
project_root = r'E:\SZQS\Client\Steel_Pipe_Bailey_Support_Cal_Client'

a = Analysis(
    ['Cal_main.py'],
    pathex=[project_root],  # 项目根路径，保证能找到模块
    binaries=[],
    datas=[
        (project_root + r'\resources\logo.ico', 'resources')  
    ],
    hiddenimports=[
        'auth',
        'Cal_client',
        'Cal_Report_UI',
        'RSA',
        'ttkbootstrap',
        'ttkbootstrap.constants',
        'tksheet',
        'requests',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='数智桥施网络版',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=project_root + r'\resources\logo.ico',  # 绝对路径图标
)
