# -*- mode: python ; coding: utf-8 -*-

project_root = r'E:\SZQS\Server\Steel_Pipe_Bailey_Support_Cal_Server'  

a = Analysis(
    ['Server_UI.py'],                        # 主程序
    pathex=[project_root],
    binaries=[],
    datas=[
    ('run_server.py', '.'),                     
    ('Cal_server.py', '.'),                     
    ('Cal_task_executor_new.py', '.'),          
    ('email_send.py', '.'),                     
    ],

    hiddenimports=[
        # 你的 UI 会启动 run_server，所以这里要确保相关模块包含
        'run_server',
        'Cal_server',
        'Cal_task_executor_new',
        'email_send',
        'fastapi',
        'uvicorn',
        'starlette',
        'requests',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Server_UI',        # 生成的可执行文件名
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,               # 为避免打包后 DLL 压缩崩溃，建议关闭
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,            
)
