# -*- mode: python ; coding: utf-8 打包指令 pyinstaller ExcelTool.spec --clean -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config/settings.ini', 'config'),           # 配置文件
        ('src/resources/styles.qss', 'src/resources'),  # 样式表
        ('src/resources/icons/newapp.ico', 'src/resources/icons')  # 图标（打包后会被主程序引用？实际上exe图标用--icon指定了）
    ],
    hiddenimports=['pandas', 'openpyxl', 'xlrd', 'xlwt'],
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
    name='绿能ExcelTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='src/resources/icons/newapp.ico'
)