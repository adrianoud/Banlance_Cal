# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['loadcalculation.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 添加可能需要的数据文件
        ('data_template.csv', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.PhotoImage',
        'matplotlib.pyplot',
        'matplotlib.backends.backend_tkagg',
        'numpy',
        'csv',
        'json',
        'os',
        'datetime',
        'openpyxl',
        'openpyxl.Workbook',
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
    name='XJ-EnergyBalanceSystem-Console',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 设置为True以查看控制台输出和错误信息
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 如果有图标文件可以指定路径
)