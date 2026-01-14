# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['loadcalculation.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 添加可能需要的数据文件
        ('data_template.csv', '.'),
        # 添加项目目录，这样已有的项目数据会被包含
        ('projects', 'projects'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.constants',
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
        'ttk',
        'tkinter.font',
        'tkinter.tix',
        'tkinter.ttk',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'pkg_resources.py2_warn',  # 针对matplotlib的兼容性问题
        'tkinter.scrolledtext',  # 可能缺失的组件
        'tkinter.simpledialog',  # 可能缺失的组件
        'tkinter.colorchooser',  # 可能缺失的组件
        'tkinter.commondialog',  # 可能缺失的组件
        'tkinter.dialog',        # 可能缺失的组件
        'tkinter.dnd',           # 可能缺失的组件
        'tkinter.filedialog',    # 可能缺失的组件
        'tkinter.messagebox',    # 可能缺失的组件
        'tkinter.ttk',           # 可能缺失的组件
        'tkinter.constants',     # 可能缺失的组件
        'tkinter.__init__',      # 可能缺失的组件
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
    name='XJ-EnergyBalanceSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为False以避免显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 如果有图标文件可以指定路径
)