# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：Streamlit + pandas + altair 单文件 exe。

构建：python -m PyInstaller excel-toolbox.spec --clean --noconfirm
产物：dist/ExcelToolbox.exe
"""

from PyInstaller.utils.hooks import collect_all, collect_submodules, copy_metadata

datas = [
    ("app.py", "."),
    (".streamlit", ".streamlit"),
    ("excel-docs", "excel-docs"),
]
binaries = []
hiddenimports = [
    "openpyxl",
    "xlsxwriter",
    "streamlit.web.cli",
    "streamlit.runtime.scriptrunner.magic_funcs",
]
hiddenimports += collect_submodules("core")
hiddenimports += collect_submodules("ui")

datas += copy_metadata("streamlit")
datas += copy_metadata("altair")

# streamlit 的 static 静态资源和 altair 的 schema 都在包目录里，必须整体收进来
for package in ("streamlit", "altair"):
    package_datas, package_binaries, package_hidden = collect_all(package)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hidden

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tests", "pytest", "matplotlib", "scipy", "IPython", "notebook"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ExcelToolbox",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)
