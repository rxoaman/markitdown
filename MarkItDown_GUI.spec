# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = [
    'markitdown',
    'markitdown.converters',
    'markitdown.converters._base_converter',
    'markitdown.converters._plain_text_converter',
    'markitdown.converters._html_converter',
    'markitdown.converters._rss_converter',
    'markitdown.converters._wikipedia_converter',
    'markitdown.converters._youtube_converter',
    'markitdown.converters._ipynb_converter',
    'markitdown.converters._bing_serp_converter',
    'markitdown.converters._pdf_converter',
    'markitdown.converters._docx_converter',
    'markitdown.converters._xlsx_converter',
    'markitdown.converters._xls_converter',
    'markitdown.converters._pptx_converter',
    'markitdown.converters._image_converter',
    'markitdown.converters._audio_converter',
    'markitdown.converters._outlook_msg_converter',
    'markitdown.converters._zip_converter',
    'markitdown.converters._epub_converter',
    'markitdown.converters._csv_converter',
    'tkinterdnd2',
]

packages_to_collect = [
    'tkinterdnd2',
    'markitdown',
    'mammoth',
    'pptx',
    'openpyxl',
    'pdfminer',
    'pdfplumber',
    'pypdfium2',
    'speech_recognition',
    'pydub',
    'markdownify',
    'defusedxml',
    'olefile',
    'charset_normalizer',
    'bs4',
    'lxml',
]

for pkg in packages_to_collect:
    try:
        ret = collect_all(pkg)
        datas += ret[0]
        binaries += ret[1]
        hiddenimports += ret[2]
    except Exception as e:
        print(f"Hook warning for {pkg}: {e}")

# Python 3.15 DLL support
python_dir = os.path.dirname(sys.executable)
abi3t_dll = os.path.join(python_dir, 'abi3t-compat', 'python3t.dll')
if os.path.exists(abi3t_dll):
    binaries.append((abi3t_dll, '.'))

python3_dll = os.path.join(python_dir, 'python3.dll')
if os.path.exists(python3_dll):
    binaries.append((python3_dll, '.'))

a = Analysis(
    ['markitdown_gui.py'],
    pathex=['.', os.path.join('packages', 'markitdown', 'src')],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'unittest', 'IPython'],
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
    name='MarkItDown_GUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
