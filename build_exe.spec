# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec文件 - 排期录入系统打包配置"""

import os
import sys

block_cipher = None
BASE = os.path.abspath('.')

# 数据文件：templates、static、data目录
datas = [
    (os.path.join(BASE, 'templates'), 'templates'),
    (os.path.join(BASE, 'static'), 'static'),
    (os.path.join(BASE, 'data', 'config.json'), os.path.join('data')),
    (os.path.join(BASE, 'data', 'item_schedule_map.json'), os.path.join('data')),
]

# 如果存在sku_mapping.json也打包
sku_map = os.path.join(BASE, 'data', 'sku_mapping.json')
if os.path.exists(sku_map):
    datas.append((sku_map, 'data'))

a = Analysis(
    ['app.py'],
    pathex=[BASE],
    binaries=[],
    datas=datas,
    hiddenimports=[
        # Flask相关
        'flask', 'jinja2', 'markupsafe', 'click', 'itsdangerous', 'werkzeug',
        # pywin32 COM
        'win32com', 'win32com.client', 'win32api', 'win32con',
        'pythoncom', 'pywintypes', 'win32com.server',
        # pdfplumber / pdfminer
        'pdfplumber', 'pdfminer', 'pdfminer.high_level',
        'pdfminer.layout', 'pdfminer.pdfpage',
        # openpyxl
        'openpyxl', 'openpyxl.styles',
        # APScheduler
        'apscheduler', 'apscheduler.schedulers.background',
        'apscheduler.triggers.interval',
        # 其他
        'email', 'email.mime', 'email.mime.text', 'email.mime.multipart',
        'email.mime.base', 'email.mime.application',
        'imaplib', 'smtplib',
        'logging', 'logging.handlers',
        'json', 'csv', 're', 'threading',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', 'matplotlib', 'numpy', 'scipy', 'pandas',
        'PIL', 'pillow', 'PyMuPDF', 'fitz', 'pypdfium2',
        'pythonnet', 'clr_loader', 'pywebview', 'bottle',
        'pytest', 'unittest',
    ],
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
    exclude_binaries=True,
    name='排期系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # 保留控制台窗口，方便看日志
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='排期系统',
)
