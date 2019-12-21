# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['src\\ical2csv_gui.py'],
             pathex=['C:\\Users\\christian\\Documents\\libreoffice-ical-importer'],
             binaries=[],
             datas=[('C:\\Users\\christian\\AppData\\Local\\Programs\\Python\\Python38-32\\Lib\\site-packages\\ics\\grammar\\contentline.ebnf', 'ics\\grammar\\')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='ical2csv_gui',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='images\\icon.ico')
