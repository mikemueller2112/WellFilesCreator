# -*- mode: python -*-
a = Analysis(['main_menu.py'],
             pathex=['c:\\Users\\mike mueller\\Desktop\\Well Files Creator'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='main_menu.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False , icon='mm.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=None,
               upx=True,
               name='main_menu')
