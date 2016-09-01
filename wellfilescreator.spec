# -*- mode: python -*-
a = Analysis(['wellfilescreator.py'],
             pathex=['c:\\Users\\mike mueller\\Desktop\\Well Files Creator'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='wellfilescreator.exe',
          debug=False,
          strip=None,
          upx=True,
          console=True , icon='(mm.ico)')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=None,
               upx=True,
               name='wellfilescreator')
