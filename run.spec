# -*- mode: python -*-

block_cipher = None


a = Analysis(['src\\run.py'],
             pathex=['src/', 'C:\\Users\\user\\Documents\\Personal Coding Projects\\PC Projects\\beta_cal'],
             binaries=[],
             datas=[],
             hiddenimports=['YahooToExcel', 'PricePoint'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='run',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
