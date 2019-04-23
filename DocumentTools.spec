# -*- mode: python -*-

block_cipher = None


a = Analysis(['DocumentTools.py'],
             pathex=['E:\\CODE\\Jianli\\ExcelRenderWordTool'],
             binaries=[],
             datas=[('C:\\Users\\xuqiu\\Desktop\\SelfScan\\venv\\lib\\site-packages\\eel\\eel.js', 'eel'), ('web', 'web')],
             hiddenimports=['bottle_websocket'],
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
          name='DocumentTools',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
