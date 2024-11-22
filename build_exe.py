import PyInstaller.__main__
import sys
import os

# 确保使用 UTF-8 编码
if sys.platform.startswith('win'):
    os.environ['PYTHONIOENCODING'] = 'utf-8'

PyInstaller.__main__.run([
    'template_processor.py',
    '--name=皮皮虾模板替换',
    '--onefile',
    '--windowed',
    '--noconsole',
    '--clean',
    '--add-data', 'template_processor.py;.',
    '--hidden-import', 'pandas',
    '--hidden-import', 'openpyxl',
    '--hidden-import', 'tkinter',
    '--collect-data', 'pandas',
    '--collect-data', 'openpyxl',
    '--noupx',
    # '--icon', 'icon.ico',  # 如果有图标的话
]) 