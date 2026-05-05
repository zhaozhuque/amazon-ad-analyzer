"""
打包脚本 - 将工具打包为独立的 EXE 文件
使用方法：python build_exe.py
"""

import subprocess
import sys
import os

def build():
    # 确保 pyinstaller 已安装
    try:
        import PyInstaller
    except ImportError:
        print("正在安装 PyInstaller...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])

    print("=" * 50)
    print("  开始打包 亚马逊广告诊断工具")
    print("=" * 50)

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',                          # 打包成单个exe
        '--windowed',                         # 不显示命令行窗口
        '--name', 'AmazonAdAnalyzer',         # exe文件名
        '--icon', 'icon.ico',                 # 图标（可选，没有就去掉这行）
        '--add-data', 'README.md;.',          # 打包README
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'numpy',
        '--hidden-import', 'pandas',
        'amazon_ad_tool.py'
    ]

    # 如果没有 icon.ico，去掉图标参数
    if not os.path.exists('icon.ico'):
        cmd = [c for c in cmd if c != '--icon' and c != 'icon.ico']

    try:
        subprocess.check_call(cmd)
        print()
        print("=" * 50)
        print("  打包成功！")
        print(f"  EXE位置: dist/AmazonAdAnalyzer.exe")
        print("=" * 50)
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        sys.exit(1)


if __name__ == '__main__':
    build()
