"""启动入口（源码运行与打包均使用本脚本，避免相对导入问题）。

用法：
    源码运行：  python launcher.py
    打包：      pyinstaller ... launcher.py
"""
from app.main import main

if __name__ == "__main__":
    main()
