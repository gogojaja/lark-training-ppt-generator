#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ASCII 启动器：加载中文命名的图形界面主程序。

避免 .bat 文件因中文路径/UTF-8 codepage 解析问题而失败。
双击 启动_PPT流程图工具.bat → 本文件 → 主程序 GUI。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    main_py = os.path.join(HERE, "PPT流程图工具.py")
    if not os.path.isfile(main_py):
        sys.stderr.write("找不到主程序: %s\n" % main_py)
        input("按回车退出…")
        sys.exit(1)

    try:
        namespace = {"__name__": "__main__", "__file__": main_py}
        with open(main_py, encoding="utf-8") as f:
            code = f.read()
        exec(compile(code, main_py, "exec"), namespace)
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("程序出错，按回车退出…")
        sys.exit(1)


if __name__ == "__main__":
    sys.exit(main())
