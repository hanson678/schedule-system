# -*- coding: utf-8 -*-
"""打包路径辅助模块
PyInstaller打包后 __file__ 指向临时目录，
可写数据(data/、uploads/)需要在exe同级目录。
"""
import os
import sys


def get_app_dir():
    """获取应用根目录
    - 打包后：sys._MEIPASS（_internal目录，含templates/static/data）
    - 开发时：源码所在目录
    """
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))
