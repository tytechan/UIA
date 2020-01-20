# -*- coding: utf-8 -*-
import sys
import os

def app_path():
    """Returns the base application path."""
    # if hasattr(sys, 'frozen'):
    #     # Handles PyInstaller
    #     return os.path.dirname(sys.executable)          #使用pyinstaller打包后的exe目录
    # return os.path.dirname(__file__)                    #没打包前的py目录

    if getattr(sys, 'frozen', False):
        root_path = os.path.dirname(sys.executable)
    elif __file__:
        root_path = os.path.dirname(__file__)
    return root_path