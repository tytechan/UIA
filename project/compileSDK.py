#!python3
# -*- coding: utf-8 -*-
import py_compile
import compileall
import os
import shutil

# 固定SDK目录路径
parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
localSDKPath = parentDirPath + "\localSDK"

def compileDir(dir=localSDKPath):
    '''
    编译路径下所有.py文件为二进制.pyc文件，隐藏源码（默认为固定SDK路径）
    :param dir: 编译路径
    :return: 编译路径下生成“__pycache__”
    '''
    compileall.compile_dir(dir)

def compileFile(fileName, dir=parentDirPath):
    '''
    编译单个.py文件为二进制.pyc文件，隐藏源码（默认为固定SDK路径下文件）
    :param fileName: 编译文件名（带后缀）
    :param dir: 编译路径
    :return: 编译路径下生成“__pycache__”
    '''
    file = r"%s\%s" %(dir, fileName)
    py_compile.compile(file)

def handleFile(dir, fileName=None):
    '''
    打包并拷贝至源文件相同路径
    :param dir: 编译路径
    :param fileName: 编译文件名（带后缀）
    :return: 打包后源文件目录下会生成相同文件名的.pyc文件，可用于直接调用
    '''
    if fileName:
        compileFile(fileName, dir)
    else:
        compileDir(dir)

    try:
        pycDir = r"%s\__pycache__" %dir
        files = os.listdir(pycDir)
        for file in files:
            if file.endswith(".pyc"):
                # print(file)
                newName = file.replace(".cpython-35", "")
                shutil.copyfile(os.path.join(pycDir, file), os.path.join(dir, newName))
    except Exception as e:
        raise e


if __name__ == "__main__":
    '''
    # 编译所有固定SDK
    handleFile(localSDKPath)

    # 编译固定SDK下单个文件
    handleFile(localSDKPath, "BrowserFunc.py")
    '''