#!python3
# -*- coding: utf-8 -*-
import re
import frozen_dir
import shutil
from config.DirAndTime import *
from PyInstaller.utils.cliutils.set_version import *


configFile = "version_control.txt"                                  # 版本控制配置文件路径
versionPath = "%s\\versions" %os.getcwd()                           # 备份路径
versionFile = "%s\\version_log.txt" %versionPath                    # 版本日志文件（\dist 目录下）
imgPath = "%s\\image\\app_logo.ico" %frozen_dir.app_path()          # 执行程序图标路径

def versionControl(appName, appDesc):
    ''' 发布前写入应用版本信息
    :param appName: 应用名称
    :param appDesc: 应用描述
    '''
    try:
        with open(configFile, "a+", encoding='UTF-8'):
            pass
        with open(configFile, "r+", encoding='UTF-8') as f:
            rawData = f.read()
            assert rawData != "", \
                "版本控制配置文件内容为空，请检查后再发布！"

            # 应用名称
            ProductName = re.findall(r"'ProductName', u\'(.+?)\'", rawData)[0]
            nameList = rawData.split(ProductName)
            rawData = appName.join(nameList)
            # 应用描述
            FileDescription = re.findall(r"'FileDescription', u\'(.+?)\'", rawData)[0]
            descriptionList = rawData.split(FileDescription)
            appDesc = appDesc.replace("\n", "").replace(" ", "")
            rawData = appDesc.join(descriptionList)
            # 版本号
            FileVersion = re.findall(r"'FileVersion', u\'(.+?)\'", rawData)[0]
            versionList = rawData.split(FileVersion)
            if FileVersion == "首次发布":
                FileVersion = "1.0.0"
            else:
                myList = FileVersion.split(".")
                i = len(myList) - 1
                while i >= 0:
                    if i == 0 or int(myList[i]) != 9:
                        myList[i] = str(int(myList[i]) + 1)
                        flag = False
                    else:
                        myList[i] = "0"
                        flag = True

                    if not flag:
                        break
                    i -= 1

                FileVersion = ".".join(myList)
            rawData = FileVersion.join(versionList)

            ProductVersion = re.findall(r"'ProductVersion', u\'(.+?)\'", rawData)[0]
            productList = rawData.split(ProductVersion)
            rawData = FileVersion.join(productList)
            # print(rawDict)
        with open(configFile, "w", encoding='UTF-8') as f:
            f.write(str(rawData))
            f.close()
        print("配置文件更新完成！")
    except AssertionError as e:
        raise e
    except Exception as e:
        raise e


def publish():
    ''' 发布并记录版本日志 '''
    try:
        # cmd = r"pyinstaller -D -i %s .\editFunc.py" %imgPath
        cmd = r"pyinstaller --version-file version_control.txt -D -i %s .\editFunc.py -y" %imgPath
        os.system(cmd)
        print("应用发布完成！")

        with open(configFile, "r+", encoding='UTF-8') as x:
            configInfo = x.read()
            FileVersion = re.findall(r"'FileVersion', u\'(.+?)\'", configInfo)[0]
            x.close()

        if not os.path.exists(versionPath):
            os.makedirs(versionPath)
        shutil.copyfile("editFunc.py", os.path.join(versionPath, "%s.py" %FileVersion))
        print("当前版本备份完成！")

        with open(versionFile, "a+", encoding='UTF-8'):
            pass
        with open(versionFile, "r+", encoding='UTF-8') as f:
            rawData = f.read()
            rawDict = eval(rawData) if rawData else dict()

            assert  FileVersion not in rawDict.keys(), "版本重复，发布前请先更新配置文件！"
            rawDict[FileVersion] = dict()
            rawDict[FileVersion]["version_info"] = appDesc
            rawDict[FileVersion]["update_time"] = "%s %s" %(getCurrentDate(), getCurrentTime())
            # print(rawDict)
        with open(versionFile, "w", encoding='UTF-8') as f:
            f.write(str(rawDict))
            # f.close()
        print("版本日志更新完成！")

    except AssertionError as e:
        raise e
    except Exception as e:
        raise e

if __name__ == "__main__":
    # 应用名称
    appName = "壳系统登陆1"
    # 应用描述
    appDesc = \
    '''
    1、abc修改修改修改修改修改修改修改修改修改修改修改修改；
    2、完成完成完成完成完成完成完成完成完成完成完成完成完成完成完成完成。
    '''

    # 更新配置文件
    versionControl(appName, appDesc)
    # 发布并记录版本信息
    publish()