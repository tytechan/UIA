# -*- coding: utf-8 -*-
import subprocess
import os

def runCMD(cmd):
    try:
        # cmd = "pipreqs ./ --encoding=utf8"
        popen = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        popen.wait()
        lines = popen.stdout.readlines()
    except BaseException as e:
        # return -1
        return e
    except Exception as e:
        raise e

if __name__ == "__main__":
    file = os.getcwd() + r"\requirements.txt"
    if os.path.exists(file):
        os.remove(file)

    runCMD("pipreqs ./ --encoding=utf8")
    print("依赖包清单生成完毕！")
    # runCMD("pip install -r requirements.txt")
    print("安装完毕！")