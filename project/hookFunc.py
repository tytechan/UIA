#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import PublicFunc
PF = PublicFunc()
if __name__ == "__main__":
    # 打开自调chrome, 非默认安装chrome需修改 path 参数
    PF.openChrome()

    # 开始录制
    PF.startHook()