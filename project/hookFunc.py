#!python3
# -*- coding: utf-8 -*-

from localSDK.BasicFunc import PublicFunc
PF = PublicFunc()

if __name__ == "__main__":
    # 打开自调chrome, 非默认安装chrome需修改 path 参数
    # PF.openChrome(r"C:\Users\47612\AppData\Local\Google\Chrome\Applicationpath")
    #
    # 打开自调Firefox（须保证geckodriver路径已配环境变量）
    # profileDir = "C:\\Users\\47612\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\ytnvvpv8.default-release"
    # PF.openFirefox(profileDir)

    # 开始录制
    PF.startHook()
