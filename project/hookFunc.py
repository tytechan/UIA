#!python3
# -*- coding: utf-8 -*-
from hooker.Hook import *


cf._init()
path = os.getcwd()
cf.set_value("path", path)

HK = Hooker()
HK.hooks()