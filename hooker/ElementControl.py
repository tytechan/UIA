#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import AppControl

class Controls(AppControl):
    def __init__(self, element):
        AppControl.__init__(self)
        self.element = element

    def click(self):
        self.element.Click()