#!python3
# -*- coding: utf-8 -*-
import tkinter
import tkinter.messagebox as msg


def getInput(title, message):
    def return_callback(event):
        # print('quit...')
        root.quit()

    def close_callback():
        msg.showinfo('提示', '请输入值并回车...')

    root = tkinter.Tk(className=title)
    root.wm_attributes('-topmost', 1)
    screenwidth, screenheight = root.maxsize()
    width = 300
    height = 100
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root.geometry(size)
    root.resizable(0, 0)
    lable = tkinter.Label(root, height=2)
    lable['text'] = message
    lable.pack()
    entry = tkinter.Entry(root)
    entry.bind('<Return>', return_callback)
    entry.pack()
    entry.focus_set()
    root.protocol("WM_DELETE_WINDOW", close_callback)
    root.mainloop()
    str = entry.get()
    root.destroy()
    return str


if __name__ == "__main__":
    a = getInput("title", "msg")
    print(a)