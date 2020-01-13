#!python3
# -*- coding: utf-8 -*-
import tkinter
import tkinter.messagebox as msg

# def returnResult(var):
#     return var

def getInput(title, message):
    def return_callback(event):
        # print('quit...')
        root.quit()

    def close_callback():
        global quitConfirm
        # msg.showinfo('提示', '请输入控件名并回车...')
        quitConfirm = msg.askokcancel("请确认", "是否退出此控件保存？")
        if quitConfirm:
            root.destroy()

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

    # btn1 = tkinter.Button(root, text="确定", command=returnResult(True)).pack()
    # btn1.grid(row=0, column=0, padx=5,pady=5)
    # btn2 = tkinter.Button(root, text="取消", font=('Helvetica 10 bold'), command=root.quit()).pack()
    # btn2.grid(row=0, column=0, padx=5,pady=5)

    root.protocol("WM_DELETE_WINDOW", close_callback)
    root.mainloop()

    # if "quitConfirm" not in locals().keys() or not quitConfirm:
    #     str = entry.get()
    #     root.destroy()
    #     return str

    try:
        assert quitConfirm == False
        str = entry.get()
        root.destroy()
        return str
    except NameError:
        str = entry.get()
        root.destroy()
        return str
    except AssertionError:
        pass


def message_askyesno(title, message):
    top = tkinter.Tk()
    top.withdraw()                          # 实现主窗口隐藏
    top.update()                            # 需要update一下
    txt = tkinter.messagebox.askyesno(title, message)
    top.destroy()
    return txt

if __name__ == "__main__":
    a = getInput("title", "msg")
    print(a)
    # message_askyesno("提示", "测试？")
