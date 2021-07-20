__author__ = 'zds'
# -*- coding: utf-8 -*-
# @Time    : 2021/6/7 0007 上午 11:54
# @Author  : zds
# @FileName: view.py
# @Software: PyCharm
from threading import Thread
from  tkinter import Tk, ttk, StringVar, LEFT,filedialog,END,Frame,BOTH,Label,Button,Entry,HORIZONTAL
import win32file
from do_main import *
from os import startfile

in_path = None
out_path = None
in_path_2_1 = None
in_path_2_2 = None
out_path_2 = None
in_path_2_3 = None
in_path_2_4 = None
a = []
b = []
c = []
d = []
e = []
f = []
g = []
data_dick2 = {}
path_j = PATH_J
p_c = PATH_TEMP
p_c1_name = PATH_TEMP_NAME1
p_c2_name = PATH_TEMP_NAME2


def thread_it(func, *args):
    '''将函数打包进线程'''
    # 创建
    t = Thread(target=func, args=args)
    # 守护 !!!
    t.setDaemon(True)
    # 启动
    t.start()
    # 阻塞--卡死界面！
    # t.join()


def two():
    start()
    set_lab()


def two2():
    start2()
    set_lab2()


def star():  # 开始多线程任务
    thread_it(two)
    thread_it(tk_do_work)


def star2():
    thread_it(two2)
    thread_it(tk_do_work2)


# 进度条开始
def start(*args):
    display(p1)
    p1.start(30)


# 进度条停止
def stop(*args):
    value = p1['value']
    p1.stop()
    p1['value'] = value


# 进度条开始
def start2(*args):
    display(p2)
    p2.start(30)


# 进度条停止
def stop2(*args):
    value = p2['value']
    p2.stop()
    p2['value'] = value


# 控件显示
def display(con):
    con.place(x=100, y=200, heigh=30)


# 控件隐藏
def forget(con):
    con.place_forget()


# end_进度条
def end_stop():
    stop()

    # forget(p1)
    # msg()

# 消息窗口
# def msg(text):
#     m = messagebox.showinfo("Notice", text)
#     return m


# 获取文件名路径
def select_path1():
    path_ = filedialog.askopenfilename(filetypes=[('Excel Document', '*.zip')])
    path_i.set(path_)


# 获取文件夹路径
def select_path3():
    path_ = filedialog.askdirectory()
    path_i.set(path_)


# 选择保存文件的路径
def select_path2():
    path_ = filedialog.asksaveasfilename(filetypes=[('Excel Document', '*.csv')], defaultextension=[('Excel Document', '*.csv')])
    path_o.set(path_)


def select_path4():
    path_ = filedialog.askopenfilename(filetypes =[('Excel Document', '*.xls'), ('Excel Document', '*.csv'), ('Excel Document', '*.xlsx')])
    path_m.set(path_)


def select_path5():
    path_ = filedialog.askopenfilename(filetypes=[('Excel Document', '*.xls'), ('Excel Document', '*.csv'), ('Excel Document', '*.xlsx')])
    path_n.set(path_)


# 选择保存文件的路径
def select_path6():
    path_ = filedialog.asksaveasfilename(filetypes=[('Excel Document', '*.csv')], defaultextension=[('Excel Document', '*.csv')])
    path_q.set(path_)


def set_lab():
    lab4.config(text="合并中......")


def set_lab2():
    # lab4.config(text="合并中......")
    pass


def callback():
    n1 = ent1.get()
    a.append(n1)
    n2 = ent2.get()
    b.append(n2)


def callback2():
    n3 = ent4.get()
    c.append(n3)
    n4 = ent5.get()
    d.append(n4)
    n5 = ent6.get()
    e.append(n5)
    n6 = ent7.get()
    f.append(n6)
    n7 = ent8.get()
    g.append(n7)


def list_n():
    global a
    global b
    global in_path
    global out_path
    try:
        in_path = a.pop()
        out_path = b.pop()
    except IOError as e:
        pass


def list_n2():
    global c
    global d
    global e
    global in_path_2_1, in_path_2_2, in_path_2_3, in_path_2_4
    global out_path_2
    try:
        in_path_2_1 = c.pop()
        in_path_2_2 = d.pop()
        out_path_2 = e.pop()
        in_path_2_3 = f.pop()  # 数组
        in_path_2_4 = g.pop()  # 数组
    except IOError as e:
        pass


def _clear():
    ent1.delete(0, END)
    ent2.delete(0, END)


def _clear2():
    ent4.delete(0, END)
    ent5.delete(0, END)
    ent6.delete(0, END)
    ent7.delete(0, END)
    ent8.delete(0, END)


def tk_do_work2():
    global in_path_2_1, in_path_2_2, in_path_2_3, in_path_2_4, path_g_2
    global out_path_2
    callback2()
    list_n2()
    s_t2 = time()
    if in_path_2_1 != "" and in_path_2_2 != "" and out_path_2 != "" and in_path_2_3 != "" and in_path_2_4 != "" and path_g_2 != "":
        text2, f2 = do_work2(in_path_2_1, in_path_2_2, in_path_2_3, in_path_2_4, out_path_2, path_g_2, s_t2)
        if f2 is False:
            _clear2()
            stop2()
            forget(p2)
            com_box_list2.current(0)
            return
    else:
        is_ok = msg("选择的路径或输入键值为空！！")
        if is_ok == 'ok':
            stop2()
            forget(p2)
        return
    msg(text2)
    if f2 is True:
        _clear2()
        stop2()
        forget(p2)
        com_box_list2.current(0)


def tk_do_work():
    global in_path
    global out_path
    callback()
    list_n()
    s_t = time()
    if in_path != "" and out_path != "":
        text, fl = do_work(in_path, out_path, s_t)
        if fl is False:
            _clear()
            lab4.config(text='')
            stop()
            forget(p1)
            return
    else:
        is_ok = msg("选择的路径为空！！")
        if is_ok == 'ok':
            lab4.config(text='')
            stop()
            forget(p1)
        return
    msg(text)
    if fl is True:
        _clear()
        lab4.config(text='')
        stop()
        forget(p1)


def go(*args):  # 处理事件，*args表示可变参数
    global path_c
    path_c = com_box_list.get()  # 打印选中的值
    # print(path_c)
    # print(type(path_c))


def go1(*args):  # 处理事件，*args表示可变参数
    global path_d
    path_d = com_box_list1.get()  # 打印选中的值


def go2(*args):
    global path_g_2
    path_g_2 = com_box_list2.get()  # 打印选中的值


def save_d(d_k, d_v):
    global path_c
    global path_s
    dic = {}
    try:
        if path.exists(path_j):
            with open(path_j, 'r') as fr:
                dic = literal_eval(fr.read())  # 读取的str转换为字典
                fr.close()
            if is_hide_file(path_j):
                set_f_display(path_j)
        dic[d_k] = d_v
        with open(path_j, 'w') as fw:
            fw.write(str(dic))
            fw.close()
        msg("保存成功！")
        path_c = ""
        path_s = ""
        com_box_list.current(0)
        set_f_hide(path_j)
    except Exception as e:
        msg("保存失败！{}".format(e))


def del_d(d):
    global path_d
    dic = {}
    try:
        if path.exists(path_j):
            with open(path_j, 'r') as fr:
                dic = eval(fr.read())  # 读取的str转换为字典
                fr.close()
            if is_hide_file(path_j):
                set_f_display(path_j)
        if d in dic.keys():
            dic.pop(d)
        with open(path_j, 'w') as fw:
            fw.write(str(dic))
            fw.close()
        msg("删除成功！")
        com_box_list1.current(0)
        set_f_hide(path_j)
    except Exception as e:
        msg("删除失败！{}".format(e))


def is_hide_file(path_file):  # 判断是否是隐藏文件
    file_flag = win32file.GetFileAttributesW(path_file)
    is_hide = file_flag & win32con.FILE_ATTRIBUTE_HIDDEN
    if is_hide == 2:
        return True
    else:
        return False


def set_f_hide(path_file):  # 设置隐藏文件
    file_flag = win32file.GetFileAttributesW(path_file)
    is_hide = file_flag & win32con.FILE_ATTRIBUTE_HIDDEN
    if is_hide != 2:
        win32api.SetFileAttributes(path_file, win32con.FILE_ATTRIBUTE_HIDDEN)  # 设置path_j为隐藏文件


def set_f_display(path_file):  # 显示隐藏文件
    file_flag = win32file.GetFileAttributesW(path_file)
    is_hide = file_flag & win32con.FILE_ATTRIBUTE_HIDDEN
    if is_hide != 0:
        win32api.SetFileAttributes(path_file, win32con.FILE_ATTRIBUTE_NORMAL)


def save_config():
    global path_s
    path_s = ent3.get()
    if path_c != "" and path_s != "":
        save_d(path_s,path_c)
    else:
        msg("存入字段为空！")


def del_config():
    if path_d != "":
        del_d(path_d)
    else:
        msg("删除的字段不能为空！")


def query_to_tuple():
    if path.exists(path_j):
        with open(path_set_j, 'r') as fr:
            dict1 = literal_eval(fr.read())  # 读取的str转换为字典
            fr.close()
            return tuple(dict1)
    else:
        return ()


def query_config():
    data1 = DATA_DICK
    data = {}
    data.update(data1)
    if path.exists(path_j):
        with open(path_set_j, 'r') as fr:
            dict2 = literal_eval(fr.read())  # 读取的str转换为字典
            data.update(dict2)
            fr.close()
    txt1 = ""
    j = 1
    for t in data.keys():
        txt1 += (t+"  ")
        if j % 3 == 0:
            txt1 += '\n'
        j += 1
    lab7.config(text=txt1)


def read_me():
    startfile(r'rm.txt')


# 创建主窗口
myWindow = Tk()
myWindow.title("金三导出表格合并工具V1.1（试用版）")
# 得到屏幕宽度
sw = myWindow.winfo_screenwidth()
# 得到屏幕高度，设置居中
sh = myWindow.winfo_screenheight()
x = (sw-600) / 2
y = (sh-300) / 2
myWindow.geometry("600x300+%d+%d" % (x, y))
myWindow.maxsize(800, 300)
myWindow.minsize(600, 300)
try:
    myWindow.iconbitmap('myico.ico')
except:
    pass

# 选项卡
notebook = ttk.Notebook(myWindow)
frameOne = Frame()
frameTwo = Frame()
frameThree = Frame()
frameFour = Frame()
# 进度条
p1 = ttk.Progressbar(frameOne, length=400, mode="indeterminate", maximum=200, orient=HORIZONTAL)
p2 = ttk.Progressbar(frameTwo, length=400, mode="indeterminate", maximum=200, orient=HORIZONTAL)
# p1.place(x=135, y=160, heigh=30)
lab1 = Label(frameOne, text='请选择需要合并的zip文件或已解压的文件夹')
lab1.place(x=20, y=10)
lab2 = Label(frameOne, text='请选择保存路径')
lab2.place(x=20, y=80)

lab4 = Label(frameOne, text="")
lab4.place(x=200, y=160)
lab8 = Label(frameOne, text="或")
lab8.place(x=392, y=40)
path_i = StringVar()
path_o = StringVar()
path_c = ""
path_d = ""
path_g_2 = ""
path_s = StringVar()
path_m = StringVar()
path_n = StringVar()
path_q = StringVar()
path_x = StringVar()
path_y = StringVar()
# # 添加enter
ent1 = Entry(frameOne, textvariable=path_i, show=None)
ent1.place(x=20, y=40, heigh=30, width=300)
ent2 = Entry(frameOne, textvariable=path_o, show=None)
ent2.place(x=20, y=110, heigh=30, width=300)
# 添加按钮
btn1 = Button(frameOne, text='选择文件', command=select_path1)
btn1.place(x=320, y=40, heigh=30)
btn7 = Button(frameOne, text='选择文件夹', command=select_path3)
btn7.place(x=412, y=40, heigh=30)
btn2 = Button(frameOne, text='保存文件', command=select_path2)
btn2.place(x=320, y=110, heigh=30)
btn3 = Button(frameOne, text='一键合并', command=star)
btn3.place(x=420, y=78, heigh=60)

# frameThree设置选项卡的界面
lab5 = Label(frameThree, text="如发现表格数据变成科学计数形式，则需将数据对应字段设置为文本")
lab5.place(x=20, y=10)
lab6 = Label(frameThree, text="已设置的表格字段查看点击")
lab6.place(x=50, y=110)
lab7 = Label(frameThree, text="", justify=LEFT)
lab7.place(x=50, y=150)
ent3 = Entry(frameThree,textvariable=path_s,show=None)
ent3.place(x=20, y=40, heigh=30, width=150)
btn4 = Button(frameThree, text='保存', command=save_config)
btn4.place(x=250, y=40, heigh=30)
btn5 = Button(frameThree, text='查询', command=query_config)
btn5.place(x=250, y=110, heigh=30)
btn6 = Button(frameThree, text='删除', command=del_config)
btn6.place(x=250, y=75, heigh=30)
# 下拉框
com_value = StringVar()  # 窗体自带的文本，新建一个值
com_box_list = ttk.Combobox(frameThree, textvariable=com_value)  # 初始化
com_box_list["values"] = ("", str)
com_box_list.current(0)  # 选择第一个
com_box_list.bind("<<ComboboxSelected>>", go)  # 绑定事件,(下拉列表框被选中时，绑定go()函数)
com_box_list.place(x=170, y=40, heigh=30, width=80)

com_value1 = StringVar()  # 窗体自带的文本，新建一个值
com_box_list1 = ttk.Combobox(frameThree, textvariable=com_value1)  # 初始化
com_box_list1["values"] = ("",) + query_to_tuple()
com_box_list1.current(0)  # 选择第一个
com_box_list1.bind("<<ComboboxSelected>>", go1)  # 绑定事件,(下拉列表框被选中时，绑定go1()函数)
com_box_list1.place(x=20, y=75, heigh=30, width=230)

# frameTwo界面
lab8 = Label(frameTwo, text="该功能用于金三导出的不同种类的两张表格合并。", justify=LEFT)
lab8.place(x=20, y=10)
lab9 = Label(frameTwo, text="请选择保存路径", justify=LEFT)
lab9.place(x=20, y=130)
lab10 = Label(frameTwo, text="表一做键\n的字段：", justify=LEFT)
lab10.place(x=390, y=35)
lab11 = Label(frameTwo, text="表二做键\n的字段：", justify=LEFT)
lab11.place(x=390, y=75)
lab12 = Label(frameTwo, text="合并方式：",justify=LEFT)
lab12.place(x=390, y=120)
# # 添加enter
ent4 = Entry(frameTwo, textvariable=path_m, show=None)
ent4.place(x=20, y=40, heigh=30, width=250)
ent5 = Entry(frameTwo, textvariable=path_n, show=None)
ent5.place(x=20, y=80, heigh=30, width=250)
ent6 = Entry(frameTwo, textvariable=path_q, show=None)
ent6.place(x=20, y=160, heigh=30, width=250)
ent7 = Entry(frameTwo, textvariable=path_x, show=None)
ent7.place(x=470, y=40, heigh=30, width=120)
ent8 = Entry(frameTwo, textvariable=path_y, show=None)
ent8.place(x=470, y=80, heigh=30, width=120)
# 添加按钮
btn8 = Button(frameTwo, text='选择第一个文件', command=select_path4)
btn8.place(x=270, y=40, heigh=30)
btn9 = Button(frameTwo, text='选择第二个文件', command=select_path5)
btn9.place(x=270, y=80, heigh=30)
btn10 = Button(frameTwo, text='保存文件', command=select_path6)
btn10.place(x=270, y=160, heigh=30)
btn11 = Button(frameTwo, text='一键合并', command=star2)
btn11.place(x=390, y=160, heigh=30, width=120)
# 下拉框
com_value2 = StringVar()  # 窗体自带的文本，新建一个值
com_box_list2 = ttk.Combobox(frameTwo, textvariable=com_value2)  # 初始化
com_box_list2["values"] = ("", "inner", "outer", "left", "right")
com_box_list2.current(0)  # 选择第一个
com_box_list2.bind("<<ComboboxSelected>>", go2)  # 绑定事件,(下拉列表框被选中时，绑定go()函数)
com_box_list2.place(x=470, y=120, heigh=30, width=80)

# frameFour关于界面
lab3 = Label(frameFour, text="""   '# @version  1.1  \n# @Author:湛江张冬松'  """, justify=LEFT)
lab3.place(x=50, y=30)
btn12 = Button(frameFour, text='操作说明', command=read_me)
btn12.place(x=280, y=40, heigh=30, width=120)

notebook.add(frameOne, text='导出的同一类型表合并')
notebook.add(frameTwo, text='不同类型表合并（非zip文件）')
notebook.add(frameThree, text='设置')
notebook.add(frameFour, text='关于和操作说明')
notebook.pack(padx=10, pady=5, fill=BOTH, expand=True)
myWindow.mainloop()
