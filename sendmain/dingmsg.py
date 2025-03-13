#  -*- coding: utf-8 -*-
"""
@File    : dingmsg-to-person.py，include menu。
@Date    : 2023-11-06
@Copyright :by BAPS V1.8
memo:消息来源文件由指定单一的mu.xls，改为可选择。
"""
import tkinter
import requests
import os
import pandas as pd
from PIL import ImageTk
from PIL import Image
import tkinter.messagebox as messagebox
import tkinter as tk
from tkinter import ttk  # 导入内部包
from miniini import filename_xls, filename_img, data_dict, filename_log  # 单独存储的全局变量
import time
from tkinter import filedialog

# 更改当前工作目录
os.chdir('D:\\pythonstart\\ddsvenv\\sendmain\\')

data_log = {}      # 存储日志数据的字典
#  生成log文件完整路径。此处不判断其是否存在，后面写入不存在的时候可自动创建。
fpath_log = os.path.join(os.getcwd(), filename_log)

def browse_file():
    # 此处不需要单独创建子窗口
    yfile_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    if yfile_path:
        process_file(yfile_path)

def process_file(yfile_path):
    # 读取xls文件
    # df = pd.read_excel(yfile_path)
    # 在这里对数据进行处理，例如打印前5行
    # print(df.head())
    menu_click('提示', f"你选择了消息源文件：{yfile_path}")

# 判断互联网是否连接，以ping通百度网页为准。
def ping(host):
    response = os.system("ping -n 3 " + host+'>null')    #-c 需要管理权限，-n不需要
    if response == 0:
        return True
    else:
        return False

def net_ck():
    hs="www.baidu.com"
    if ping(hs) == True:
        menu_click('提示', '网络测试通过！')
    else:
        menu_click('错误', '网络不通！')

# 读取日志到字典。
def read_log_f(log_path):
    global data_log
    # 引用全局变量data_log{}
    with open(log_path, 'r', encoding='ansi', errors='ignore') as file:    # 文本通常格式为ansi，不是utf-8
        lines = file.readlines()
        j = 0      # 增加计数的j，用作关键字。时间和手机均有重复，致使字典数据不完整。
        # 遍历每一行
        for line in lines:
            #  将每一行用*号分割成3个字符串。
            k1, v1, v2 = line.strip().split('*', 2)
            data_log[j] = [k1, v1, v2]
            j = j+1

# 将日志转成execl文件
def logtoexecl():
    global data_log
    # 将字典转换为DataFrame，直接转换达不到效果——元素值个数成了execl文件的行数，记录数成了列数。
    # 因此提取字典的元素值，逐行读出字典数据，写入execl——不必用循环，用list列表即可。
    values = list(data_log.values())
    # 将值转换为DataFrame
    df = pd.DataFrame(values, columns=['推送日期', '手机号', '已推送消息'])
    # 将DataFrame保存为Excel文件
    df.to_excel('secfile/logxls.xlsx', index=False)
    menu_click('提示', '已成功将日志写入文件secfile/logxls.xlsx！')

# 清空日志
def clear_log(fp_log):
    global data_log
    question = "确定要删除日志记录吗？"
    result = messagebox.askquestion(question, "确认删除", icon='warning', type='yesnocancel')
    if result == 'yes':
       with open(fp_log, 'w') as f:
            f.truncate(0)
       data_log={}            # 同时清空log字典

# 写入推送成功的消息日志
def write_log():
    filepath_log = os.path.join(os.getcwd(), filename_log)
    with open(filepath_log, "a") as f:
        current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        for key in data_dict.keys():
            # 将各个数值用*隔开。因为半角逗号可能是消息的一部分。
            f.write("{}*{}*{}\n".format(current_time,key, data_dict[key]))

# 读取指定exl文件，生成推送数据字典。
def read_exl_file(file_path):
    global data_dict    # 引用全局变量data_dict{}
    df = pd.read_excel(file_path)
    df['消息2'] = df['消息2'].fillna(value='')  #  将缺失值(NaN——Not a Number)替换为空。若为0效果较差。

    #  遍历DataFrame中的每一行数据
    for index, row in df.iterrows():
        #  提取手机和推送消息
        key = row['手机']
        value = row['姓名']+",〖"+row['消息1']+"〗"+str(row['消息2'])   # 组合生成推送消息，加str是因为备注为数字时会报错。

        #  将手机和推送消息作为键值对添加到字典中
        data_dict[key] = value

# 指定背景图片
def get_img(filename, width, height):
    im = Image.open(filename).resize((width, height))
    im = ImageTk.PhotoImage(im)
    return im

# 菜单程序
def menu_click(men1,mecon):
    messagebox.showinfo(men1, mecon)    # messagebox.showinfo('提示',f"{mecon}")——都可以

# 推送函数
def do_sendmsg():
    # 先判断网络状态，不通则退出，返回错误代码300。
    if ping('www.baidu.com') == False:
       menu_click('错误', '网络不通，不能推送！')
       exit(300)
    #  应用的唯一标识key
    appkey = 'dingxdafefawfase'   #替换成自己的appkey
    #  应用的密钥
    appsecret = 'TZVoSkT-9dya-H9iv2334-234234'   #替换成自己的appsecret
    #  推送消息时使用的微应用的AgentID，
    agent_id = '27446367243'    #替换成自己的agent_id
    token = get_access_token(appkey, appsecret)

    for key in data_dict.keys():
        userid_list =getUserIdByPhone(key, token['access_token'])
        ret = send_message(token['access_token'], {
            "agent_id": agent_id,
            "userid_list": userid_list,
            "msg": {
                "msgtype": "text",
                "text": {
                    "content": data_dict[key]
                },
            },
        })
    #  推送成功。此处if语句缩格在循环之外，检测的是最后一条推送操作的状态。若是逐条检测状态，则应将if缩至for之内！！！
    if ret['errcode'] == 0:
       messagebox.showinfo('提示', '恭喜恭喜，推送消息成功！')
       write_log()         # 考虑到逐条检测过于频繁，不取。此处统一将本次推送全部写入日志。
    else:
       messagebox.showinfo('错误', '推送消息失败！可重试！')

# 在子窗口中显示待推送消息
def show_table():
    global data_dict  #  引用全局变量data_dict{}
    subwin_wid = 800
    subwin_hig = 600
    newwin=tkinter.Toplevel(window)    # 创建子窗口
    # newwin = tk.Tk()
    newwin.title('浏览推送消息')
    newwin.geometry(f"{subwin_wid}x{subwin_hig}")

    #  创建滚动条
    scrollbar = ttk.Scrollbar(newwin)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    #  创建画布，并将滚动条与画布关联起来
    canvas = tk.Canvas(newwin, yscrollcommand=scrollbar.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    #  创建 treeview，并放置在画布上
    tree = ttk.Treeview(canvas, columns=['1', '2', '3'], show='headings')

    #  将滚动条的移动与画布关联起来
    scrollbar.config(command=tree.yview)
    #  tree.column('1', width=100, anchor='center')
    tree.column('1', width=int(float(f"{subwin_wid / 10}")), anchor='center')
    tree.column('2', width=int(float(f"{subwin_wid / 8}")), anchor='center')
    tree.column('3', width=int(float(f"{subwin_wid * 4 / 5}")), anchor='center')
    tree.heading('1', text='序号')
    tree.heading('2', text='手机')
    tree.heading('3', text='待推送消息')

    #  插入数据
    i=0
    for  key in data_dict.keys():
        tree.insert("", i, values=[str(i + 1), key, data_dict[key]])
        i = i + 1
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    newwin.mainloop()

# 在子窗口中显示推送日志
def show_log():
    global data_log  #  引用全局变量data_dict{}
    read_log_f(fpath_log)    # 调用函数读取日志到字典
    subwin_wid = 800
    subwin_hig = 600
    logwin=tkinter.Toplevel(window)    # 创建子窗口
    # logwin = tk.Tk()
    logwin.title('浏览日志内容')
    logwin.geometry(f"{subwin_wid}x{subwin_hig}")

    #  创建滚动条
    scrollbar = ttk.Scrollbar(logwin)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    #  创建画布，并将滚动条与画布关联起来
    canvas = tk.Canvas(logwin, yscrollcommand=scrollbar.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    #  创建 treeview，并放置在画布上
    tree = ttk.Treeview(canvas, columns=['1', '2', '3'], show='headings')

    #  将滚动条的移动与画布关联起来
    scrollbar.config(command=tree.yview)
    tree.column('1', width=int(float(f"{subwin_wid / 6}")), anchor='center')
    tree.column('2', width=int(float(f"{subwin_wid / 8 }")), anchor='w')
    tree.column('3', width=int(float(f"{subwin_wid*2 / 3 }")), anchor='center')
    tree.heading('1', text='推送时间')
    tree.heading('2', text='手机')
    tree.heading('3', text='已推送消息')

    #  插入数据
    for key, values in data_log.items():
        v1, v2, v3 = values
        tree.insert("", key,values=[v1, v2, v3])    # i可以不显示，但不可缺。
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    logwin.mainloop()

#  创建主窗口
window = tk.Tk()

# 退出程序，关闭所有子窗体
def prc_quit():
    for child in window.winfo_children():
        child.destroy()
        #  退出整个程序的mainloop
    window.destroy()
    exit(0)               # 退出程序

#  获取屏幕宽度和高度
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
#  设置窗口大小为屏幕大小
window.geometry(f"{screen_width}x{screen_height}")
window.title('考核激励消息精准推送')

# 显示背景图片，在画布canvas中显示。
window.resizable(False, False)
canvas_win = tk.Canvas(window, width=screen_width, height=screen_height)
im_win = get_img(filename_img, screen_width, screen_height)
canvas_win.create_image(screen_width/2, screen_height/2, image=im_win)
canvas_win.pack()

#  创建菜单栏
menu_bar = tk.Menu(window)

#  创建菜单项
file_menu = tk.Menu(menu_bar, tearoff=0)
edit_menu = tk.Menu(menu_bar, tearoff=0)
fssys_menu = tk.Menu(menu_bar, tearoff=0)

#  向菜单项添加子菜单项
file_menu.add_command(label="选择消息来源", command=browse_file)
file_menu.add_command(label="浏览推送内容", command=show_table)
file_menu.add_command(label="执行推送命令", command=do_sendmsg)
# file_menu.add_command(label="write", command=write_log)    # write_log
# file_menu.add_command(label="read", command=)
file_menu.add_separator()
file_menu.add_command(label="退出程序", command=prc_quit)

edit_menu.add_command(label="显示日志", command=show_log)
edit_menu.add_command(label="导出表格", command=logtoexecl)
# edit_menu.add_command(label="打印", command=lambda: menu_click('后续实现'))
# edit_menu.add_command(label="Paste", command=menu_click)

fssys_menu.add_command(label="清空日志", command=lambda: clear_log(fpath_log))
# 不加lambda:，直接上clear_log(fpath_log)，则程序一启动就会运行。
fssys_menu.add_command(label="网络测试", command=lambda: net_ck())
fssys_menu.add_separator()
fssys_menu.add_command(label="程序版本", command=lambda: menu_click('程序版本信息', '考核激励精准推送程序V1.6'))
#  将菜单项添加到菜单栏
menu_bar.add_cascade(label="推送消息", menu=file_menu)
menu_bar.add_cascade(label="查询统计", menu=edit_menu)
menu_bar.add_cascade(label="系统维护", menu=fssys_menu)

#  将菜单栏添加到主窗口
window.config(menu=menu_bar)

# 通过应用参数获取令牌
def get_access_token(appkey, appsecret):
    url = 'https://oapi.dingtalk.com/gettoken'
    params = {
        'appkey': appkey,
        'appsecret': appsecret
    }
    res = requests.get(url, params=params)
    return res.json()

# 通过手机获取钉钉端userid
def getUserIdByPhone(mobile,ac_token):
    res = requests.get(
        url='https://oapi.dingtalk.com/user/get_by_mobile',
        params={
            'access_token': ac_token,  #  第二步获取到的access_token
            'mobile': mobile  #  用户电话
        }
    )
    #  如果包含userId键，则返回userId信息，否则返回0。
    if "userid" in res.json():
       return res.json()["userid"]
    else:
       return 0

# 推送钉钉消息API
def send_message(access_token, body):
    url = 'https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2'
    params = {
        'access_token': access_token,
    }
    res = requests.post(url, params=params, json=body)
    return res.json()

# __main__函数须放在mainloop之前，否则不能执行
if __name__ == '__main__':
    #global data_dict  #  引用全局变量data_dict{}
    # 测试推送数据文件是否存在——使用os.path.join()函数拼接当前目录和文件名，生成完整路径
    filepath_xls = os.path.join(os.getcwd(), filename_xls)
    if os.path.exists(filepath_xls):
    #  调用函数，生成推送字典data_dict
        read_exl_file(filepath_xls)
    else:
        messagebox.showinfo('错误', f"请确认推送文件'{filepath_xls}'是否在当前目录下面！")
        exit(100)        # 退出程序，返回错误代码100.

    # 测试背景图片文件是否存在
    filepath_img = os.path.join(os.getcwd(), filename_img)
    if not os.path.exists(filepath_img):
       messagebox.showinfo('错误', f"请确认背景图片'{filepath_img}'是否在secfile目录下面！")
       exit(200)        # 退出程序，返回错误代码200.

#  运行主循环
window.mainloop()