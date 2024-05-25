import time
import datetime
import requests
import json
import tkinter as tk
import threading
import inspect
import ctypes

LOG_LINE_NUM = 0
uids = {'宋嘉豪': 'UID_h3LqysLXKMtYxkfj51EcUndhuY1K', '李金珂': 'UID_MuvlcijJW3GPCYy8vSFs0zGCx4t3',
        '刘婧琦': 'UID_zK7oZOs783unN1lyxgG4oxGgvrQL', '李晓璇': 'UID_oqvBOHzPHZI1cmdfKljqFgg8wNO2',
        '狄晓宇': 'UID_Tc59wyssm85u3qzop89PFoBg8Ihx', '贾亚锞': 'UID_JSN4h0Rgp859FFv5zZ7aZqXdaRjK'}
WxPusher_uids = []
token = ''
userid = r'918179'
addressid = r'20118719'  # 此项每个人不同
loop_thread = None
closeTime = None


# 主函数
def check_info():
    global token
    global userid
    global addressid
    url = r'http://59.255.61.45/api/api/sms_service/api/sms/list'
    head = {
        "Host": "59.255.61.45",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) gjyj/3.0.1 "
                      "Chrome/85.0.4183.121 Electron/10.1.3 Safari/537.36",
        "Token": token,
        "Areacode": "410503000000000",
        "Userid": userid,
        "Departid": "2253",
        "Addressid": addressid,
        "Isadmin": "0"
    }
    try:
        r = requests.get(url, headers=head, timeout=5).json()
        code = r['code']
        info = r['data']['smsInfoList']
    except Exception as e:
        write_log_to_Text('error: ' + str(e))
        code = 200
        info = []
    return code, info


def get_message_data(check_infoID):
    global token
    url = r'http://59.255.61.45/backapi/acceptAnnouncement.do'
    nowtime = datetime.datetime.now()
    starttime = datetime.datetime.now() - datetime.timedelta(days=30)
    head = {
        "Host": "59.255.61.45",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) gjyj/3.0.1 "
                      "Chrome/85.0.4183.121 Electron/10.1.3 Safari/537.36",
        "Token": token,
    }
    data = {
        "totalCount": 200,
        "curPage": 1,
        "pageSize": 10,
        "loading": "true",
        "announcementAcceptCond.startDate": starttime.strftime('%Y-%m-%d') + ' 00:00:00',
        "announcementAcceptCond.endDate": nowtime.strftime('%Y-%m-%d') + ' 23:59:59'
    }
    try:
        r = requests.post(url, headers=head, data=data, timeout=5).json()
        if r['status'] != 200:
            send_message('status:' + str(r['status']) + 'message:' + str(r['message']))
            annId = 0
            recordId = 0
        else:
            annId = 0
            recordId = 0
            infodata_list = r['data']['annList']
            for i in range(0, 5):
                if (check_infoID in infodata_list[i].values()) and (infodata_list[i]['currentAnnAnnAcceptRecord']['acceptStatus'] == '0'):
                    infodata = infodata_list[i]
                    recordId = infodata['currentAnnAnnAcceptRecord']['recordId']
                    annId = infodata['annID']
                    break
    except Exception as e:
        write_log_to_Text('error: ' + str(e))
        annId = 0
        recordId = 0
    # 接收
    if annId != 0 and recordId != 0:
        url2 = r'http://59.255.61.45/backapi/acceptAnnouncement.do?action=detailAcceptAnnGotoWordTemplate'
        head2 = {
            "Host": "59.255.61.45",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) gjyj/3.0.1 "
                          "Chrome/85.0.4183.121 Electron/10.1.3 Safari/537.36",
            "Token": token,
        }
        data = {
            "ann.annID": annId,
            "annAcceptRecordVO.recordId": recordId
        }
        try:
            r2 = requests.post(url2, headers=head2, data=data, timeout=5).json()
            write_log_to_Text('info: ' + str(r2['data']['logList']))
            send_message("信息接收完成!!")
        except Exception as e:
            write_log_to_Text('error: ' + str(e))
            send_message("信息接收错误!!")
            stop()
    else:
        send_message("接收出错!!")
        stop()


def send_message(message):
    global WxPusher_uids
    url = 'http://wxpusher.zjiecode.com/api/send/message'

    headers = {
        'Content-Type': 'application/json'
    }

    data = {
        'appToken': 'AT_NgmzoOx5z0S9rjJYCnaYJa6mpsa6MirE',
        'content': message,
        # 'summary': '消息摘要', # 消息摘要，显示在微信聊天页面或者模版消息卡片上，限制长度100，可以不传，不传默认截取content前面的内容。
        'contentType': 1,  # 内容类型 1表示文字 2表示html(只发送body标签内部的数据即可，不包括body标签) 3表示markdown
        # 'topicIds': [], # 发送目标的topicId，是一个数组！！！，也就是群发，使用uids单发的时候， 可以不传。
        'uids': WxPusher_uids  # 发送目标的UID，是一个数组。注意uids和topicIds可以同时填写，也可以只填写一个。
        # 'url': "" # 原文链接，可选参数
    }
    requests.post(headers=headers, url=url, data=json.dumps(data), timeout=5)


def _async_raise(tid, exctype):
    """raises the exception, performs cleanup if needed"""
    tid = ctypes.c_long(tid)
    if not inspect.isclass(exctype):
        exctype = type(exctype)
    res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
    if res == 0:
        raise ValueError("invalid thread id")
    elif res != 1:
        # """if it returns a number greater than one, you're in trouble,
        # and you should call it again with exc=NULL to revert the effect"""
        ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
        raise SystemError("PyThreadState_SetAsyncExc failed")


def stop_thread(thread):
    _async_raise(thread.ident, SystemExit)


# 功能函数
# 日志动态打印
def write_log_to_Text(logmsg):
    global LOG_LINE_NUM
    current_time = get_current_time()
    scrollbar_position = log_data_Text.yview()[1]
    logmsg_in = str(current_time) + " " + str(logmsg) + "\n"  # 换行
    if LOG_LINE_NUM <= 7:
        log_data_Text.insert('end', logmsg_in)
        LOG_LINE_NUM = LOG_LINE_NUM + 1
    else:
        log_data_Text.delete(1.0, 2.0)
        log_data_Text.insert('end', logmsg_in)
    # 检查滚动条是否在最底部
    if scrollbar_position == 1.0:
        # 在最底部时，执行自动滚动
        log_data_Text.see(tk.END)
    log_data_Text.update_idletasks()


# 获取当前时间
def get_current_time():
    current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    return current_time


# 复选
def select_option1(name, status):
    if status.get():
        WxPusher_uids.append(uids[name])
        write_log_to_Text('info: ' + name + "已选择")
    else:
        WxPusher_uids.remove(uids[name])
        write_log_to_Text('info: ' + name + "已取消选择")


# 设置关闭时间
def set_closeTime(status):
    global closeTime
    if status.get():
        closeTime = datetime.datetime.now() + datetime.timedelta(days=1)
        closeTime = closeTime.replace(hour=8, minute=0, second=0, microsecond=0)
        write_log_to_Text('info: 在' + closeTime.strftime('%Y-%m-%d %H:%M:%S') + '停止')
    else:
        closeTime = None
        write_log_to_Text('info: 次日停止已取消')


# 禁用输入框和复选
def disabled_element():
    input_text['state'] = 'disabled'
    C1['state'] = 'disabled'
    C2['state'] = 'disabled'
    C3['state'] = 'disabled'
    C4['state'] = 'disabled'
    C5['state'] = 'disabled'
    C6['state'] = 'disabled'
    C7['state'] = 'disabled'


# 启用输入框和复选
def normal_element():
    input_text['state'] = 'normal'
    C1['state'] = 'normal'
    C2['state'] = 'normal'
    C3['state'] = 'normal'
    C4['state'] = 'normal'
    C5['state'] = 'normal'
    C6['state'] = 'normal'
    C7['state'] = 'normal'


# 按钮开关
def switch():
    global loop_thread
    global token
    if input_text.get():
        if CheckVar1.get() or CheckVar2.get() or CheckVar3.get() or CheckVar4.get() or CheckVar5.get() or CheckVar6.get():
            if start_button['text'] == '开始':
                start_button['text'] = '停止'
                token = input_text.get()
                disabled_element()
                loop_thread = threading.Thread(target=start)
                loop_thread.start()
            else:
                stop()
        else:
            write_log_to_Text("error: 推送选择不能为空！")
    else:
        write_log_to_Text("error: token不能为空！")


# 开始程序
def start():
    global closeTime
    global loop_thread
    while True:
        if closeTime:
            if datetime.datetime.now() > closeTime:
                print(2)
                print(loop_thread)
                stop()
        code, info = check_info()
        if code != 200:
            send_message("请重新登录系统！")
            stop()
        if info:
            infoDetail = info[0]
            write_log_to_Text('infoID:' + str(infoDetail['infoID']) + ',  content:' + str(infoDetail['content']))
            send_message('infoID:' + str(infoDetail['infoID']) + ',  content:' + str(infoDetail['content']))
            time.sleep(1)
            get_message_data(int(infoDetail['infoID']))
        write_log_to_Text('info: 程序正在运行！')
        time.sleep(20)


# 停止程序
def stop():
    global loop_thread
    threadID = loop_thread
    start_button['text'] = '开始'
    normal_element()
    write_log_to_Text("info: 程序已停止！")
    loop_thread = None
    stop_thread(threadID)


# 关闭GUI界面
def closeGUI():
    global loop_thread
    if loop_thread:
        stop_thread(loop_thread)
        loop_thread = None
    root.destroy()


# 窗口
root = tk.Tk()
root.title('yjzh')
root.geometry('450x300')
root.resizable(False, False)
# 框架1
frame1 = tk.Frame(root)
frame1.grid(row=0, column=0, sticky='w', padx=5, pady=5)
# token标题
log_label = tk.Label(frame1, text="token:")
log_label.grid(row=0, column=0, sticky='w', padx=10)
# token输入框
input_text = tk.Entry(frame1, show=None)
input_text.grid(row=0, column=1, sticky='w')
# 日志标题
log_label = tk.Label(root, text="推送选择")
log_label.grid(row=1, column=0, sticky='w', padx=10)
# 框架2
frame2 = tk.Frame(root)
frame2.grid(row=2, column=0, sticky='w', padx=5)
CheckVar1 = tk.IntVar()
CheckVar2 = tk.IntVar()
CheckVar3 = tk.IntVar()
CheckVar4 = tk.IntVar()
CheckVar5 = tk.IntVar()
CheckVar6 = tk.IntVar()
C1 = tk.Checkbutton(frame2, text="宋嘉豪", variable=CheckVar1, onvalue=1, offvalue=0,
                    command=lambda: select_option1('宋嘉豪', CheckVar1))
C2 = tk.Checkbutton(frame2, text="李金珂", variable=CheckVar2, onvalue=1, offvalue=0,
                    command=lambda: select_option1('李金珂', CheckVar2))
C3 = tk.Checkbutton(frame2, text="刘婧琦", variable=CheckVar3, onvalue=1, offvalue=0,
                    command=lambda: select_option1('刘婧琦', CheckVar3))
C4 = tk.Checkbutton(frame2, text="李晓璇", variable=CheckVar4, onvalue=1, offvalue=0,
                    command=lambda: select_option1('李晓璇', CheckVar4))
C5 = tk.Checkbutton(frame2, text="狄晓宇", variable=CheckVar5, onvalue=1, offvalue=0,
                    command=lambda: select_option1('狄晓宇', CheckVar5))
C6 = tk.Checkbutton(frame2, text="贾亚锞", variable=CheckVar6, onvalue=1, offvalue=0,
                    command=lambda: select_option1('贾亚锞', CheckVar6))
C1.grid(row=0, column=0, sticky='w', padx=5)
C2.grid(row=0, column=1, sticky='w', padx=5)
C3.grid(row=0, column=2, sticky='w', padx=5)
C4.grid(row=1, column=0, sticky='w', padx=5)
C5.grid(row=1, column=1, sticky='w', padx=5)
C6.grid(row=1, column=2, sticky='w', padx=5)
# 日志标题
log_label = tk.Label(root, text="日志")
log_label.grid(row=4, column=0, sticky='w', padx=10)
# 框架3
frame3 = tk.Frame(root)
frame3.grid(row=5, column=0, sticky='w', padx=5)
# 滚动条
scroll = tk.Scrollbar(frame3)
scroll.pack(side=tk.RIGHT, fill=tk.Y)
# 日志框
log_data_Text = tk.Text(frame3, width=60, height=9)
log_data_Text.pack()
# 两个控件关联
scroll.config(command=log_data_Text.yview)
log_data_Text.config(yscrollcommand=scroll.set)
# 开始/停止
start_button = tk.Button(root, text="开始", bg="lightblue", width=10, command=switch)
start_button.grid(row=3, column=0, sticky='w', padx=10)
# 次日停止
CheckVar7 = tk.IntVar()
C7 = tk.Checkbutton(root, text="次日停止", variable=CheckVar7, onvalue=1, offvalue=0,
                    command=lambda: set_closeTime(CheckVar7))
C7.grid(row=3, column=0, sticky='w', padx=100)
root.protocol("WM_DELETE_WINDOW", closeGUI)  # 更改协议绑定
root.mainloop()
