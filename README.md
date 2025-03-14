这是一个基于钉钉平台的实时消息推送小程序。
应用场景：企业组织中，经常需要同时向组织内不同成员推送不同的实时消息。
    推送依据和实现逻辑：根据手机号找到组织成员在钉钉上的用户ID，再调用API接口，向用户ID推送指定的文本数据。
使用准备：
    1.获取本组织钉钉管理员权限。获取本组织的appkey和appsecret，用以替换代码中的对应部分。
    2.创建H5应用，记录其agent_id，替换代码中对应部分。
文件目录：
    sendmain\dingmsg.py    主程序
    sendmain\minini.py     全局变量
    sendmain\mu.xls        待推送信息
    sendmain\la11.ico      图标
    sendmain\nuitka.txt    非程序运行必须，使用nuitka编译的一些记录
    sendmain\null          程序运行检测网络情况的过程文件
    sendmain\secfile\bg.png    主程序背景图片
    sendmain\secfile\dolog.txt 推送日志
    sendmain\secfile\logxls    推送日志的表格形式
操作说明：下周空了整一个。
