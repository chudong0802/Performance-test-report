import requests
import json
import time
from datetime import datetime
future = datetime.strptime('2020-12-31','%Y-%m-%d')
#当前时间
now = datetime.now()
#求时间差
delta = future - now
#企业微信机器人的webhook
send_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=5377ab2b-870a-48af-ae9f-14f0db4bb214"
headers = {"Content-Type": "text/plain"}

#获取每日一句
r = requests.get('http://open.iciba.com/dsapi?').json()
content = r.get('content')
note = r.get('note')


# Markdown类型消息
#标题（支持1至6级标题，注意#与文字中间要有空格）
# 字体颜色(只支持3种内置颜色)绿色：info、灰色：comment、橙红：warning
def send_msg_markdown():
    send_data = {
        "msgtype": "markdown",
        "markdown": {
            "content": "### 距2020年12月31日\n" +
                       "## <font color=\"info\">Bug Free</font> \n"+
                       "# <font color=\"warning\">剩余{}天</font> \n ".format(delta.days) +
                       "\n>####  {} \n".format(note)
        }
    }
    res = requests.post(url=send_url, headers=headers, json=send_data)
    print(res.text)
# Markdown类型消息不能@所有人，所以选择text方式@所有人
 # userid的列表，提醒群中的指定成员(@某个成员)，
# @all表示提醒所有人，如果开发者获取不到userid，可以使用mentioned_mobile_list
def send_msg_txt():
    send_data = {
        "msgtype": "text",
        "text": {
            "mentioned_list": ["@all"]
        }
    }
    res = requests.post(url=send_url, headers=headers, json=send_data)
    print(res.text)

if __name__ == '__main__':
    send_msg_markdown()
    send_msg_txt()

