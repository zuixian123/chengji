# -*- codeing = utf-8 -*-
# @Time : 2022/9/20 19:57
# @Author : 张智成
# @File : 有不及格测试.py
# @Software : PyCharm


import time
import os
from selenium import webdriver, common
import re
import random
from time import localtime
from requests import get, post
from datetime import datetime, date
from zhdate import ZhDate
import sys
def get_access_token():
    # appId
    app_id = "wxbbb8a76f47df69f0"
    # appSecret
    app_secret = "b47e21ea34934df47fc58ad3cd573cf6"
    post_url = ("https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={}&secret={}"
                .format(app_id, app_secret))
    try:
        access_token = get(post_url).json()['access_token']
    except KeyError:
        print("获取access_token失败，请检查app_id和app_secret是否正确")
        os.system("pause")
        sys.exit(1)
    # print(access_token)
    return access_token

def send_message(to_user, access_token, text1, text2):
    url = "https://api.weixin.qq.com/cgi-bin/message/template/send?access_token={}".format(access_token)
    week_list = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
    year = localtime().tm_year
    month = localtime().tm_mon
    day = localtime().tm_mday
    today = datetime.date(datetime(year=year, month=month, day=day))
    week = week_list[today.isoweekday() % 7]
    # 获取所有生日数据
    birthdays = {}
    for k, v in config.items():
        if k[0:5] == "birth":
            birthdays[k] = v
    data = {
        "touser": to_user,
        "template_id": config["template_id"],
        "url": "http://weixin.qq.com/download",
        "topcolor": "#FF0000",
        "data": {
            "date": {
                "value": "{} {}".format(today, week),
            },
            "text1": {
                "value": text1
            },
            "text2": {
                "value": text2
            }
        }
    }
    headers = {
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                      'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36'
    }
    response = post(url, headers=headers, json=data).json()
    if response["errcode"] == 40037:
        print("推送消息失败，请检查模板id是否正确")
    elif response["errcode"] == 40036:
        print("推送消息失败，请检查模板id是否为空")
    elif response["errcode"] == 40003:
        print("推送消息失败，请检查微信号是否正确")
    elif response["errcode"] == 0:
        print("推送消息成功")
    else:
        print(response)

if __name__ == "__main__":
    try:
        with open("config.txt", encoding="utf-8") as f:
            config = eval(f.read())
    except FileNotFoundError:
        print("推送消息失败，请检查config.txt文件是否与程序位于同一路径")
        os.system("pause")
        sys.exit(1)
    except SyntaxError:
        print("推送消息失败，请检查配置文件格式是否正确")
        os.system("pause")
        sys.exit(1)

    # 获取accessToken
    accessToken = get_access_token()
    # 接收的用户
    users = config["user"]

    # 环境配置
    chromedriver = "C:\Program Files (x86)\Google\Chrome\Application"
    os.environ["webdriver.ie.driver"] = chromedriver


    driver = webdriver.Edge()  # 选择Chrome浏览器
    driver.get('http://jwglxt.zua.edu.cn/eams/loginExt.action')  # 打开网站

    time.sleep(1)
    # driver.find_element_by_link_text('登录').click() # 点击“账户登录”

    # username = "2107211021"  # 请替换成你的用户名
    # password = "Liuhang123."  # 请替换成你的密码
    username = config["username"]
    password = config["password"]

    driver.find_element_by_id('username').click()  # 点击用户名输入框
    driver.find_element_by_id('username').clear()  # 清空输入框
    driver.find_element_by_id('username').send_keys(username)  # 自动敲入用户名

    driver.find_element_by_id('password').click()  # 点击密码输入框
    driver.find_element_by_id('password').clear()  # 清空输入框
    driver.find_element_by_id('password').send_keys(password)  # 自动敲入密码

    # 采用class定位登陆按钮
    driver.find_element_by_class_name('submit').click() # 点击“登录”按钮
    # 采用xpath定位登陆按钮，
    # driver.find_element_by_xpath('//*[@id="root"]/div/div[3]/form/button').click()

    time.sleep(1)

    # driver.find_element_by_id('signIn').click() # 点击“签到”

    # win32api.keybd_event(122, 0, 0, 0)  # F11的键位码是122，按F11浏览器全屏
    # win32api.keybd_event(122, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    driver.get('http://jwglxt.zua.edu.cn/eams/teach/grade/course/person!search.action?semesterId=117&projectType=')
    text = driver.page_source
    driver.close()
    # print(text)
    a = re.findall('%">(.*?)</th>', text)
    b = re.findall('"c?o?l?o?r?:?r?e?d?">|<td>(.*?\s*)[\s<]', text)
    c = re.findall('</td><td style=".*?">(.*?)\n', text)
    d = re.findall('</td><td>(.*?)\n',text)
    score = []
    classname = []
    credit = []
    grade = []
    m = len(b)//12
    for i in range(0,m):
        name = b[i*12+3].strip('\t').strip('\n').strip('\t')
        xf = b[i*12+5]
        gra = d[i].strip('\t')
        if gra == "":
            gra = '暂无'
        sco = c[i].strip('\t')
        classname.append(name)
        credit.append(xf)
        grade.append(gra)
        score.append(sco)

    file = open("chengji.txt", "w", encoding="utf-8")
    file.close()

    with open("chengji.txt", "a", encoding="utf-8")as w:
        for i in range(0, m):
            w.write( classname[i] + " " + '成绩：' + score[i] + " " + '绩点：' + grade[
                i] + '\n\n')
    with open("chengji.txt", "r", encoding="utf-8")as w:
        text1 = w.read()
    # driver.close()
        # 公众号推送消息
    print("您的成绩如下：")
    print(text1)
    # exit(0)
    text2 = ""
    for user in users:
        send_message(user, accessToken, text1, text2)
    os.system("pause")

    # 代码调用：
    # python.exe JDSignup.py
    # 可以将这行命令添加到Windows计划任务，每天运行，从而实现每日自动签到领取京豆。



