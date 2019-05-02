#!/usr/bin/env python
#-*- coding: utf-8 -*-

import requests
import re,os,sys
import urllib
import getpass
from lxml import etree
from PIL import Image
from bs4 import BeautifulSoup

import icode

class Who:
    def __init__(self, user, pswd):
        self.user = user
        self.pswd = pswd

def Getgrade(response):
    html = response.content
    soup = BeautifulSoup(html, 'lxml')
    trs = soup.find(id="Datagrid1").findAll("tr")
    Grades = []
    keys = []
    tds = trs[0].findAll("td")
    tds = tds[:2] + tds[3:5] + tds[6:9]
    for td in tds:
        keys.append(td.string)
    for tr in trs[1:]:
        tds = tr.findAll("td")
        tds = tds[:2] + tds[3:5] + tds[6:9]
        values = []
        for td in tds:
            values.append(td.string)
        one = dict((key, value) for key, value in zip(keys, values))
        Grades.append(one)
    return Grades


class University:
    def __init__(self, student, baseurl):
        # reload(sys)
        self.student = student
        self.baseurl = baseurl
        self.session = requests.session()
        self.session.headers['User-Agent'] = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'

    def Login(self):
        url = self.baseurl+'/default2.aspx'
        res = self.session.get(url)
        cont = res.content
        selector = etree.HTML(cont)
        __VIEWSTATE = selector.xpath('//*[@id="form1"]/input/@value')[0]
        imgurl = self.baseurl + '/CheckCode.aspx'
        imgres = self.session.get(imgurl, stream=True)
        img = imgres.content
        with open('code.jpg', 'wb') as f:
            f.write(img)

        code=icode.verfication('code.jpg')

        # jpg = Image.open('{}/code.jpg'.format(os.getcwd()))
        # jpg.show()
        # jpg.close
        # code = input('输入验证码：')

        RadioButtonList1 = u"学生"
        data = {
            "__VIEWSTATE": __VIEWSTATE,
            "txtUserName": self.student.user,
            "TextBox1": self.student.pswd,
            "TextBox2": self.student.pswd,
            "txtSecretCode": code,
            "RadioButtonList1": RadioButtonList1,
            "Button1": "",
            "lbLanguage": ""
        }
        loginres = self.session.post(url, data=data)
        logcont = loginres.text
        pattern = re.compile(
            '<form name="Form1".*?action=(.*?) id="Form1">', re.S)
        res = re.findall(pattern, logcont)
        try:
            if res[0][17:29] == self.student.user:
                print('Login succeed!')
        except:
            print('Login failed! Maybe Wrong password ! ! !')
            return
        pattern = re.compile('<span id="xhxm">(.*?)</span>')
        xhxm = re.findall(pattern, logcont)
        name = xhxm[0].replace('同学', '')
        self.student.urlname = urllib.parse.quote_plus(str(name))
        return True

    def GetGrade(self):
        self.session.headers['Referer'] = self.baseurl + '/xs_main.aspx?xh=' + self.student.user
        gradeurl = self.baseurl + '/xscjcx.aspx?xh='+self.student.user + '&xm='+self.student.urlname+'&gnmkdm=N121605'
        graderesponse = self.session.get(gradeurl)

        gradecont = graderesponse.content
        soup = BeautifulSoup(gradecont, 'lxml')
        __VIEWSTATE = soup.findAll(name="input")[2]["value"]
        self.session.headers['Referer'] = gradeurl
        data = {
            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",
            "__VIEWSTATE": __VIEWSTATE,
            "hidLanguage": "",
            "ddlXN": "",
            "ddlXQ": "",
            "ddl_kcxz": "",
            "btn_zcj": u'历年成绩'
        }
        grares = self.session.post(gradeurl, data=data)
        grades = Getgrade(grares)

####################写入xls#######################
        import xlwt
        file = xlwt.Workbook(encoding='utf-8')        # 指定file以utf-8的格式打开
        sheet = file.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

        style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
        al = xlwt.Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al

        sheet.col(0).width = 256 * 40
        lists=list(grades[0])      #写入头
        num=[2,4,6]              #输出的值
        for i in range(len(num)):
            sheet.write(0,i,lists[num[i]],style)

        line=1
        for dict in grades:
            for i in range(len(num)):
                sheet.write(line,i,dict[lists[num[i]]],style)
            line+=1

        file.save(user+'.xls')  # 保存并打开
        os.system('start '+user+'.xls')
        os.system('del code.jpg')

if __name__ == "__main__":
    url = 'http://202.206.243.3'
    user=input('学号：')
    print('密码不可见')
    pswd=getpass.getpass('password:')
    #user = '2018xxxxxxxx'
    #pswd = "xxxxxxxxxxxx"

    who = Who(user, pswd)
    univ = University(who, url)
    if univ.Login():
        univ.GetGrade()
