# -*- coding: utf-8 -*-

"""
@author:随时静听
@file: DocumentTools.py
@time: 2019/04/17
@email:d1314ziting@163.com
"""
from eel import *
import threading
import sqlite3
import json
import os
import win32ui
from win32com.shell import shell,shellcon
import win32gui
import time
from pyexcel_xlsx import get_data
from docxtpl import DocxTemplate
import  math
#全局配置
BASE_DIR=os.path.dirname(os.path.abspath(__file__))
# print BASE_DIR
DB_DIR=os.path.join(BASE_DIR,'db')
DB_FILE=os.path.join(DB_DIR,'data_db.db')
CONF_FILE= os.path.join(os.path.join(BASE_DIR,'CONF'),'app.conf')
#模板管理表
WORD_TPL_SQL='''
CREATE TABLE IF NOT EXISTS template(
id INTEGER PRIMARY KEY AUTOINCREMENT,
title NVARCHAR(64) NOT NULL,
wordtpl NVARCHAR(256) NOT NULL
)
'''

WORD_TPL_PATH=os.path.join(BASE_DIR,'wordTpl')
# print WORD_TPL_PATH


def Num2MoneyFormat( change_number ):
    """
    .转换数字为大写货币格式( format_word.__len__() - 3 + 2位小数 )
    change_number 支持 float, int, long, string
    """
    format_word = [u"分", u"角",u"元",
               u"拾",u"百",u"千",u"万",
               u"拾",u"百",u"千",u"亿",
               u"拾",u"百",u"千",u"万",
               u"拾",u"百",u"千",u"兆"]

    format_num = [u"零",u"壹",u"贰",u"叁",u"肆",u"伍",u"陆",u"柒",u"捌",u"玖"]
    if type( change_number ) == str:
        # - 如果是字符串,先尝试转换成float或int.
        if '.' in change_number:
            try:    change_number = float( change_number )
            except: raise ValueError, '%s   can\'t change'%change_number
        else:
            try:    change_number = int( change_number )
            except: raise ValueError, '%s   can\'t change'%change_number

    if type( change_number ) == float:
        real_numbers = []
        for i in range( len( format_word ) - 3, -3, -1 ):
            if change_number >= 10 ** i or i < 1:
                real_numbers.append( int( round( change_number/( 10**i ), 2)%10 ) )

    elif isinstance( change_number, (int, long) ):
        real_numbers = [ int( i ) for i in str( change_number ) + '00' ]

    else:
        raise ValueError, '%s   can\'t change'%change_number

    zflag = 0                       #标记连续0次数，以删除万字，或适时插入零字
    start = len(real_numbers) - 3
    change_words = []
    for i in range(start, -3, -1):  #使i对应实际位数，负数为角分
        if 0 <> real_numbers[start-i] or len(change_words) == 0:
            if zflag:
                change_words.append(format_num[0])
                zflag = 0
            change_words.append( format_num[ real_numbers[ start - i ] ] )
            change_words.append(format_word[i+2])

        elif 0 == i or (0 == i%4 and zflag < 3):    #控制 万/元
            change_words.append(format_word[i+2])
            zflag = 0
        else:
            zflag += 1

    if change_words[-1] not in ( format_word[0], format_word[1]):
        # - 最后两位非"角,分"则补"整"
        change_words.append(u"整")

    return ''.join(change_words)

def numToRMB(num):
    if not num:
        return '0.00'

    num=str(num)
    ret=num.find('.')
    if ret!=-1:
        num2=num[:ret]
    else:
        num2=num
    n=len(num2)/3
    r=len(num2)%3
    # print r
    if r==0:
        n=n-1
    # print n
    result=""
    if n>0:#30000
        if r!=0:
            result=num2[:r]
            num2=num2[r:]
        else:
            result=num2[:3]
            num2=num2[3:]

        for i in range(n):
            result+=','+num2[ i*3:(i+1)*3]


        return result+".00"
    else:
        return  num2+".00"






def BrowseCallbackProc(hwnd, msg, lp, data):
    if msg== shellcon.BFFM_INITIALIZED:
        win32gui.SendMessage(hwnd, shellcon.BFFM_SETSELECTION, 1, data)
    elif msg == shellcon.BFFM_SELCHANGED:
        # Set the status text of the
        # For this message, 'lp' is the address of the PIDL.
        pidl = shell.AddressAsPIDL(lp)
        try:
            path = shell.SHGetPathFromIDList(pidl)
            # print path
            win32gui.SendMessage(hwnd, shellcon.BFFM_SETSTATUSTEXT, 0, path)
        except shell.error:
            pass
#  保存配置文件
def saveConf(confDic,conf=CONF_FILE):
    if not conf:
        return

    if not os.path.exists(os.path.dirname(conf)):
        os.makedirs(os.path.dirname(conf))
    with open(conf,'w') as f:
        json.dump(confDic,f)
        return  True

# 读取EXCEL文件中的数据
def read_from_xls(filename):
    if not os.path.exists(filename):
        return
    data=get_data(filename)

    return  data





# 渲染word模板生成word文件并写入到配置的导出目录
def renderWord(excelLineData,wordTpl,reportName,v):
    if not excelLineData :
        return
    if not wordTpl:
        return
    n=7-len(excelLineData)
    if n>0:
        excelLineData.extend([""]*n)
    # print excelLineData[5]
    dataDic={
        'bidder':excelLineData[0],#投标人
        'beneficiary':excelLineData[1],#受益人
        'bidName':excelLineData[2],#投标项目名称
        'projectNo':excelLineData[3],#工程编号
        'expenses':numToRMB(excelLineData[4]),#担保金额
        'expensesCN':Num2MoneyFormat(excelLineData[4]),#担保金额 大写
        'expiryBidDateY':str(excelLineData[5])[:4],#截标日期 年
        'expiryBidDateM':str(excelLineData[5])[5:7],#截标日期 月
        'expiryBidDateD':str(excelLineData[5])[8:10],#截标日期 日
        'letterDateY':str(excelLineData[6])[:4],#出函日期 年
        'letterDateM':str(excelLineData[6])[5:7],#出函日期 月
        'letterDateD':str(excelLineData[6])[8:10],#出函日期 日
        'letterDate':str(excelLineData[6]),#出函日期
    }
    doc = DocxTemplate(wordTpl)
    doc.render(dataDic)
    # print reportName
    doc.save(reportName)

    progressBar(v)












# 获取app 配置文件

@expose
def loadAppConf():
    reportPath =BASE_DIR
    conf = CONF_FILE
    appConf={u'reportPath':reportPath}

    if not os.path.exists(conf):
        return  appConf
    with open(conf,'r') as f:
        fappconf=json.load(f)

        if fappconf:

            appConf.update(fappconf)

    return appConf
@expose
def startRun(filename,tplIdArr):
    if not filename:
        return False
    reportPath=loadAppConf().get('reportPath','')
    if not reportPath:
        AlertInfo("请配置路径导出位置！","alert")
    if not os.path.exists(filename):
        AlertInfo("选择的EXCEL文件不存在！", "alert")
    excelData=read_from_xls(filename)
    for key,item in excelData.items():
        lines=item
        break
    #     去除首行

    linesall=lines[1:]
    lines=[]
    for line in linesall:
        if line:
            lines.append(line)

    # for i, line in enumerate(lines):
    #     print i,
    #     if len(line) > 0:
    #         print line[0]

    tplLst = []
    conn = getConn()
    cursor = conn.cursor()
    sql = "select * from  template WHERE id=?"

    for id in tplIdArr:
        cursor.execute(sql, (id,))
        ret = cursor.fetchone()
        if ret:
            tplLst.append(ret)
    cursor.close()
    conn.close()
    # print tplLst
    thread_lst=[]
    v=1.0/(len(lines)*len(tplLst))
    v=math.ceil(v*100)
    for i,line in enumerate(lines):
        for y, tpl in enumerate(tplLst):
            # print line
            # print i
            reportName=tpl[1]+"_"+str(i)+str(y)+'.docx'
            reportName=os.path.join(reportPath,reportName)
            t=threading.Thread(target=renderWord,args=(line,tpl[2],reportName,v))
            thread_lst.append(t)
    for t in thread_lst:
        t.start()
    for t in thread_lst:
        t.join()

    return True


# print shell.SHGetPathFromIDList(result[0])
def openDirDialog(path=""):
    wnd=win32gui.FindWindow(None,u"文档转转工具")
    # print wnd
    if not path:
        path=""
    if not os.path.exists(path):
        pidl=shell.SHBrowseForFolder(wnd,None,"Plese select Path")[0]

    else:
        desktop = shell.SHGetDesktopFolder()
        cb, pidl, extra = desktop.ParseDisplayName(0, None, path)
        pidl=shell.SHBrowseForFolder(wnd,None,"Plese select Path")[0]
    try:
        selectPath= shell.SHGetPathFromIDList(pidl)
    except:
        selectPath=path
    try:
        saveConf({"reportPath":selectPath.decode('gbk')})
    except:
        saveConf({"reportPath": selectPath})

    return  selectPath

init('web/')

@expose
def getUserInfo():
    '''
    将作者信息暴漏出去
    :return:
    '''
    userinfo={
        "username":"随时静听",
        "email":"d1314ziting@163.com",
    }
    return json.dumps(userinfo)
@expose
def getVersion():
    '''
    程序版本信息暴漏出去
    :return:
    '''
    version={
        "version":"Version 1.0",
        "copyright":"随时静听 [d1314ziting@163.com]",
    }
    return  json.dumps(version)

def moveTpl(filename):
    msg={'status':0}

    if not os.path.exists(WORD_TPL_PATH):

        os.makedirs(WORD_TPL_PATH)
    if not filename:

        return msg# 文件存在
    # print filename
    if os.path.exists(filename):

        name=os.path.split(filename)[1]
        oldname,ext=os.path.splitext(name)
        if ext=="" or not (ext =='.doc' or ext==".docx"):
            msg['status']=1#文件后缀不合法
            return msg
        newfilename=oldname+time.strftime("%Y%m%d_%H%M%S",time.localtime())+ext
        newfilename=time.strftime("%Y%m%d_%H%M%S",time.localtime())+ext
        filepath=os.path.join(WORD_TPL_PATH,newfilename)
        with open(filename,'rb') as f:
            with open(filepath,'wb') as fw:
                fw.write(f.read())
                msg['status']=3# 文件写入成功
                msg['path']=filepath
    return msg

@expose
def openFile():
    # wnd = win32gui.FindWindow(None, u"文档转转工具")
    dlg=win32ui.CreateFileDialog(1)
    dlg.DoModal()
    filename=dlg.GetPathName()
    # print filename
    try:
        return filename.decode('gbk')
    except:
        return  ""

'''
def moveTpl11(filename,data):
    msg={'status':0}
    print 9999999
    if not data:
        print 6666
        return msg
    if not filename:
        print 1444
        return msg#

    oldname,ext=os.path.splitext(filename)
    print ext
    if ext=="" or not (ext =='.doc' or ext==".docx"):
        msg['status']=1#文件后缀不合法
        return msg
    newfilename=oldname+time.strftime("%Y%m%d_%H%M%S",time.localtime())+ext
    filepath=os.path.join(WORD_TPL_PATH,newfilename)

    with open(filepath,'wb') as fw:
        fw.write(data.encode('utf-8'))
        msg['status']=3# 文件写入成功
        msg['path']=filepath


    return msg
'''







@expose
def opendir(path):
    # dlg=win32ui.CreateFileDialog(0)
    # flag=dlg.DoModal()
    # print dlg.GetPathName()
    pass

    result=openDirDialog(path)

    try:
        return result.decode('gbk')
    except:
        return  result





'''
需要数据表结构存储Excel 转换成word的对应关系

//如果开发成高扩展性和使用性更强的工具，需要费时，但是可以做到任意Excel数据到word填充，可批量定制
1. 建立规则（ Excel和word模板指定）
2. Excel确定那些列的数据填充到word，并设置Excel到word关联的字段变量，动态添加和删除
3. 手动修改word模板，然后上传，其他高级使用直接在程序中拖动或者选择的方式来处理

//高级点的暂时先不写，先写一个简单点的，也就是可以添加Excel，Word映射的并可以修改Excel或者word,Excel其实是模板名称

模板管理表
id integer auto_increament primary key, 模板id
title nvarchar(64) not null,模板名称
wordtpl nvarchar(512),word模板路径

//模板管理表建立的sql语句
CREATE TABLE IF NOT EXISTS template(
id INTERGER PRIMARY KEY AUTOINCREMENT,
title NVARCHAR(64) NOT NULL,
wordtpl NVARCHAR(256) NOT NULL,
)
'''





#数据库文件和目录初始化

def getConn():
    conn=None
    #判断数据库路径是否存在，不在就创建
    if not os.path.exists(DB_DIR):
        try:
            os.makedirs(DB_DIR)
        except:
            pass
            #在这里通知客户端环境初始化失败

            exit()#程序退出
    #判断数据库文件是存在
    conn=sqlite3.connect(DB_FILE)

    return  conn

def initDB(conn,sql=WORD_TPL_SQL):
    if not conn or  not sql:
        #游标无效，sql为空或者none 就退出
        exit()

    cursor=conn.cursor()
    cursor.execute(sql)
    cursor.close()


# 环境初始化操作
conn=getConn()
# 数据初始化
initDB(conn)
conn.close()
# 数据加载
@expose
def loadData():
    conn=getConn()
    cursor=conn.cursor()
    sql="select * from template"
    cursor.execute(sql)
    return cursor.fetchall()


# 添加到数据库
@expose
def insertDB(name,fname):
    # print name,fname
    if not name:
        return
    if not fname:
        return
    ret=moveTpl(fname)
    newfile=ret.get("path","")
    if not newfile:
        return
    sql="INSERT INTO template(title,wordtpl) VALUES (?,?)"
    conn=getConn()
    cursor=conn.cursor()
    cursor.execute(sql,(name,newfile))
    conn.commit()
    conn.close()
    return True

@expose
def DeleteData(id):
    if not id:
        return False
    conn = getConn()
    cursor = conn.cursor()

    sql = "select * from template WHERE id=?"
    cursor.execute(sql, id)
    filepath=cursor.fetchone()
    if filepath:
        filepath=filepath[2]
    if os.path.join(filepath):
        if os.path.exists(filepath):
            os.remove(filepath)
    sql="DELETE from template WHERE id=?"

    cursor.execute(sql,id)
    conn.commit()
    conn.close()
    return True

@expose
def editData(id,name,filePath):
    if not id:
        return False
    if not filePath:
        return False
    sql="select * from template WHERE id=?"
    conn=getConn()
    cursor=conn.cursor()
    cursor.execute(sql,(id,))
    item=cursor.fetchone()
    if item[2]==filePath and name==item[1]:
        return True
    ret=moveTpl(filePath)
    # print ret
    newfile=ret.get('path')
    sql="UPDATE template set title=? , wordtpl=? WHERE id=?"
    cursor.execute(sql,(name,newfile,id))
    cursor.close()
    conn.commit()
    conn.close()
    return  True


@expose
def insertDB1(name,tplpath):
    # print name,tplpath
    if not name:

        return
    if not tplpath:

        return
    ret=moveTpl(tplpath)

    newfile=ret.get("path","")
    if not newfile:

        return
    sql="INSERT INTO template(title,wordtpl) VALUES (?,?)"
    conn=getConn()
    cursor=conn.cursor()
    cursor.execute(sql,(name,newfile))
    cursor.close()

    return True



# start('main.html')
start('templates/index.html',size=(1100,700),templates='templates')
if __name__ == '__main__':
    pass