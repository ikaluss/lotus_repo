# -*- coding: utf-8 -*-
import telnetlib

from fastapi import FastAPI
from fastapi import Query
from pydantic import BaseModel
from typing import List
import datetime
import time

import PyExcel

def connectToRouter(hn,un,pw,finish_symbol):
    try:
        tn = telnetlib.Telnet(hn,timeout = 5)
        tn.read_until(b'Username: ')
        tn.write(un.encode('ascii') + b'\n')
        tn.read_until(b'Password: ')
        tn.write(pw.encode('ascii') + b'\n')
        tn.read_until(finish_symbol.encode('ascii'))
        return tn
    except:
        print('Router Connection Failed')
        return -1
    
def disconnectToRouter(tnt):
    try:
        tnt.write(b'exit \n')
    except AttributeError:
        print('Connection closed!')

def getResult(tnn,FinishSymbol):
    try:
        tnn.write(b'sh ip ospf nei' + b'\n')
        return tnn.read_until(FinishSymbol.encode('ascii')).count(b'FULL')
    except:
        return -2


#status: -1:connection failed; -2:get status failed; -3:other errors; 0,1,2:actual status
def checkLineStatus(HostIP,UserName,PassWord,FinishSymbol):
    try:
        tn = connectToRouter(HostIP,UserName,PassWord,FinishSymbol)
        if tn == -1:
            return -1
        test = getResult(tn,FinishSymbol)
        disconnectToRouter(tn)
        tn.close()
        return test
    
    except:
        return -3
        

Host_FIS_CT  = '192.168.231.217'  # 电信FIS路由器
Host_OA_CT   = '192.168.232.117'  # 电信OA路由器
username_CT  = 'cisco'            # 登录用户名  
pw_normal_CT = 'ideal'            # 登录密码

Host_OA_CU   = '192.168.232.217'  # 联通OA路由器
Host_FIS_CU  = '192.168.231.117'  # 联通FIS路由器
username_CU  = 'netstar'
pw_normal_CU = 'netstar'
pw_enable_CU = 'cisco'            # 登录密码

Host_OA_CM   = '192.168.232.17'   # 移动OA路由器
Host_FIS_CM  = '192.168.231.17'   # 移动FIS路由器
username_CM  = 'netstar'
pw_normal_CM = 'netstar'
pw_enable_CM = 'cisco'            # 登录密码

finish_normal = '>'             # 命令提示符（标识着上一条命令已执行完毕）
finish_enable = '#'             # 命令提示符（标识着上一条命令已执行完毕）

flag_OA_CM = 2
flag_OA_CU = 2
flag_OA_CT = 2
flag_FIS_CM = 2
flag_FIS_CU = 2
flag_FIS_CT = 2

standartStatus = 2

def getPreviousStatus():
    try:
        filePath = "lineStatus.xls"
        sheetName = "status"
        getData = PyExcel.ExcelData(filePath, sheetName)
        status = getData.readExcel()
        print(status[0]["flag_OA_CM"],status[0]["flag_OA_CU"],status[0]["flag_OA_CT"],status[0]["flag_FIS_CM"],status[0]["flag_FIS_CU"],status[0]["flag_FIS_CT"])
        return status[0]["flag_OA_CM"],status[0]["flag_OA_CU"],status[0]["flag_OA_CT"],status[0]["flag_FIS_CM"],status[0]["flag_FIS_CU"],status[0]["flag_FIS_CT"]
    except:
        print("get PreviousStatus error")

def setPreviousStatus(param_OA_CM,param_OA_CU,param_OA_CT,param_FIS_CM,param_FIS_CU,param_FIS_CT):
    try:
        filePath = "lineStatus.xls"
        sheetName = "status"
        getData = PyExcel.ExcelData(filePath, sheetName)
        value = [param_OA_CM,param_OA_CU,param_OA_CT,param_FIS_CM,param_FIS_CU,param_FIS_CT]
        getData.overWriteExcel(value)
        print("Set PreviousStatus Success!")
    except:
        print("Set PreviousStatus Error!")


def generateLogs(HostIP,UserName,PassWord,FinishSymbol,HostName,flag):

    temp = checkLineStatus(HostIP,UserName,PassWord,FinishSymbol)
    try:
        if temp != flag:
            data_path = "a.xls"
            sheetname = "专线日志"
            get_data = PyExcel.ExcelData(data_path, sheetname)
        
            now = datetime.datetime.now()
            now.strftime('%Y-%m-%d %H:%M')

            if temp == standartStatus:
                value = [str(now).split('.')[0], "2", HostName, str(temp), HostName + "已恢复"]
                get_data.writeExcel(value,0)
            else:
                value = [str(now).split('.')[0], "2", HostName, str(temp), HostName + "中断"]
                get_data.writeExcel(value,0)
            
            flag = temp
        else:
            print("Don't worry, %s fine!" % HostName)
        
        return flag
    except:
        print("oops")
        return flag
    



def monitorStatus():
    flag_OA_CM, flag_OA_CU, flag_OA_CT, flag_FIS_CM, flag_FIS_CU, flag_FIS_CT = getPreviousStatus()

    flag_OA_CM = generateLogs(Host_OA_CM,username_CM,pw_normal_CM,finish_normal,"移动OA专线",flag_OA_CM)
    flag_OA_CU = generateLogs(Host_OA_CU,username_CM,pw_normal_CM,finish_normal,"联通OA专线",flag_OA_CU)
    flag_OA_CT = generateLogs(Host_OA_CT,username_CT,pw_normal_CT,finish_enable,"电信OA专线",flag_OA_CT)

    flag_FIS_CM = generateLogs(Host_FIS_CM,username_CM,pw_normal_CM,finish_normal,"移动FIS专线",flag_FIS_CM)
    flag_FIS_CU = generateLogs(Host_FIS_CU,username_CM,pw_normal_CM,finish_normal,"联通FIS专线",flag_FIS_CU)
    flag_FIS_CT = generateLogs(Host_FIS_CT,username_CT,pw_normal_CT,finish_enable,"电信FIS专线",flag_FIS_CT)

    setPreviousStatus(flag_OA_CM, flag_OA_CU, flag_OA_CT, flag_FIS_CM, flag_FIS_CU, flag_FIS_CT)


SleepTime   = 30                # 间隔时间，秒

if __name__ == "__main__":
    # print(monitorStatus())
    # getPreviousStatus()
    checkTimes = 1
    while 1 > 0:
        try:
            monitorStatus()
            print('This is the %d times' % checkTimes)
            checkTimes = checkTimes + 1
            time.sleep(SleepTime)
        except:
            print("something wrong happened!")

    

