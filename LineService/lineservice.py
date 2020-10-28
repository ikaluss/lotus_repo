# -*- coding: utf-8 -*-
import telnetlib

from fastapi import FastAPI
from fastapi import Query
from pydantic import BaseModel
from typing import List

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

BandWidth_OA_CM = '8Mbps'
BandWidth_OA_CU = '40Mbps'
BandWidth_OA_CT = '40Mbps'
BandWidth_FIS_CM = '4Mbps'
BandWidth_FIS_CU = '16Mbps'
BandWidth_FIS_CT = '16Mbps'

LineLogsFilePath = u'Z:\guotao\LineLogs\LineLogs.csv'
##LineLogsFilePath = u'/test/LineLogs.csv'


app = FastAPI()

class Item(BaseModel):
    name: str
    price: float
    is_offer: bool = None



@app.get("/getStatusAll/")
def getStatusAll():
    status_OA_CM = checkLineStatus(Host_OA_CM,username_CM,pw_normal_CM,finish_normal)
    status_OA_CU = checkLineStatus(Host_OA_CU,username_CM,pw_normal_CM,finish_normal)
    status_OA_CT = checkLineStatus(Host_OA_CT,username_CT,pw_normal_CT,finish_enable)

    status_FIS_CM = checkLineStatus(Host_FIS_CM,username_CM,pw_normal_CM,finish_normal)
    status_FIS_CU = checkLineStatus(Host_FIS_CU,username_CM,pw_normal_CM,finish_normal)
    status_FIS_CT = checkLineStatus(Host_FIS_CT,username_CT,pw_normal_CT,finish_enable)
    return [{"name":"移动OA","status":status_OA_CM,'bandwidth':BandWidth_OA_CM},{"name":"联通OA","status":status_OA_CU,'bandwidth':BandWidth_OA_CU},{"name":"电信OA","status":status_OA_CT,'bandwidth':BandWidth_OA_CT},{"name":"移动FIS","status":status_FIS_CM,'bandwidth':BandWidth_FIS_CM},{"name":"联通FIS","status":status_FIS_CU,'bandwidth':BandWidth_FIS_CU},{"name":"电信FIS","status":status_FIS_CT,'bandwidth':BandWidth_FIS_CT}]

@app.get("/getLineLogs/")
def getLineLogs():
    
    result = []
    f = open(LineLogsFilePath,'r+')
    
    for line in f:
        (startTime,startComment,endTime,engineer,isProblemSolved,method,backup) = line.split(',')
        line = {}
        line['startTime'] = startTime
        line['startComment'] = startComment
        line['endTime'] = endTime
        line['isProblemSolved'] = isProblemSolved
##        result.append(line)
        result.insert(0,line)
        
    return result

@app.get("/getLineLogsRealTime/")
def getLineLogsRealTime():
    
    data_path = r'a.xls'
    sheetname = "专线日志"
    get_data = PyExcel.ExcelData(data_path, sheetname)
    result = get_data.readExcel()
    # print(result)
    result.reverse()
    return result

@app.get("/getBM/")
def getBM():
    data_path = r'Z:\liuying\BM单\运维和项目跟踪CIT-5.xlsx'
    sheetname = "BM跟踪表"
    get_data = PyExcel.ExcelData(data_path, sheetname)
    result = get_data.readExcel()
##    print(result)
    return result





##result = getLineLogs()
##print(result)


# print(getLineLogsRealTime())
##status = checkLineStatus(Host_OA_CM,username_CM,pw_normal_CM,finish_normal)
##print(status)
##status = checkLineStatus(Host_OA_CU,username_CM,pw_normal_CM,finish_normal)
##print(status)
##status = checkLineStatus(Host_OA_CT,username_CT,pw_normal_CT,finish_enable)
##print(status)
##
##status = checkLineStatus(Host_FIS_CM,username_CM,pw_normal_CM,finish_normal)
##print(status)
##status = checkLineStatus(Host_FIS_CU,username_CM,pw_normal_CM,finish_normal)
##print(status)
##status = checkLineStatus(Host_FIS_CT,username_CT,pw_normal_CT,finish_enable)
##print(status)
