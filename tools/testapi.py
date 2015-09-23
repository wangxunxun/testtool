#coding=utf-8
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

import httplib2
import json
import xlwt
import sqlite3
import xlrd
from dominate.tags import table
import datetime
import pymysql
import requests

def zp(r):
    try:
        print(json.dumps(json.loads(r.text), indent=4,ensure_ascii=False, encoding="utf-8"))
        return json.dumps(json.loads(r.text), indent=4,ensure_ascii=False, encoding="utf-8")
    except Exception, e:
        print r.text

class oprMysql:
    def __init__(self,host,user,passwd,db,port,charset):
        self.host = host
        self.user = user
        self.passwd = passwd
        self.db = db
        self.charset = charset
        self.port = port
        self.conn=pymysql.connect(host = self.host,user = self.user,passwd = self.passwd,db = self.db,port = self.port,charset = self.charset)
        self.cur = self.conn.cursor()
            
    def getTablesName(self):
        self.cur.execute('show tables')
        result = self.cur.fetchall()
        tablesname = []
        i = 0
        while i <len(result):
            tablesname.append(result[i][0])
            i = i+1
        return tablesname
    
    def getTableHeader(self,tablename):
        self.cur.execute('describe %s'%tablename)
        result = self.cur.fetchall()
        header = []
        for i in result:
            header.append(i[0])       
        return header
    
    def getTableData(self,tablename):
        self.cur.execute('select * from %s'%tablename)
        data = self.cur.fetchall()
        return data
    
    def getTablesHeader(self,tablesname):
        data = {}
        for tablename in tablesname:
            self.cur.execute('describe %s'%tablename)
            tables = self.cur.fetchall()
            header = []
            for i in tables:
                header.append(i[0])   
            data.setdefault(tablename,header)                              
        return data
    
    def getTablesData(self,tablesname):
        data = {}
        for tablename in tablesname:
            self.cur.execute('select * from %s'%tablename)
            tabledata = self.cur.fetchall()
            data.setdefault(tablename,tabledata)
        return data
    
    def closeConnect(self):                  
        self.cur.close()
        self.conn.commit()  
        self.conn.close()
        
    def toExcel(self,excelpath):     
        tablesname = self.getTablesName()
        tablesheader = self.getTablesHeader(tablesname)
        tablesdata = self.getTablesData(tablesname)
        self.closeConnect()
        excel = oprexcel(excelpath)
        excel.saveTables(tablesname, tablesheader, tablesdata)
    

        
        
class oprsqlite:
    def __init__(self,sqlpath):
        self.sqlpath = sqlpath
        self.cx = sqlite3.connect(sqlpath)
        self.cu = self.cx.cursor()

    def getTablesName(self):
        self.cu.execute('select name from sqlite_master where type = "table" order by name')        
        tables = self.cu.fetchall()
        tablesname = []
        i = 0
        while i <len(tables):
            tablesname.append(tables[i][0])
            i = i+1
        return tablesname
    
    def getTableHeader(self,tablename):
        self.cu.execute('PRAGMA table_info(%s)'%tablename)
        tables = self.cu.fetchall()
        header = []
        for i in tables:
            header.append(i[1])          
        return header
    
    def getTableData(self,tablename):
        self.cu.execute('select * from %s'%tablename)
        data = self.cu.fetchall()
        return data
    
    def getTablesHeader(self,tablesname):
        data = {}
        for tablename in tablesname:
            self.cur.execute('describe %s'%tablename)
            tables = self.cur.fetchall()
            header = []
            for i in tables:
                header.append(i[0])   
            data.setdefault(tablename,header)                                   
        return data
    
    def getTablesData(self,tablesname):
        data = {}
        for tablename in tablesname:
            self.cu.execute('select * from %s'%tablename)
            tabledata = self.cu.fetchall()
            data.setdefault(tablename,tabledata)
        return data
    
    def closeConnect(self):
        self.cx.close()

    def toExcel(self,excelpath):     
        tablesname = self.getTablesName()
        tablesheader = self.getTablesHeader(tablesname)
        tablesdata = self.getTablesData(tablesname)
        self.closeConnect()
        excel = oprexcel(excelpath)
        excel.saveTables(tablesname, tablesheader, tablesdata)          
              
class oprexcel:
    def __init__(self,excelpath):
        self.excelpath = excelpath
    
    def saveTable(self,tablename,tableheader,tabledata):
        f = xlwt.Workbook()
        sheet1 = f.add_sheet(tablename,cell_overwrite_ok=True) 
        for i in range(0,len(tableheader)):
            sheet1.write(0,i,tableheader[i])
        i = 0
        while i < len(tabledata):
            for j in range(0,len(tabledata[i])):
                sheet1.write(i+1,j,str(tabledata[i][j]).decode('UTF-8'))
            i = i+1
        
        f.save(self.excelpath)
        
    def saveTables(self,tablesname,tablesheader,tablesdata):
        f = xlwt.Workbook()
        for tablename in tablesname:
            print(tablename)
            sheet1 = f.add_sheet(tablename,cell_overwrite_ok=True) 
            for i in range(0,len(tablesheader.get(tablename))):
                sheet1.write(0,i,tablesheader.get(tablename)[i])
            i = 0
            while i < len(tablesdata.get(tablename)):
                for j in range(0,len(tablesdata.get(tablename)[i])):
                    sheet1.write(i+1,j,str(tablesdata.get(tablename)[i][j]).decode('UTF-8'))
                i = i+1       
        
        f.save(self.excelpath)
        print('Save successfully')

        
class readExcel():
    def __init__(self,excelpath):
        self.excelpath = excelpath
        self.data = xlrd.open_workbook(self.excelpath)
        
    def getSheetNames(self):
        return self.data.sheet_names()
        
    def readTable(self,tablename):

        table = self.data.sheet_by_name(tablename)
        nrows = table.nrows
        ncols = table.ncols
        header = table.row_values(0)
        tabledata = []
        i = 1
        while i < nrows:
            rdata = table.row_values(i)
            if rdata[0]:
                if isinstance(rdata[0],float):
                    rdata[0] = int(rdata[0])
                j = 0
                row = {}
                while j <len(header):                         
                    row.setdefault(header[j],rdata[j])
                    j = j+1   
                tabledata.append(row)
            i=i+1
        return tabledata
    
    def readTables(self,tablesname):
        tablesdata = {}
        for tablename in tablesname:
            table = self.data.sheet_by_name(tablename)
            nrows = table.nrows
            ncols = table.ncols
            header = table.row_values(0)
            tabledata = []
            i = 1
            while i < nrows:
                rdata = table.row_values(i)
                if rdata[0]:    
                    if isinstance(rdata[0],float):
                        rdata[0] = int(rdata[0])
                    j = 0
                    row = {}
                    while j <len(header):                
                        row.setdefault(header[j],rdata[j])
                        j = j+1           
                    tabledata.append(row)
                i=i+1
            tablesdata.setdefault(tablename,tabledata)
        return tablesdata

class sendAPI:
    def __init__(self,url,data,method='POST',contentType = 'json'):
        self.url =url
        self.data = data
        self.method = method
        self.contentType = contentType 
        
    def run(self):
        if self.method == 'POST' and self.contentType == 'json':
            successCount = 0
            failCount =0
            result = {}
            requestsdata = []
            responses = []
            response = {}
            start = datetime.datetime.now()
            if isinstance(self.data, list):
                
                for i in self.data:
                    requestsdata.append(i)
                    r = requests.post(self.url, data=i)   

                    response.setdefault("text",r.text)
                    response.setdefault("status_code",r.status_code)                
                    responses.append(response)


                    
                    if r.status_code==200:
                        successCount = successCount +1
                    else:
                        failCount = failCount +1
            elif isinstance(self.data, dict): 
                requestsdata.append(self.data)
                r = requests.post(self.url, data=self.data)   

                response.setdefault("text",r.text)
                response.setdefault("status_code",r.status_code) 
            
                responses.append(response)
                
                if r.status_code==200:
                    successCount = successCount +1
                else:
                    failCount = failCount +1
            else:
                return u"json格式不正确"
            end = datetime.datetime.now()
            duration = end - start
            duration =  duration.seconds +float(duration.microseconds)/1000000
            result.setdefault('duration',duration)
            result.setdefault('successCount',successCount)
            result.setdefault('failCount',failCount)
            result.setdefault("requests",requestsdata)
            result.setdefault("responses",responses)
            print(result)
            return result
        if self.method == 'GET' and self.contentType == 'json':
            successCount = 0
            failCount =0
            result = {}
            requestsdata = []
            responses = []
            response = {}
            start = datetime.datetime.now()
            if isinstance(self.data, list):
                
                for i in self.data:
                    requestsdata.append(i)
                    r = requests.get(self.url, params=i)   
                    response.setdefault("text",r.text)
                    response.setdefault("status_code",r.status_code)                
                    responses.append(response)


                    
                    if r.status_code==200:
                        successCount = successCount +1
                    else:
                        failCount = failCount +1
            elif isinstance(self.data, dict): 
                requestsdata.append(self.data)
                r = requests.get(self.url, params=self.data)   

                response.setdefault("text",r.text)
                response.setdefault("status_code",r.status_code)                
                responses.append(response)
                
                if r.status_code==200:
                    successCount = successCount +1
                else:
                    failCount = failCount +1
            else:
                return u"json格式不正确"
            end = datetime.datetime.now()
            duration = end - start
            duration =  duration.seconds +float(duration.microseconds)/1000000
            result.setdefault('duration',duration)
            result.setdefault('successCount',successCount)
            result.setdefault('failCount',failCount)
            result.setdefault("requests",requestsdata)
            result.setdefault("responses",responses)
            print(result)
            return result
        else:
            print('no support')
            
 
                    
        


if __name__ == '__main__':
    
#    userdata = [{"version_number":"115541","download_url":"34343434343"}]
#    userdata =[{"phoneNumber":"18627802682","password":"123456"}]
#    userdata =[{"phoneNumber":"18627802681","password":"1234576"}]
    userdata =[{"phoneNumber":"18627802681"},{"phoneNumber":"18627802681"}]
#    userdata =[{"email":"59853844@qq.com","password":"111111"}]
#    userdata = [{"sessionID":"cb8857de5da711e5869b001a7dda7113"}]
    cc = sendAPI("http://127.0.0.1:5000/api/v1.0/User/EmitCode",userdata,method="GET")
#    cc = sendAPI("http://127.0.0.1:5000/api/v1.0/User/Logout",userdata,method="GET")
#    cc = sendAPI("http://127.0.0.1:5000/api/v1.0/User/Login",userdata,method="POST")
#    cc = sendAPI("http://127.0.0.1:5000/Passenger/User/Regist",userdata,method="POST")
    cc.run()
#    cc.run()
#    a =oprMysql('69.164.202.55','test','test','test',3306,"utf8")

#    a.toExcel('D:/demo.xls')
    
#    b = readExcel('D:/demo.xls')


#    print(content)