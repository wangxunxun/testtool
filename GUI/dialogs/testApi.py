#coding=utf-8
'''
Created on 2015年7月30日

@author: xun
'''
from PySide import QtGui,QtCore
import os
from tools.ConnectMysql import mysqlconnect
from win32api import ShellExecute
from win32con import SW_SHOW, SW_SHOWNOACTIVATE, SW_SHOWNORMAL
from pymysql.constants.ER import PASSWORD_ANONYMOUS_USER
from tools.testapi import oprMysql,sendAPI,readExcel
import json
from tools.CommonTool import CommonTool


class TestApi(QtGui.QDialog):
    def __init__(self, parent=None):
        super(TestApi, self).__init__(parent)
        self.setWindowTitle(self.trUtf8("TestApi"))
#        self.setWindowFlags(QtCore.Qt.WindowSystemMenuHint)
        self.resize(900, 400)
        
        
        self.host = QtGui.QLabel(self)
        self.host.setText(self.trUtf8("Host"))
        self.hostLineEdit = QtGui.QLineEdit(self)  
        self.hostLineEdit.setText("69.164.202.55")
        
        self.user = QtGui.QLabel(self)
        self.user.setText(self.trUtf8("User"))
        self.userLineEdit = QtGui.QLineEdit(self)  
        self.userLineEdit.setText("test")
        
        self.passwd = QtGui.QLabel(self)
        self.passwd.setText(self.trUtf8("Passwd"))
        self.passwdLineEdit = QtGui.QLineEdit(self)
        self.passwdLineEdit.setEchoMode(QtGui.QLineEdit.Password)

        
        self.db = QtGui.QLabel(self)
        self.db.setText(self.trUtf8("DB"))
        self.dbLineEdit = QtGui.QLineEdit(self)  
        self.dbLineEdit.setText("test")
        
        self.port = QtGui.QLabel(self)
        self.port.setText(self.trUtf8("Port"))
        self.portLineEdit = QtGui.QLineEdit(self)  
        self.portLineEdit.setText("3306")
        
        self.charset = QtGui.QLabel(self)
        self.charset.setText(self.trUtf8("Charset"))
        self.charsetLineEdit = QtGui.QLineEdit(self)  
        self.charsetLineEdit.setText("utf8")
        
        self.output = QtGui.QLabel(self)
        self.output.setText(self.trUtf8("输出文件夹"))
        self.outputLineEdit = QtGui.QLineEdit(self) 
        self.outputLineEdit.setPlaceholderText(u"D:/testcasefolder")
        self.chooseOutPutButton=  QtGui.QPushButton(self.trUtf8("选择"))
        
        self.savename = QtGui.QLabel(self)
        self.savename.setText(self.trUtf8("excel文件名"))
        self.savenameLineEdit = QtGui.QLineEdit(self)  
        
        self.connectButton = QtGui.QPushButton(self.trUtf8("转换为表格"))

        
        self.runButton = QtGui.QPushButton(self.trUtf8("Run"))
        self.jiaochengButton = QtGui.QPushButton(self.trUtf8("常用语句查询"))
        
        
        self.url = QtGui.QLabel(self)
        self.url.setText(self.trUtf8("Url"))
        self.urlLineEdit = QtGui.QLineEdit(self)  
        self.urlLineEdit.setText("http://apis.baidu.com/heweather/weather/free")
#        self.urlLineEdit.setText("http://120.24.255.213:5000/Passenger/User/Regist")

        
        
        self.headers = QtGui.QLabel(self)
        self.headers.setText(self.trUtf8("Headers"))
        self.headersLineEdit = QtGui.QLineEdit(self) 
        self.headersLineEdit.setText('''{"apikey":"761b0c47d570195fbae8125c69d10659"}''') 
        
        self.form_type = QtGui.QRadioButton(self)
        self.form_type.setText(self.trUtf8("Form"))
        self.form_type.setChecked(True)
        self.excel_type = QtGui.QRadioButton(self)
        self.excel_type.setText(self.trUtf8("EXCEL"))
        self.json_type = QtGui.QRadioButton(self)
        self.json_type.setText(self.trUtf8("Json"))
        
        self.buttongroup1 = QtGui.QButtonGroup(self)
        self.buttongroup1.addButton(self.form_type)
        self.buttongroup1.addButton(self.excel_type)
        self.buttongroup1.addButton(self.json_type)
        
        self.get_type = QtGui.QRadioButton(self)
        self.get_type.setText(self.trUtf8("GET"))
        self.get_type.setChecked(True)
        self.post_type = QtGui.QRadioButton(self)
        self.post_type.setText(self.trUtf8("POST"))
        

        
        self.result = QtGui.QLabel(self)
        self.result.setText(self.trUtf8("Result"))
        self.resultTextEdit = QtGui.QTextEdit(self)
        

        self.errorTipLable2 = QtGui.QLabel()      
       


        self.execlname = QtGui.QLabel(self)
        self.execlname.setText(self.trUtf8("数据文件"))
        self.execlnameLineEdit = QtGui.QLineEdit(self)  
        self.execlnameLineEdit.setEnabled(False)
        
        self.chooseExcelButton = QtGui.QPushButton(self.trUtf8("选择"))


        self.sheetname = QtGui.QLabel(self)
        self.sheetname.setText(self.trUtf8("表格名"))
        self.chooseSheet = QtGui.QComboBox()
        self.chooseSheet.addItem(self.trUtf8("请选择"))


        self.script = QtGui.QLabel(self)
        self.script.setText(self.trUtf8("Please input params"))
        self.scriptTextEdit = QtGui.QTextEdit(self)
        self.scriptTextEdit.setText('''{"city":"beijing"}''')
#        self.scriptTextEdit.setText('''{"phoneNumber":"18627802681","password":"1234576"}''')
        scriptlayout = QtGui.QGridLayout()
        scriptlayout.addWidget(self.form_type,0,0)
        scriptlayout.addWidget(self.excel_type,0,1)
        scriptlayout.addWidget(self.json_type,0,2)

        
        scriptlayout.addWidget(self.url,1,0)
        scriptlayout.addWidget(self.urlLineEdit,1,1)
        scriptlayout.addWidget(self.get_type,2,0)
        scriptlayout.addWidget(self.post_type,2,1)
        scriptlayout.addWidget(self.headers,3,0)
        scriptlayout.addWidget(self.headersLineEdit,3,1)        
        
        scriptlayout.addWidget(self.execlname,4,0)
        scriptlayout.addWidget(self.execlnameLineEdit,4,1)        
        scriptlayout.addWidget(self.chooseExcelButton,4,2)
        scriptlayout.addWidget(self.sheetname,5,0)
        scriptlayout.addWidget(self.chooseSheet,5,1)
        
        scriptlayout.addWidget(self.script,6,0)
        scriptlayout.addWidget(self.scriptTextEdit,7,0,1,2)
        scriptlayout.addWidget(self.runButton,8,0)
        scriptlayout.setSpacing(0)


  
        resultlayout = QtGui.QVBoxLayout()
        resultlayout.addWidget(self.result)
        resultlayout.addWidget(self.resultTextEdit)
        resultlayout.addWidget(self.errorTipLable2)
        
        hbox1 = QtGui.QHBoxLayout()
        hbox1.addWidget(self.host)
        hbox1.addWidget(self.hostLineEdit)
        
        hbox2 = QtGui.QHBoxLayout()
        hbox2.addWidget(self.user)
        hbox2.addWidget(self.userLineEdit)
        
        hbox3 = QtGui.QHBoxLayout()
        hbox3.addWidget(self.passwd)
        hbox3.addWidget(self.passwdLineEdit)
        
        hbox4 = QtGui.QHBoxLayout()
        hbox4.addWidget(self.db)
        hbox4.addWidget(self.dbLineEdit)
        
        hbox5 = QtGui.QHBoxLayout()
        hbox5.addWidget(self.port)
        hbox5.addWidget(self.portLineEdit)
        
        hbox6 = QtGui.QHBoxLayout()
        hbox6.addWidget(self.charset)
        hbox6.addWidget(self.charsetLineEdit)
        
        hbox7 = QtGui.QHBoxLayout()
        hbox7.addWidget(self.output)
        hbox7.addWidget(self.outputLineEdit)
        hbox7.addWidget(self.chooseOutPutButton)
        
        hbox8 = QtGui.QHBoxLayout()
        hbox8.addWidget(self.savename)
        hbox8.addWidget(self.savenameLineEdit)
        
        left_layout = QtGui.QVBoxLayout()
        left_layout.addLayout(hbox1)
        left_layout.addLayout(hbox2)
        left_layout.addLayout(hbox3)
        left_layout.addLayout(hbox4)
        left_layout.addLayout(hbox5)
        left_layout.addLayout(hbox6)
        left_layout.addLayout(hbox7)
        left_layout.addLayout(hbox8)
        left_layout.addWidget(self.connectButton)



        
        leftlayout = QtGui.QGridLayout()

        leftlayout.addWidget(self.host,0,0)
        leftlayout.addWidget(self.hostLineEdit,0,1,1,2)
        leftlayout.addWidget(self.user,1,0)
        leftlayout.addWidget(self.userLineEdit,1,1,1,2)
        leftlayout.addWidget(self.passwd,2,0)
        leftlayout.addWidget(self.passwdLineEdit,2,1,1,2)
        leftlayout.addWidget(self.db,3,0)
        leftlayout.addWidget(self.dbLineEdit,3,1,1,2)
        leftlayout.addWidget(self.port,4,0)
        leftlayout.addWidget(self.portLineEdit,4,1,1,2)
        leftlayout.addWidget(self.charset,5,0)
        leftlayout.addWidget(self.charsetLineEdit,5,1,1,2)     
        leftlayout.addWidget(self.output,6,0)   
        leftlayout.addWidget(self.outputLineEdit,6,1)  
        leftlayout.addWidget(self.chooseOutPutButton,6,2)  
        
        leftlayout.addWidget(self.savename,7,0)   
        leftlayout.addWidget(self.savenameLineEdit,7,1,1,2)  

        leftlayout.addWidget(self.connectButton,8,1,1,2)  

     
             
        rightlayout = QtGui.QGridLayout()
        rightlayout.addLayout(scriptlayout,0,0)
        rightlayout.addLayout(resultlayout,1,0)
        rightlayout.setRowStretch(0,10)
        rightlayout.setRowStretch(1,20)
        
        mainlayout = QtGui.QHBoxLayout()
        mainlayout.addLayout(leftlayout)
        mainlayout.addLayout(rightlayout)
        mainlayout.setStretch(0,10)
        mainlayout.setStretch(1,30)

        self.setLayout(mainlayout)
        self.chooseFormType()
        QtCore.QObject.connect(self.connectButton, QtCore.SIGNAL('clicked()'), self.toExcel)
        self.chooseOutPutButton.clicked.connect(self.chooseFolder)
        self.chooseExcelButton.clicked.connect(self.chooseFile)
#        self.connectButton.clicked.connect(self.connectMysql)
        self.runButton.clicked.connect(self.run)
        self.form_type.clicked.connect(self.chooseFormType)
        self.json_type.clicked.connect(self.chooseFormType)
        self.excel_type.clicked.connect(self.chooseExcelType)
#        self.runButton.clicked.emit("canshu") 有参数时需要用此方法发送参数
        self.jiaochengButton.clicked.connect(self.openexcel)
    def toExcel(self):
        
        host = self.hostLineEdit.text()
        user = self.userLineEdit.text()
        passwd = self.passwdLineEdit.text()
        db = self.dbLineEdit.text()
        port = int(self.portLineEdit.text())
        charset = self.charsetLineEdit.text()
        output = self.outputLineEdit.text()
        name = self.savenameLineEdit.text()
        if not host:
            self.resultTextEdit.append(u"host不能为空")

        elif not user:
            self.resultTextEdit.append(u"user不能为空")

        elif not passwd:
            self.resultTextEdit.append(u"passwd不能为空")

        elif not db:
            self.resultTextEdit.append(u"db不能为空")

        elif not port:
            self.resultTextEdit.append(u"port不能为空")

        elif not charset:
            self.resultTextEdit.append(u"charset不能为空")

        elif not output:
            self.resultTextEdit.append(u"output不能为空")

        elif not name:
            self.resultTextEdit.append(u"excel文件名不能为空")

        else:
            try:
                self.oprmysql = oprMysql(host,user,passwd,db,port,charset)
            except Exception as e:
                self.resultTextEdit.append(str(e))
                return
            path = output+"/"+name+".xls"
            try:
                self.oprmysql.toExcel(path)
                
            except Exception as e:
                self.resultTextEdit.append(str(e))
                return
                
            ShellExecute(0,"open",path,"","",SW_SHOW)

                
                
            


    def chooseFolder(self):
        self.dir =QtGui.QFileDialog.getExistingDirectory(self, self.trUtf8("选择文件夹"))
        
        if len(self.dir) == 0:
            {}
        else:

            self.outputLineEdit.setText(self.dir.replace('\\',"/"))        

    def chooseRadio(self):
        if self.get_type.isChecked():
            return "GET"
        if self.post_type.isChecked():
            return "POST"
        
    def chooseRadioGroup1(self):
        if self.form_type.isChecked():
            return "JSON"
        if self.excel_type.isChecked():
            return "EXCEL"
        
    def chooseFormType(self):
        self.execlname.hide()
        self.execlnameLineEdit.hide()
        self.chooseExcelButton.hide()
        self.sheetname.hide()
        self.chooseSheet.hide()
        self.script.show()
        self.scriptTextEdit.show()
    
    def chooseExcelType(self):
        self.execlname.show()
        self.execlnameLineEdit.show()
        self.chooseExcelButton.show()
        self.sheetname.show()
        self.chooseSheet.show()
        self.script.hide()
        self.scriptTextEdit.hide()
 
 
    def chooseFile(self):
        self.file = QtGui.QFileDialog.getOpenFileName(self, self.trUtf8("选择.xls文件"), ".", self.trUtf8("Image Files(*.xls )"))
        if self.file==(u'', u''):
            return;

        else:
            i = 1
            if self.chooseSheet.count() !=1:
                self.chooseSheet.clear()                
                self.chooseSheet.addItem(self.trUtf8("请选择"))
                

            self.execlnameLineEdit.setText(self.file[0])
            sheets =readExcel(self.file[0]).getSheetNames()

            i =0
            while i<len(sheets):
                self.chooseSheet.addItem(sheets[i])
                i=i+1
                       
    def run(self):
        tool = CommonTool()
        self.errorTipLable2.hide()
        jsondata = self.scriptTextEdit.toPlainText()
        url = self.urlLineEdit.text()
        
        headers = self.headersLineEdit.text()
        if headers:
            
            try:
                headers = json.loads(headers)
            except:
                self.errorTipLable2.setText(self.trUtf8("Headers不支持该格式"))
                self.errorTipLable2.show()
                return
            if not isinstance(headers, dict):
                self.errorTipLable2.setText(self.trUtf8("Headers不支持该格式"))
                self.errorTipLable2.show()
                return
        request_type = self.chooseRadio()
        if self.form_type.isChecked():
            if not url:
                self.errorTipLable2.setText(self.trUtf8("url不能为空"))
                self.errorTipLable2.show()
                return
            elif not jsondata:
                self.errorTipLable2.setText(self.trUtf8("参数不能为空"))
                self.errorTipLable2.show()
                return
            else :
                try:
                    newjsondata = json.loads(jsondata)
                except:
                    self.errorTipLable2.setText(self.trUtf8("参数不支持该格式"))
                    self.errorTipLable2.show()
                    return
    
    
            if isinstance(newjsondata, dict):
                sendapi = sendAPI(url,headers,newjsondata,request_type)
                result = sendapi.run()
        
                duration = str(result.get("duration"))
        
                responses = result.get("responses")
                requests = result.get("requests")
                status_codes = result.get("status_codes")
                failCount = str(result.get("failCount"))
                successCount = str(result.get("successCount"))
        
        
                i = 0
                while i <len(requests):



                    req = tool.changeToJson(requests[i])
                    if isinstance(responses[i], unicode):
                        try:
                            res = json.loads(responses[i])
                            res = tool.changeToJson(res)
                        except:
                            
                            res = tool.changeToJson(responses[i])
                    else:                        
                        res = tool.changeToJson(responses[i])

                    self.resultTextEdit.append("request:\n"+req)
                    self.resultTextEdit.append("response:\n"+res)
                    self.resultTextEdit.append("status_code:\n"+str(status_codes[i]))
                
                    
                    self.resultTextEdit.append("-"*100)
                    i = i+1
        
                self.resultTextEdit.append("duration:"+duration+"s")
                self.resultTextEdit.append("failCount:"+failCount)
                self.resultTextEdit.append("successCount:"+successCount)
                self.resultTextEdit.append("-"*100)
                self.resultTextEdit.append("-"*100)
                self.resultTextEdit.moveCursor(QtGui.QTextCursor.End)
    
            elif isinstance(newjsondata, list):
                sendapi = sendAPI(url,headers,newjsondata,request_type)
                result = sendapi.run()
        
                duration = str(result.get("duration"))
        
                responses = result.get("responses")
                requests = result.get("requests")
                failCount = str(result.get("failCount"))
                successCount = str(result.get("successCount"))
                status_codes = result.get("status_codes")
        
        
                i = 0
                while i <len(requests):
                    req = tool.changeToJson(requests[i])

                    if isinstance(responses[i], unicode):
                        try:
                            res = json.loads(responses[i])
                            res = tool.changeToJson(res)
                        except:
                            
                            res = tool.changeToJson(responses[i])
                    else:                        
                        res = tool.changeToJson(responses[i])
                    self.resultTextEdit.append("request:\n"+req)
                    self.resultTextEdit.append("response:\n"+res)
                    self.resultTextEdit.append("status_code:\n"+str(status_codes[i]))
                    self.resultTextEdit.append("-"*100)
                    i = i+1
        
                self.resultTextEdit.append("duration:"+duration+"s")
                self.resultTextEdit.append("failCount:"+failCount)
                self.resultTextEdit.append("successCount:"+successCount)
                self.resultTextEdit.append("-"*100)
                self.resultTextEdit.append("-"*100)
                self.resultTextEdit.moveCursor(QtGui.QTextCursor.End)
    
            else:
                self.errorTipLable2.setText(self.trUtf8("参数不支持该格式"))
                self.errorTipLable2.show()
        elif self.excel_type.isChecked():
            if not self.urlLineEdit.text():
                self.errorTipLable2.setText(self.trUtf8("url不能为空"))
                self.errorTipLable2.show()
            elif not self.execlnameLineEdit.text():
                self.errorTipLable2.setText(self.trUtf8("数据文件不能为空"))
                self.errorTipLable2.show()
            elif self.chooseSheet.currentText()==u"请选择":
                
                self.errorTipLable2.setText(self.trUtf8("表格名不能为空"))
                self.errorTipLable2.show()      
            else:          
                
                
                excel_path = self.execlnameLineEdit.text()
                sheet_name = self.chooseSheet.currentText()
                read_excel = readExcel(excel_path)
                data = read_excel.readTable(sheet_name)
                sendapi = sendAPI(url,headers,data,request_type)
                result = sendapi.run()
                duration = str(result.get("duration"))
        
                responses = result.get("responses")
                requestsdata = result.get("requests")
                failCount = str(result.get("failCount"))
                successCount = str(result.get("successCount"))
                status_codes = result.get("status_codes")
        
        
                i = 0
                while i <len(requestsdata):
                    req = tool.changeToJson(requestsdata[i])
                    if isinstance(responses[i], unicode):
                        try:
                            res = json.loads(responses[i])
                            res = tool.changeToJson(res)
                        except:
                            
                            res = tool.changeToJson(responses[i])
                    else:                        
                        res = tool.changeToJson(responses[i])
                    self.resultTextEdit.append("request:\n"+req)
                    self.resultTextEdit.append("response:\n"+res)
                    self.resultTextEdit.append("status_code:\n"+str(status_codes[i]))
                    self.resultTextEdit.append("-"*100)
                    i = i+1
        
                self.resultTextEdit.append("duration:"+duration+"s")
                self.resultTextEdit.append("failCount:"+failCount)
                self.resultTextEdit.append("successCount:"+successCount)
                self.resultTextEdit.append("-"*100)
                self.resultTextEdit.append("-"*100)
                self.resultTextEdit.moveCursor(QtGui.QTextCursor.End)
        else:
            pass
            
            
        
        
    def stop(self):
        self.mysql.close() 
        self.disConnectButton.setDisabled(True)
        self.connectButton.setEnabled(True)
        self.runButton.setDisabled(True)
        self.resultTextEdit.append("已断开连接")
    def openexcel(self):
        ShellExecute(0,"open",u"Model\mysql操作.xls","","",SW_SHOW)
        
        
        
        
        
        
        
        
        