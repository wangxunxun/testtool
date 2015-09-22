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


class MysqlUI(QtGui.QDialog):
    def __init__(self, parent=None):
        super(MysqlUI, self).__init__(parent)
        self.setWindowTitle(self.trUtf8("Mysql Tool"))
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
        
        self.connectButton = QtGui.QPushButton(self.trUtf8("Connect"))
        self.disConnectButton = QtGui.QPushButton(self.trUtf8("Disconnect"))
        self.disConnectButton.setDisabled(True)
        self.runButton = QtGui.QPushButton(self.trUtf8("Run"))
        self.runButton.setDisabled(True)
        self.jiaochengButton = QtGui.QPushButton(self.trUtf8("常用语句查询"))
        
        self.script = QtGui.QLabel(self)
        self.script.setText(self.trUtf8("Please input script"))
        self.scriptTextEdit = QtGui.QTextEdit(self)
        
        self.result = QtGui.QLabel(self)
        self.result.setText(self.trUtf8("Result"))
        self.resultTextEdit = QtGui.QTextEdit(self)

        
       
        
        buttonlayout = QtGui.QHBoxLayout()
        buttonlayout.addWidget(self.connectButton)
        buttonlayout.addWidget(self.disConnectButton)
       
        scriptlayout = QtGui.QGridLayout()
        scriptlayout.addWidget(self.script,0,0)
        scriptlayout.addWidget(self.scriptTextEdit,1,0,1,2)
        scriptlayout.addWidget(self.runButton,2,0)
        scriptlayout.addWidget(self.jiaochengButton,2,1)
        
        resultlayout = QtGui.QVBoxLayout()
        resultlayout.addWidget(self.result)
        resultlayout.addWidget(self.resultTextEdit)
        
        

        leftlayout = QtGui.QGridLayout()

        leftlayout.addWidget(self.host,0,0)
        leftlayout.addWidget(self.hostLineEdit,0,1)
        leftlayout.addWidget(self.user,1,0)
        leftlayout.addWidget(self.userLineEdit,1,1)
        leftlayout.addWidget(self.passwd,2,0)
        leftlayout.addWidget(self.passwdLineEdit,2,1)
        leftlayout.addWidget(self.db,3,0)
        leftlayout.addWidget(self.dbLineEdit,3,1)
        leftlayout.addWidget(self.port,4,0)
        leftlayout.addWidget(self.portLineEdit,4,1)
        leftlayout.addWidget(self.charset,5,0)
        leftlayout.addWidget(self.charsetLineEdit,5,1)        
        leftlayout.addLayout(buttonlayout,6,1)
     
             
        rightlayout = QtGui.QGridLayout()
        rightlayout.addLayout(scriptlayout,0,0)
        rightlayout.addLayout(resultlayout,1,0)
        rightlayout.setRowStretch(0,10)
        rightlayout.setRowStretch(1,30)
        
        mainlayout = QtGui.QHBoxLayout()
        mainlayout.addLayout(leftlayout)
        mainlayout.addLayout(rightlayout)
        mainlayout.setStretch(0,10)
        mainlayout.setStretch(1,30)

        self.setLayout(mainlayout)
        QtCore.QObject.connect(self.connectButton, QtCore.SIGNAL('clicked()'), self.connectMysql)
        
#        self.connectButton.clicked.connect(self.connectMysql)
        self.runButton.clicked.connect(self.run)
#        self.runButton.clicked.emit("canshu") 有参数时需要用此方法发送参数
        self.disConnectButton.clicked.connect(self.stop)
        self.jiaochengButton.clicked.connect(self.openexcel)
    def connectMysql(self):
        host = self.hostLineEdit.text()
        user = self.userLineEdit.text()
        passwd = self.passwdLineEdit.text()
        db = self.dbLineEdit.text()
        port = int(self.portLineEdit.text())
        
        charset = self.charsetLineEdit.text()
        self.mysql = mysqlconnect(host,user,passwd,db,port,charset)
        result = self.mysql.connect()
        if  isinstance(result, str):
            self.resultTextEdit.append("error:"+result)
            self.resultTextEdit.moveCursor(QtGui.QTextCursor.End)
            
        else:    
            self.disConnectButton.setEnabled(True)
            self.connectButton.setDisabled(True)
            self.runButton.setEnabled(True)
            self.resultTextEdit.append("连接成功")
            self.resultTextEdit.moveCursor(QtGui.QTextCursor.End)
        


        
    def run(self):
        script = self.scriptTextEdit.toPlainText()
        result = self.mysql.execute(script)
        self.resultTextEdit.append("Execute:" +script)
        if  isinstance(result, str):
            self.resultTextEdit.append("error:"+result)
            self.resultTextEdit.moveCursor(QtGui.QTextCursor.End)
        else:
            self.resultTextEdit.append("返回结果的数量："+ str(result[0]))
            self.resultTextEdit.append("返回结果：" + str(result[1]))
            self.resultTextEdit.moveCursor(QtGui.QTextCursor.End)
        
        
    def stop(self):
        self.mysql.close() 
        self.disConnectButton.setDisabled(True)
        self.connectButton.setEnabled(True)
        self.runButton.setDisabled(True)
        self.resultTextEdit.append("已断开连接")
    def openexcel(self):
        ShellExecute(0,"open",u"Model\mysql操作.xls","","",SW_SHOW)
        
        
        
        
        
        
        
        
        