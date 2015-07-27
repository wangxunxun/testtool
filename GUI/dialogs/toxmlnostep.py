#coding=utf-8

'''
Created on 2015年7月22日

@author: xun
'''


from PySide import QtGui,QtCore
import os
from tools import ExportXmlNoStep
from tools import exportxml
from win32api import ShellExecute
from win32con import SW_SHOW, SW_SHOWNOACTIVATE, SW_SHOWNORMAL


class toXmlUI(QtGui.QDialog):
    def __init__(self, parent=None):
        super(toXmlUI, self).__init__(parent)

        self.setWindowTitle(self.trUtf8("Excel to XML"))
        self.setWindowFlags(QtCore.Qt.WindowSystemMenuHint)

        self.resize(450, 200)
        
        self.execlname = QtGui.QLabel(self)
        self.execlname.setText(self.trUtf8("用例文件"))
        self.execlnameLineEdit = QtGui.QLineEdit(self)  
        self.execlnameLineEdit.setPlaceholderText(u"D:/testcase.xls")
        
        self.chooseExcelButton = QtGui.QPushButton(self.trUtf8("选择"))
        self.caseModelButton = QtGui.QPushButton(self.trUtf8("用例模板"))



        
        self.sheetname = QtGui.QLabel(self)
        self.sheetname.setText(self.trUtf8("表格名"))
#        self.sheetnameLineEdit = QtGui.QLineEdit(self) 
#        self.sheetnameLineEdit.setPlaceholderText(u"表格名称，格式如:Sheet1")
        
        
        self.output = QtGui.QLabel(self)
        self.output.setText(self.trUtf8("输出文件夹"))
        self.outputLineEdit = QtGui.QLineEdit(self) 
        self.outputLineEdit.setPlaceholderText(u"D:/testcasefolder")
        self.chooseOutPutButton=  QtGui.QPushButton(self.trUtf8("选择"))
        
      
        self.savename = QtGui.QLabel(self)
        self.savename.setText(self.trUtf8("XML文件名"))
        self.savenameLineEdit = QtGui.QLineEdit(self)  
        self.savenameLineEdit.setPlaceholderText(u"testcase") 
        

        self.okButton = QtGui.QPushButton(self.trUtf8("确定"))
        self.cancelButton =  QtGui.QPushButton(self.trUtf8("取消"))


        self.chooseSheet = QtGui.QComboBox()
        self.chooseSheet.addItem(self.trUtf8("请选择"))





        self.errorTipLable = QtGui.QLabel()
        self.errorTipLable.setObjectName("tip")




        mainlayout = QtGui.QGridLayout()
        mainlayout.addWidget(self.execlname, 0, 0)
        mainlayout.addWidget(self.execlnameLineEdit, 0, 1)
        mainlayout.addWidget(self.chooseExcelButton, 0, 2)
        mainlayout.addWidget(self.caseModelButton,0,3)
        mainlayout.addWidget(self.sheetname, 1, 0)
#        mainlayout.addWidget(self.sheetnameLineEdit, 1, 1)
        mainlayout.addWidget(self.chooseSheet,1,1)

        mainlayout.addWidget(self.output, 2, 0)
        mainlayout.addWidget(self.outputLineEdit, 2, 1)
        mainlayout.addWidget(self.savename, 3, 0)
        mainlayout.addWidget(self.savenameLineEdit, 3,1 )
        mainlayout.addWidget(self.chooseOutPutButton,2,2)


        mainlayout.addWidget(self.errorTipLable, 4, 1)


        okLayout = QtGui.QHBoxLayout()
        okLayout.addStretch()
        okLayout.addWidget(self.okButton)
        okLayout.addSpacing(75)
        okLayout.addWidget(self.cancelButton)
        okLayout.addSpacing(80)

        mainlayout.addLayout(okLayout, 5, 0, 1, 4)

        self.setLayout(mainlayout)
        self.okButton.clicked.connect(self.getFormData)
        self.cancelButton.clicked.connect(self.reject)

        self.chooseExcelButton.clicked.connect(self.chooseFile)
        self.chooseOutPutButton.clicked.connect(self.chooseFolder)
        self.caseModelButton.clicked.connect(self.openexcel)
        
      

    def openexcel(self):
        ShellExecute(0,"open","Model\Test Case Without Step.xls","","",SW_SHOW)

        

        
    def test(self):
        print 111
        
    def chooseFolder(self):
        self.dir =QtGui.QFileDialog.getExistingDirectory(self, self.trUtf8("选择文件夹"))
        
        if len(self.dir) == 0:
            {}
        else:

            self.outputLineEdit.setText(self.dir.replace('\\',"/"))

        

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
            sheets = ExportXmlNoStep.exceloperate(self.file[0]).getSheetNames()
            i =0
            while i<len(sheets):
                self.chooseSheet.addItem(sheets[i])
                i=i+1
        
    def getFormData(self):


        execlname = self.execlnameLineEdit.text()
        sheetname = self.chooseSheet.currentText()
        output = self.outputLineEdit.text()
        savename = self.savenameLineEdit.text()

        if len(execlname) == 0:
            self.errorTipLable.setText(self.trUtf8("用例文件不能为空"))
            self.errorTipLable.show()
        
        elif len(sheetname) ==0:
            self.errorTipLable.setText(self.trUtf8("表格名不能为空"))
            self.errorTipLable.show()
        
        elif len(output) ==0:
            self.errorTipLable.setText(self.trUtf8("输出文件夹不能为空"))
            self.errorTipLable.show()    
            
        elif len(savename) ==0:
            self.errorTipLable.setText(self.trUtf8("XML文件名不能为空"))
            self.errorTipLable.show()  
        elif len(execlname) >= 100:
            self.errorTipLable.setText(self.trUtf8("用例文件最大长度为100"))
            self.errorTipLable.show()
        elif len(sheetname) >= 100:
            self.errorTipLable.setText(self.trUtf8("表格名最大长度为100"))
            self.errorTipLable.show()
        elif len(output) >= 100:
            self.errorTipLable.setText(self.trUtf8("输入文件夹最大长度为100"))
            self.errorTipLable.show()            
        elif len(savename) >= 100:
            self.errorTipLable.setText(self.trUtf8("XML文件名最大长度为100"))
            self.errorTipLable.show()            
            
                    
        elif len(execlname) != 0 and len(sheetname) !=0 and len(output) !=0 and len(savename) !=0:
        
            if os.path.exists(execlname):
                try:     
                    sheets = ExportXmlNoStep.exceloperate(execlname).getSheetNames()   
                except:
                    self.errorTipLable.setText(self.trUtf8("该用例文件没有表格"))
                
                if sheetname in sheets:
                    if os.path.exists(output):
                        
                        xmlfile = output.replace('/',"\\")+"\\"+savename+".xml"
                        aa =ExportXmlNoStep.changetoxml(execlname,sheetname,output,savename)  
                        try:                      
                            aa.run()
                            ShellExecute(0,"open",xmlfile,"","",SW_SHOWNOACTIVATE)
                            self.errorTipLable.setText(self.trUtf8("成功转换成XML文件"))
                        except:
                            self.errorTipLable.setText(self.trUtf8("请参照用例模板设计用例"))
                        
                    else:
                        os.mkdir(output)
                        xmlfile = output.replace('/',"\\")+"\\"+savename+".xml"
                        print xmlfile
                        aa =ExportXmlNoStep.changetoxml(execlname,sheetname,output,savename)
                        try:                      
                            aa.run()
                            ShellExecute(0,"open",xmlfile,"","",SW_SHOWNOACTIVATE)
                            self.errorTipLable.setText(self.trUtf8("成功转换成XML文件"))
                        except:
                            self.errorTipLable.setText(self.trUtf8("请参照用例模板设计用例"))
                else:
                    self.errorTipLable.setText(self.trUtf8("表格名不存在"))
                    
            else:
                self.errorTipLable.setText(self.trUtf8("用例文件不存在"))