#coding=utf-8

'''
Created on 2015年7月22日

@author: xun
'''


from PySide import QtGui,QtCore
import os
from tools import exportxml
from blinker.base import Signal
from PySide.QtCore import SIGNAL, SLOT

class toXmlUI2(QtGui.QDialog):
    def __init__(self, parent=None):
        super(toXmlUI2, self).__init__(parent)

        self.setWindowTitle(self.trUtf8("Excel to XML"))
        self.setWindowFlags(QtCore.Qt.WindowSystemMenuHint)

        self.resize(450, 200)
        
        self.execlname = QtGui.QLabel(self)
        self.execlname.setText(self.trUtf8("用例文件"))
        self.execlnameLineEdit = QtGui.QLineEdit(self)  
        self.execlnameLineEdit.setPlaceholderText(u"测试用例地址，格式如:D:/testcase.xls")
        
        self.chooseExcelButton = QtGui.QPushButton(self.trUtf8("选择Excel"))


        
        self.sheetname = QtGui.QLabel(self)
        self.sheetname.setText(self.trUtf8("表格名"))
        self.sheetnameLineEdit = QtGui.QLineEdit(self) 
        self.sheetnameLineEdit.setPlaceholderText(u"表格名称，格式如:Sheet1")
        
        
        self.output = QtGui.QLabel(self)
        self.output.setText(self.trUtf8("输出文件夹"))
        self.outputLineEdit = QtGui.QLineEdit(self) 
        self.outputLineEdit.setPlaceholderText(u"输出文件夹，格式如:D:/testcasefolder")
      
        self.savename = QtGui.QLabel(self)
        self.savename.setText(self.trUtf8("XML文件名"))
        self.savenameLineEdit = QtGui.QLineEdit(self)  
        self.savenameLineEdit.setPlaceholderText(u"生成的XML文件名，格式如：testcase") 
        

        self.okButton = QtGui.QPushButton(self.trUtf8("确定"))
        self.cancelButton =  QtGui.QPushButton(self.trUtf8("取消"))








        self.errorTipLable = QtGui.QLabel()
        self.errorTipLable.setObjectName("tip")




        mainlayout = QtGui.QGridLayout()
        mainlayout.addWidget(self.execlname, 0, 0)
        mainlayout.addWidget(self.execlnameLineEdit, 0, 1)
        mainlayout.addWidget(self.chooseExcelButton, 0, 2)
        mainlayout.addWidget(self.sheetname, 1, 0)
        mainlayout.addWidget(self.sheetnameLineEdit, 1, 1)

        mainlayout.addWidget(self.output, 2, 0)
        mainlayout.addWidget(self.outputLineEdit, 2, 1)
        mainlayout.addWidget(self.savename, 3, 0)
        mainlayout.addWidget(self.savenameLineEdit, 3,1 )


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

        self.chooseExcelButton.clicked.connect(self.open)

    def test(self):
        print 111

    def open(self):
        self.path = QtGui.QFileDialog.getOpenFileName(self, self.trUtf8("Open Image"), ".", self.trUtf8("Image Files(*.xls )"))
        if len(self.path) == 0:
            QtGui.QMessageBox.information(None, self.trUtf8("Path"), self.trUtf8("You didn't select any files.")); 
        else:
            QtGui.QMessageBox.information(None, self.trUtf8("Path"), self.trUtf8("You selected ") + self.path); 
        
    def getFormData(self):

        execlname = self.execlnameLineEdit.text()
        sheetname = self.sheetnameLineEdit.text()
        output = self.outputLineEdit.text()
        savename = self.savenameLineEdit.text()

        if len(execlname) == 0:
            self.errorTipLable.setText(self.tr("FileName is required."))
            self.errorTipLable.show()
        
        if len(sheetname) ==0:
            self.errorTipLable.setText(self.tr("SheetName is required."))
            self.errorTipLable.show()
        
        if len(output) ==0:
            self.errorTipLable.setText(self.tr("Output folder is required."))
            self.errorTipLable.show()    
            
        if len(savename) ==0:
            self.errorTipLable.setText(self.tr("XML Name is required."))
            self.errorTipLable.show()  
        
        if len(execlname) != 0 and len(sheetname) !=0 and len(output) !=0 and len(savename) !=0:
        
            if os.path.exists(execlname):
                try:     
                    sheets = exportxml.exceloperate(execlname).getSheetNames()   
                except:
                    self.errorTipLable.setText(self.tr("There is no sheets."))
                    print 3434
                
                if sheetname in sheets:
                    if os.path.exists(output):
                        aa =exportxml.changetoxml(execlname,sheetname,output,savename)
                        aa.run()
                        self.errorTipLable.setText(self.tr("Succeful."))
                    else:
                        os.mkdir(output)
                        aa =exportxml.changetoxml(execlname,sheetname,output,savename)
                        aa.run()
                        self.errorTipLable.setText(self.tr("Succeful."))
                else:
                    self.errorTipLable.setText(self.tr("The sheet name is not existed."))
        

            
            else:
                self.errorTipLable.setText(self.tr("The excel file is not existed."))