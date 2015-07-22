#coding=utf-8

'''
Created on 2015年7月22日

@author: xun
'''


from PySide import QtGui,QtCore
import os
from tools import exportxml

class toXmlUI(QtGui.QDialog):
    def __init__(self, parent=None):
        super(toXmlUI, self).__init__(parent)

        self.setWindowTitle(self.tr("To XML"))
        self.setWindowFlags(QtCore.Qt.WindowSystemMenuHint)

        self.resize(450, 290)

        self.execlname = QtGui.QLabel(self)
        self.execlname.setText(self.tr("File Name:"))
        self.execlnameLineEdit = QtGui.QLineEdit(self)  
        
        self.sheetname = QtGui.QLabel(self)
        self.sheetname.setText(self.tr("Sheet Name:"))
        self.sheetnameLineEdit = QtGui.QLineEdit(self) 
        
        
        self.output = QtGui.QLabel(self)
        self.output.setText(self.tr("Output folder:"))
        self.outputLineEdit = QtGui.QLineEdit(self) 
      
        self.savename = QtGui.QLabel(self)
        self.savename.setText(self.tr("XML Name:"))
        self.savenameLineEdit = QtGui.QLineEdit(self)   
        

        self.okButton = QtGui.QPushButton(self.tr("OK"))
        self.cancelButton =  QtGui.QPushButton(self.tr("Cancel"))








        self.errorTipLable = QtGui.QLabel()
        self.errorTipLable.setObjectName("tip")
        self.errorTipLable.hide()



        mainlayout = QtGui.QGridLayout()
        mainlayout.addWidget(self.execlname, 0, 0)
        mainlayout.addWidget(self.execlnameLineEdit, 0, 1)
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
                        self.errorTipLable.setText(self.tr("succeful."))
                    else:
                        os.mkdir(output)
                        aa =exportxml.changetoxml(execlname,sheetname,output,savename)
                        aa.run()
                        self.errorTipLable.setText(self.tr("succeful."))
                else:
                    self.errorTipLable.setText(self.tr("The sheet name is not existed."))
        

            
            else:
                self.errorTipLable.setText(self.tr("The excel file is not existed."))
     