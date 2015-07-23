#coding=utf-8
'''
Created on 2015年7月22日

@author: xun
'''

from PySide import QtGui,QtCore
import GUI.dialogs.toxml
from PySide.QtCore import *
from PySide.QtGui import *
from GUI.dialogs.toxml import toXmlUI
from GUI.dialogs.toxml2 import toXmlUI2

class MainWindow(QtGui.QMainWindow):    
    def __init__(self):
        super(MainWindow, self).__init__()        
        self.initUI()
        
    def initUI(self):        

        self.statusBar().showMessage('statusbar:troy test')
        menubar = self.menuBar()
        toolMenu = menubar.addMenu(u'&工具')
        helpmenu = menubar.addMenu(u'&帮助')
#        toXmlAction = self.createAction(u'&Testlink Excel To Xml', self.toXml)
        toXmlAction2 = self.createAction(u'&Testlink Excel To Xml', self.toXml2)
        aboutUsAction = self.createAction(u'&关于我们',self.test)
        helpmenu.addAction(aboutUsAction)
#        toolMenu.addAction(toXmlAction)
        toolMenu.addAction(toXmlAction2)
             
        self.setGeometry(300, 300, 500, 350)
        self.setWindowTitle(u'测试工具')            
        self.show()

    def test(self):
        self.statusBar().showMessage('troy test',9000)   


    def toXml(self):
        dialog = toXmlUI()
        dialog.exec_()
        
    def toXml2(self):
        dialog = toXmlUI2()
        dialog.exec_()
        
          
    def closeEvent(self, event):        
        reply = QtGui.QMessageBox.question(self, 'Message',
            u"Are you sure to quit?", QtGui.QMessageBox.Yes | 
            QtGui.QMessageBox.No, QtGui.QMessageBox.No)

        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore() 
    

        
    def createAction(self,text,slot=None,shortcut=None, icon=None,
               tip=None,checkable=False,signal="triggered()"):
        action = QAction(text, self)
        if icon is not None:
            action.setIcon(QIcon("./images/%s.png" % icon))
        if shortcut is not None:
            action.setShortcut(shortcut)
        if tip is not None:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        if slot is not None:
            self.connect(action, SIGNAL(signal), slot)
        if checkable:
            action.setCheckable(True)
        return action
        