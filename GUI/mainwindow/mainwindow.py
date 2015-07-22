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


class MainWindow(QtGui.QMainWindow):    
    def __init__(self):
        super(MainWindow, self).__init__()        
        self.initUI()
        
    def initUI(self):        

        self.statusBar().showMessage('statusbar:Ready')
        menubar = self.menuBar()
        toolMenu = menubar.addMenu('&Tool')
        helpmenu = menubar.addMenu('&Help')
        toXmlAction = self.createAction('&ToXml', self.toXml)
        aboutUsAction = self.createAction(u'&关于我们',self.test)
        helpmenu.addAction(aboutUsAction)
        toolMenu.addAction(toXmlAction)
             
        self.setGeometry(300, 300, 500, 350)
        self.setWindowTitle('Test Tools')            
        self.show()

    def test(self):
        self.statusBar().showMessage('You have created a new file!',9000)   


    def toXml(self):
        dialog = toXmlUI()
        dialog.exec_()
          

    

        
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
        