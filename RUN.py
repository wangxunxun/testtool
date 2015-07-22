#coding=utf-8
'''
Created on 2015年7月22日

@author: xun
'''
import sys

from PySide import QtGui,QtCore
from PySide.QtCore import *
from PySide.QtGui import *
from GUI.mainwindow.mainwindow import MainWindow



def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
    
if __name__ == '__main__':
    main()   
