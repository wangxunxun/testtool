#coding=utf-8
'''
Created on 2015年7月22日

@author: xun
'''
import sys  
  
from cx_Freeze import setup, Executable  
from distutils.core import setup  
import py2exe 



#python mysetup.py py2exe
# Setup
setup ( 
       name = "Test TOOL",
       description = 'Some Test including Glade, Python and GTK in win32',     
       version = '1.0', 
       
       
       windows = [{
                        'script': 'Testtool.py'
                  }],
       data_files=[("Model",
                   ["Model/Test Case.xls","Model/Test Case Without Step.xls","Model/全国中小企业股转系统官网前台_Test Case_v1.3.xls",
                    "Model/QQ音乐_Android V3.6.1.9_Normal Test Result_Beta.xls",
                    "Model/mysql操作.xls"])]
      )
