#coding=utf-8
'''
Created on 2015年7月22日

@author: xun
'''
import sys  
  
from cx_Freeze import setup, Executable  
from distutils.core import setup  
import py2exe  
''' 
options ={ 'py2exe':
                {
                    'dll_excludes':['w9xpopen.exe'] #This file is for win9x platform
                }
        }
'''
# Setup
setup ( 
       name = "Test TOOL",
       description = 'Some Test including Glade, Python and GTK in win32',     
       version = '1.0', 
       
       
       windows = [{
                        'script': 'RUN.py'
                  }],
       data_files=[("Model",
                   ["Model/TestCase.xls"])]
      )
'''
setup(console=["RUN.py"],
      data_files=[("Model",
                   ["Model/TestCase.xls"])]
)
'''