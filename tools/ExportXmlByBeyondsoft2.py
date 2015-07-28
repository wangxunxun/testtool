#coding=utf-8
'''
Created on 2015年2月28日

@author: xun
'''
#coding=utf-8
from xml.dom import minidom
import xlrd
import os
    
class ExportXmlNoStep:
    def __init__(self,testdata,outputfolder,filename):
        self.testdata = testdata
        self.output = outputfolder
        self.filename = filename
        
    def export(self):
        impl = minidom.getDOMImplementation()
        dom = impl.createDocument(None, None, None)
        inittestsuite = dom.createElement("testsuite")
        inittestsuite.setAttribute("name", "")   
        i=0
        testsuite_s =[]
        while i<len(self.testdata):                        
            testsuite =  dom.createElement("testsuite")                    
            testsuite_s.append(testsuite)    
            testsuite_s[i].setAttribute("name", self.testdata[i].get("testsuite").encode('utf-8'))   
            testcase_s=[]
            summary_s =[]
            preconditions_s =[]
            importance_s =[]          
            j=0
            while j<len(self.testdata[i].get("testcases")):        
                testcase = dom.createElement("testcase")
                testcase_s.append(testcase)
                summary = dom.createElement("summary")
                summary_s.append(summary)
                preconditions = dom.createElement("preconditions")
                preconditions_s.append(preconditions)
                importance = dom.createElement("importance")
                importance_s.append(importance)                                        
                testcase_s[j].setAttribute("name", self.testdata[i].get("testcases")[j].encode('utf-8'))                                                                                    
                summary_text = dom.createTextNode(self.testdata[i].get("summary")[j].encode('utf-8'))
                precondition_text = dom.createTextNode(self.testdata[i].get("precondition")[j].encode('utf-8'))
                importance_text = dom.createTextNode(self.testdata[i].get("importance")[j].encode('utf-8'))            
                summary_s[j].appendChild(summary_text)                
                preconditions_s[j].appendChild(precondition_text)                              
                importance_s[j].appendChild(importance_text)


                testcase_s[j].appendChild(summary_s[j])
                testcase_s[j].appendChild(preconditions_s[j])
                testcase_s[j].appendChild(importance_s[j])            

                                                     
                testsuite_s[i].appendChild(testcase_s[j])                                
                j=j+1        
            inittestsuite.appendChild(testsuite_s[i])                
            i=i+1    
        dom.appendChild(inittestsuite)
        f=file(self.output+"/"+self.filename+".xml",'w')
        dom.writexml(f,'',' ','\n','utf-8')
        f.close()

class readexcel:
    def __init__(self,testexcel,sheetname):
        self.testexcel = testexcel   
        self.sheetname = sheetname
        self.data = xlrd.open_workbook(self.testexcel)
        self.table = self.data.sheet_by_name(self.sheetname)    
        
    def getsheetname(self):
        return self.sheetname
        
    def read(self):
        self.testdata = []
        self.testsuites = []
        self.testcases =[]
        self.testsite =[]
        self.result =[]
        self.summary =[]
        self.preconditon =[]
        self.importance =[]
        i =1
        j =self.table.nrows
        while i<j:

            testsuite = {}
            if self.table.cell(i,0).value and self.table.cell(i,2).value:
                testsuite.setdefault("testsuite",unicode(self.table.cell(i,0).value))
                self.testcases.append(unicode(self.table.cell(i,2).value))
                self.preconditon.append(unicode(self.table.cell(i,4).value))
                self.importance.append(unicode(self.table.cell(i,7).value))
                self.testsite.append(unicode(self.table.cell(i,5).value))
                self.result.append(unicode(self.table.cell(i,6).value))
                self.testsuites.append(testsuite)
                
                
                
            elif not self.table.cell(i,0).value and self.table.cell(i,2).value:

                
                self.testcases.append(unicode(self.table.cell(i,2).value))
                self.preconditon.append(unicode(self.table.cell(i,4).value))
                self.importance.append(unicode(self.table.cell(i,7).value))
                self.testsite.append(unicode(self.table.cell(i,5).value))
                self.result.append(unicode(self.table.cell(i,6).value))
            i=i+1
        self.testsite = self.newline(self.testsite)
        self.result = self.newline(self.result)
        self.summary = self.getSummary(self.testsite, self.result)

        self.testdata.append(self.testsuites)
        self.testdata.append(self.testcases)
        self.testdata.append(self.preconditon)
        self.testdata.append(self.importance)
        self.testdata.append(self.summary)
        return self.testdata

    def newline(self,data):
        i = 0
        newdata = []
        while i<len(data):
            c = data[i].replace("\n","<br>")
            newdata.append(c)
            i = i+1
        return newdata            
            
    def getSummary(self,testsite,result):
        i = 0
        summary = []
        while i <len(self.testcases):
            t = testsite[i]
            r = result[i]
            site = "<span style='font-weight:bold;font-size:18px;color:#ee82ee;'>"+"Steps (Input):"+"</span>"+"<br>"
            re = "<span style='font-weight:bold;font-size:18px;color:#ee82ee;'>"+"Expected Output:"+"</span>"+"<br>"
            summary.append(site+t +"<br>"*3+ re +r)
            i=  i+1
        return summary
        
    def suitedis(self):        
        j=1
        dis =[]        
        while j<self.table.nrows:
            if self.table.cell(j,0).value:
                dis.append(j)
            j=j+1
        dis.append(self.table.nrows)
        return dis                    
    
    def casecount(self,i,j):
        k=1
        start =0
        end =0
        while k<i:
            if self.table.cell(k,2).value:
                start = start +1
            k=k+1
        l=1
        while l<j:
            if self.table.cell(l,2).value:
                end = end +1
            l=l+1
        return end - start
    
    def datacase(self):
        test =[]
        a = self.suitedis()
        i =0
        while i<len(a)-1:
            bb=self.casecount(a[i], a[i+1])
            test.append(bb)                        
            i=i+1
        return test
    
    def realdatacase(self,data):
        i =0
        new =[0]
        while i<len(data):
            j =0 
            c = 0
            while j<=i:
                c = c+data[j]            
                j=j+1
            new.append(c)
            i = i+1
        return new
    
    def case(self):
        data = self.read()
        a =self.realdatacase(self.datacase())
        i=1
        allcases=[]        
        while i<len(a):
            cases = []
            case = data[1][a[i-1]:a[i]]
            summary = data[4][a[i-1]:a[i]]
            precondition = data[2][a[i-1]:a[i]]
            importance = data[3][a[i-1]:a[i]]
                          
            cases.append(case)
            cases.append(summary)
            cases.append(precondition)
            cases.append(importance)    
            allcases.append(cases)                            
            i=i+1
        return allcases
    
    def testsuite(self):
        i = 0
        data = self.read()
        suite =[]
        test = self.case()

        while i<len(data[0]):
            aa = {}
            aa.setdefault("testsuite",data[0][i].get("testsuite"))
            aa.setdefault("testcases",test[i][0])
            aa.setdefault("summary",test[i][1])
            aa.setdefault("precondition",test[i][2])
            aa.setdefault("importance",test[i][3])
            suite.append(aa)
            i=i+1
        return suite


        

class changetoxml:
    def __init__(self,excel,sheetname,output,filename):
        self.excel = excel
        self.sheetname = sheetname
        self.output = output
        self.filename = filename
        
    
    def run(self):
        self.importexcel = readexcel(self.excel,self.sheetname)
        
        self.testdata = self.importexcel.testsuite()
        self.xml = ExportXmlNoStep(self.testdata,self.output,self.filename)
        self.xml.export()
        

class exceloperate:   
    def __init__(self,testexcel):
        self.testexcel = testexcel   
        self.data = xlrd.open_workbook(self.testexcel)
        
    def getSheetNames(self):   
             
        return self.data.sheet_names()
 

        
            

        



if __name__ == "__main__":
    
    

    testexcel = raw_input("Please input the path of your excel file(like 'D:/testexcel.xls'):\n")
    sheetname = raw_input("Please input your sheetname of testcase(like 'Sheet1'):\n")
    output = raw_input("Please input your output folder (like 'D:/testcase') :\n")
    filename = raw_input("Please input your filename (like 'testcase'):\n")
    
    if os.path.exists(testexcel):     
        sheets = exceloperate(testexcel).getSheetNames()   
        
    else:
        print("The excel file is not existed.")    
    if sheetname in sheets:
        {}
    else:
        print("The sheet name is not existed.")   
    if os.path.exists(output):
        {}
    else:
        os.mkdir(output)
                
    aa =changetoxml(testexcel,sheetname,output,filename)
    aa.run()
    print("ok")

