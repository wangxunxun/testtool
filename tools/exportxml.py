#coding=utf-8
from xml.dom import minidom
import xlrd
import os
from idlelib.IOBinding import encoding

    
class exportxml:
    def __init__(self,testdata,outputfolder,filename):
        self.testdata = testdata
        self.output = outputfolder
        self.filename = filename
        
    def export(self):
        impl = minidom.getDOMImplementation()
        dom = impl.createDocument(None, None, None)
        testsuite =  dom.createElement("testsuite")
        inittestsuite = dom.appendChild(testsuite)
        i=0
        testsuite_s =[]
        while i<len(self.testdata):                        
            testsuite =  dom.createElement("testsuite")        
            inittestsuite.setAttribute("name", "")   
            testsuite_s.append(testsuite)
            testsuite_s[i] =  dom.createElement("testsuite")       
            testsuite_s[i].setAttribute("name", str(self.testdata[i].get("testsuite")))   
            testcase_s=[]
            summary_s =[]
            preconditions_s =[]
            execution_type_s =[]
            importance_s =[]
            steps_s=[]            
            j=0
            while j<len(self.testdata[i].get("testcases")):        
                testcase = dom.createElement("testcase")
                testcase_s.append(testcase)
                summary = dom.createElement("summary")
                summary_s.append(summary)
                preconditions = dom.createElement("preconditions")
                preconditions_s.append(preconditions)
                execution_type = dom.createElement("execution_type")
                execution_type_s.append(execution_type)
                importance = dom.createElement("importance")
                importance_s.append(importance)
                steps = dom.createElement("steps")
                steps_s.append(steps)                                        
                testcase_s[j].setAttribute("name", str(self.testdata[i].get("testcases")[j]))                                                                                    
                summary_text = dom.createTextNode(str(self.testdata[i].get("summary")[j]))
                precondition_text = dom.createTextNode(str(self.testdata[i].get("precondition")[j]))
                execution_type_text = dom.createTextNode(str(self.testdata[i].get("execution_type")[j]))
                importance_text = dom.createTextNode(str(self.testdata[i].get("importance")[j]))            
                summary_s[j].appendChild(summary_text)                
                preconditions_s[j].appendChild(precondition_text)                
                execution_type_s[j].appendChild(execution_type_text)                
                importance_s[j].appendChild(importance_text)
                step_s=[]
                step_number_s =[]
                actions_s =[]
                expectedresults_s =[]
                step_execution_type_s =[]
                testcase_s[j].appendChild(summary_s[j])
                testcase_s[j].appendChild(preconditions_s[j])
                testcase_s[j].appendChild(execution_type_s[j])
                testcase_s[j].appendChild(importance_s[j])            
                k=0
                while k<len(self.testdata[i].get("steps")[j]):
                    step = dom.createElement("step")
                    step_number = dom.createElement("step_number")
                    actions = dom.createElement("actions")
                    expectedresults = dom.createElement("expectedresults")
                    step_execution_type = dom.createElement("execution_type")                            
                    step_s.append(step)
                    step_number_s.append(step_number)
                    actions_s.append(actions)
                    expectedresults_s.append(expectedresults)
                    step_execution_type_s.append(step_execution_type)       
                    step_number_text = dom.createTextNode(str(self.testdata[i].get("steps")[j][k].get("stepid")))
                    actions_text = dom.createTextNode(str(self.testdata[i].get("steps")[j][k].get("action")))
                    expectedresults_text = dom.createTextNode(str(self.testdata[i].get("steps")[j][k].get("result")))
                    step_execution_type_text = dom.createTextNode(str(self.testdata[i].get("steps")[j][k].get("execution_type")))                    
                    step_number_s[k].appendChild(step_number_text)
                    actions_s[k].appendChild(actions_text)
                    expectedresults_s[k].appendChild(expectedresults_text)
                    step_execution_type_s[k].appendChild(step_execution_type_text)                
                    step_s[k].appendChild( step_number_s[k])
                    step_s[k].appendChild(actions_s[k])
                    step_s[k].appendChild(expectedresults_s[k])
                    step_s[k].appendChild(step_execution_type_s[k])                
                    steps_s[j].appendChild(step_s[k])
                    testcase_s[j].appendChild(steps_s[j])                                                          
                    testsuite_s[i].appendChild(testcase_s[j])                            
                    k=k+1        
                j=j+1        
            inittestsuite.appendChild(testsuite_s[i])                
            i=i+1    
        dom.appendChild(inittestsuite)
        f=open(self.output+"/"+self.filename+".xml",'w',encoding = "utf-8")
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
        self.summary =[]
        self.preconditon =[]
        self.casetype = []
        self.importance =[]
        self.steps = []
        self.casestep = []
        i =1
        j =self.table.nrows
        while i<j:
            step ={}
            testsuite = {}
            if self.table.cell(i,0).value and self.table.cell(i,1).value and self.table.cell(i,7).value:
                testsuite.setdefault("testsuite",self.table.cell(i,0).value)
                self.testcases.append(self.table.cell(i,1).value)
                self.summary.append(self.table.cell(i,2).value)
                self.preconditon.append(self.table.cell(i,3).value)
                self.casetype.append(self.table.cell(i,4).value)
                self.importance.append(self.table.cell(i,5).value)
                step.setdefault("stepid",self.table.cell(i,6).value)
                step.setdefault("action",self.table.cell(i,7).value)
                step.setdefault("result",self.table.cell(i,8).value)
                step.setdefault("execution_type",self.table.cell(i,9).value)
                self.steps.append(step)
                self.testsuites.append(testsuite)
                
                
            elif not self.table.cell(i,0).value and not self.table.cell(i,1).value and self.table.cell(i,7).value:

                step.setdefault("stepid",self.table.cell(i,6).value)
                step.setdefault("action",self.table.cell(i,7).value)
                step.setdefault("result",self.table.cell(i,8).value)
                step.setdefault("execution_type",self.table.cell(i,9).value)
                self.steps.append(step)
                
            elif not self.table.cell(i,0).value and self.table.cell(i,1).value and self.table.cell(i,7).value:

                
                self.testcases.append(self.table.cell(i,1).value)
                self.summary.append(self.table.cell(i,2).value)
                self.preconditon.append(self.table.cell(i,3).value)
                self.casetype.append(self.table.cell(i,4).value)
                self.importance.append(self.table.cell(i,5).value)
                step.setdefault("stepid",self.table.cell(i,6).value)
                step.setdefault("action",self.table.cell(i,7).value)
                step.setdefault("result",self.table.cell(i,8).value)
                step.setdefault("execution_type",self.table.cell(i,9).value)
                self.steps.append(step)
            i=i+1
        self.testdata.append(self.testsuites)
        self.testdata.append(self.testcases)
        self.testdata.append(self.summary)
        self.testdata.append(self.preconditon)
        self.testdata.append(self.casetype)
        self.testdata.append(self.importance)
        self.testdata.append(self.steps)

        return self.testdata
            
    def suitedis(self):        
        j=1
        dis =[]        
        while j<self.table.nrows:
            if self.table.cell(j,0).value:
                dis.append(j)
            j=j+1
        dis.append(self.table.nrows)
        return dis
    def casedis(self):
        j=1
        dis =[]
        while j<self.table.nrows:
            if self.table.cell(j,1).value:
                dis.append(j)
            j=j+1
        dis.append(self.table.nrows)
        return dis
    
    def nocaserow(self):
        a = self.casedis()
        i= 0
        count = []
        while i<len(a):
            j = 0
            k =0
            while j<a[i]:
                if self.table.cell(j,7).value =="":
                    k=k+1                    
                j=j+1
            count.append(k)
            i=i+1
        return count
    
    def realcasedis(self):
        casedis = self.casedis()
        nocaserow = self.nocaserow()
        realcasedis = []
        i = 0
        while i<len(casedis):
            realrow = casedis[i]-nocaserow[i]
            realcasedis.append(realrow)
            i=i+1
        return realcasedis
                    
    
    def datastep(self):
        data = self.read()
        a = self.realcasedis()
        test =[]
        i =0
        while i<len(a)-1:
            test.append(data[6][a[i]-1:a[i+1]-1])
            i=i+1
        return test
    
    def casecount(self,i,j):
        k=1
        start =0
        end =0
        while k<i:
            if self.table.cell(k,1).value:
                start = start +1
            k=k+1
        l=1
        while l<j:
            if self.table.cell(l,1).value:
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
            summary = data[2][a[i-1]:a[i]]
            precondition = data[3][a[i-1]:a[i]]
            execution_type = data[4][a[i-1]:a[i]]
            importance = data[5][a[i-1]:a[i]]
            steps = self.datastep()[a[i-1]:a[i]]                                
            cases.append(case)
            cases.append(summary)
            cases.append(precondition)
            cases.append(execution_type)
            cases.append(importance)
            cases.append(steps)            
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
            aa.setdefault("execution_type",test[i][3])
            aa.setdefault("importance",test[i][4])
            aa.setdefault("steps",test[i][5])
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
        self.xml = exportxml(self.testdata,self.output,self.filename)
        self.xml.export()
        

class exceloperate:   
    def __init__(self,testexcel):
        self.testexcel = testexcel   
        self.data = xlrd.open_workbook(self.testexcel)
        
    def getSheetNames(self):   
             
        return self.data.sheet_names()
 

        
            

        



if __name__ == "__main__":
    
    d = changetoxml("C:/Users/xun/workspace/testtool/Model/Test Case.xls","Function Test","D:/","222")
    d.run()
'''    
    testexcel =input("Please input the path of your excel file(like 'D:/testexcel.xls'):\n")
    sheetname = input("Please input your sheetname of testcase(like 'Sheet1'):\n")
    output = input("Please input your output folder (like 'D:/testcase') :\n")
    filename = input("Please input your filename (like 'testcase'):\n")
    
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
'''