#coding=utf-8
'''
Created on 2015年7月23日

@author: xun
'''
import json



class CommonTool:
    def changeToJson(self,data):
        data = json.dumps(data,indent=4,ensure_ascii=False)
        return data.replace("\\", "")
        
    def zp(self,r):
        try:
            print(json.dumps(json.loads(r.text), indent=4,ensure_ascii=False, encoding="utf-8"))
            return json.dumps(json.loads(r.text),ensure_ascii=False)
        except Exception, e:
            print(e)
            print(r.text)
            return(r.text)
