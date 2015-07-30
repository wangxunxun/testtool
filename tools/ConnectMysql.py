'''
Created on 2015年7月30日

@author: xun
'''
import pymysql


class mysqlconnect:
    def __init__(self,host,user,passwd,db,port,charset):
        self.host = host
        self.user = user
        self.passwd = passwd
        self.db = db
        self.charset = charset
        self.port = port
        
    def connect(self):
        try:
            self.conn=pymysql.connect(host = self.host,user = self.user,passwd = self.passwd,db = self.db,port = self.port,charset = self.charset)
            self.cur = self.conn.cursor()
            return self.conn
        except Exception as e:
            return str(e)   
    
    
    def execute(self,script):
        try:
            array = []
            self.cur.execute(script)            
            result = self.cur.fetchall()
            count = len(result)
            array.append(count)
            array.append(str(result))
            if script[:6] == "insert" or script[:6] == "update" or script[:6] == "delete":
                try:
                    self.conn.commit()
                except:
                    self.conn.rollback()
            return array

        except Exception as e:
            return str(e)

    def close(self):                  
        self.cur.close()
        self.conn.commit()  
        self.conn.close()
        
if __name__ == "__main__":  
#    conn= pymysql.connect(host='69.164.202.55',user='test',passwd='test',db='test',port = 3306,charset='utf8')
#    conn = mysqlconnect('69.164.202.55','test','test','test',3306,"utf8").connect()
#   cur = conn.cursor()
#    print(cur.execute("show tables"))
#    print(cur.fetchone())
#    for r in cur:
#        print(r)
    
   
    a =mysqlconnect('69.164.202.55','test','test','test',3306,"utf8")
    a.execute("show tables")
    a.execute(" insert into mytable values ('abccs','f','1977-07-07','china')")
    a.execute("update mytable set birth = '1988-10-10' where name = 'abccs'")
    a.execute("delete from mytable where name = 'abccs'")
    a.execute("select * from mytable")
#    a.execute("show tables")
#    a.execute(" Create TABLE mytable (name VARCHAR(20), sex CHAR(1),birth DATE, birthaddr VARCHAR(20))")
#    a.execute('show tables')
#    a.execute('describe mytable')
#    a.execute("insert into message value (1,'1wanggang','wanggang2','3434',1)")
#    a.execute("select * from roles")
#    a.execute("describe roles")
    a.close()

