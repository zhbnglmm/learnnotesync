import pypyodbc
import win32com.client

class AccessConnector(object):
    def __init__(self,sql):
        self.path=r'C:\Users\Alan\OneDrive\LearnNote\learnnote3.0.accdb'
        conn = pypyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + self.path + ";Uid=;Pwd=;")
        self.sql=sql
        self.cursor = conn.cursor()

    def recordset(self):
        con = win32com.client.Dispatch(r'ADODB.Connection')
        DSN = 'PROVIDER=Microsoft.ACE.OLEDB.12.0;DATA SOURCE=' + self.path + ';'
        con.Open(DSN)
        self.rs = win32com.client.Dispatch(r'ADODB.Recordset')
        self.rs.Cursorlocation = 3
        self.rs.Open(self.sql, con)
        self.list=[]
        for row in self.cursor.execute(self.sql):
            dict ={}
            for x in range(0,self.rs.Fields.Count):
                dict[self.rs.Fields(x).Name.lower()]=row[x]
            self.list.append(dict)
        return self.list
    def update(self):
        self.cursor.execute(self.sql)
        self.cursor.commit()


        
if __name__ == '__main__':
    a = AccessConnector("SELECT Top 2 * FROM tblLearnNote")
    print(a.recordset())