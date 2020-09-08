import pypyodbc
import win32com.client

class AccessConnector(object):
    def __init__(self,sql):
        self.path=r'C:\Users\Alan\OneDrive\LearnNote\learnnote3.0.accdb'
        conn = pypyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + self.path + ";Uid=;Pwd=;")
        self.sql=sql
        self.cursor = conn.cursor()
        con = win32com.client.Dispatch(r'ADODB.Connection')
        DSN = 'PROVIDER=Microsoft.ACE.OLEDB.12.0;DATA SOURCE=' + self.path + ';'
        con.Open(DSN)
        self.rs = win32com.client.Dispatch(r'ADODB.Recordset')
        self.rs.Cursorlocation = 3
        self.rs.Open(self.sql, con)
    def recordset(self):
        return {"fields":self.rs.Fields,"rows":self.cursor.execute(self.sql)}

if __name__ == '__main__':
    a = AccessConnector("SELECT Top 2 * FROM tblLearnNote")
    for b in a.recordset()["fields"]:
        print(b.name)