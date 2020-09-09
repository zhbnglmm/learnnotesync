#-------------------------------------------------------------------------------
# Name:        DataBaseHandler
# Purpose:     To provide a customized solution for data storage
#
# Author:      ALAN.ZHANG
#
# Created:     23/12/2015
# Copyright:   (c) ALAN.ZHANG 2015
# Remark:
#-------------------------------------------------------------------------------
import pyodbc
import os.path
import datetime
# import io
# import sys
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8')
class accesshandler(object):
    def __init__(self,accessfullname,tablename):
        if not os.path.isfile(accessfullname):
            raise Exception("DataBaseHandler Error:The Access Files does not exist:"+accessfullname)
        else:
            self.__dir=accessfullname
        if os.path.splitext(self.__dir)[1]!=".accdb":
            raise Exception("DataBaseHandler Error:A .accdb file was expected.")
        cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
        cursor = cnxn.cursor()
        for tbl in cursor.tables():
            if tbl.table_name==tablename:
                self.__tablename=tablename
                cursor.close()
                cnxn.close()
                #print self.__tablename
                break
        else:
            cursor.close()
            cnxn.close()
            raise Exception("DataBaseHandler Error:Table does not exsit:"+tablename)
    def printfieldnamelist(self):
        cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
        cursor = cnxn.cursor()
        cursor.execute("SELECT * FROM "+self.__tablename)
        columns = [column[0] for column in cursor.description]
        cursor.close()
        cnxn.close()
        print (columns)
        for i, column in enumerate(columns):
            print (i,column)
    def __getcolumntype(self,columnname):
        cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
        cursor = cnxn.cursor()
        cursor.execute("SELECT * FROM "+self.__tablename)
        columns = [column[0] for column in cursor.description]
        #print("cursor.description",cursor.description)
        columnstypes = [column[1] for column in cursor.description]
        cursor.close()
        cnxn.close()
        ismatch=False
        idex=0
        for idx,col in enumerate(columns):
            if col==columnname:
                ismatch=True
                idex=idx
        if ismatch==True:
            return columnstypes[idex]
        else:
            return ""
    def createsql(self,colDic):

        #No matter what data type of database field is, the value of each colDic shall be a string.
        # this function will get field data type from database and automatically generate a corresponding sql. ie, with or without "'"
        # However, exception will be raised when insert "a" to an int field, since that field can only receive "1","2"......
        # some keyword can not be used as column name, e.g. level
        cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
        cursor = cnxn.cursor()
        for (fieldname, data) in colDic.items():
            if (not fieldname) or (not type(fieldname) is str):
                cursor.close()
                cnxn.close()
                raise Exception("DataBaseHandler Error:can not receive a non-string/null fieldname")
            # if (not type(data) is str) or (not type(data) == "bs4.element.NavigableString"):
            #     cursor.close()
            #     cnxn.close()
            #     print(" type(data)", type(data))
            #     raise Exception("DataBaseHandler Error:can not receive a non-string data")
        sqlFrontHalf = """INSERT INTO """+self.__tablename+" ("
        sqlRearHalf = ") VALUES("
        for (fieldname, data) in colDic.items():
            if data is not None:
                sqlFrontHalf = sqlFrontHalf + fieldname+","
                #print("self.__getcolumntype(fieldname)",self.__getcolumntype(fieldname))
                typeName = self.__getcolumntype(fieldname).__name__ if hasattr(self.__getcolumntype(fieldname),__name__) else ""
                if not (typeName.find("datetime") or typeName.find("unicode")):
                    sqlRearHalf = sqlRearHalf + data + ","
                else:
                    sqlRearHalf = sqlRearHalf + " '" + data + "',"
        sqlFrontHalf = sqlFrontHalf[:-1]
        sqlRearHalf = sqlRearHalf[:-1]
        sql = sqlFrontHalf + sqlRearHalf
        sql = sql + ');'
        cursor.close()
        cnxn.close()
        return sql
    def insertdata(self,sql):
         cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
         cursor = cnxn.cursor()
         cursor.execute(sql)
         cursor.commit()
         cursor.close()
         cnxn.close()

    def closedatabase(self):
         cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
         cursor = cnxn.cursor()
         cursor.close()
         cnxn.close()

    def selectdata(self, fieldList=None):
        cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
        cursor = cnxn.cursor()
        if not fieldList:
            data = cursor.execute("select * from " + self.__tablename +" order by ID")
        else:
            fieldStr = ""
            for f in fieldList:
                fieldStr = f + ","
            fieldStr =  fieldStr[:-1]
            data = cursor.execute("select "+ fieldStr +" from " + self.__tablename)  
        # cursor.commit()
        # cursor.close()
        # cnxn.close()
        return data
    def executesql(self,sql):
        cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ self.__dir)
        cursor = cnxn.cursor()
        cursor.execute(sql)
        cursor.commit()
        cursor.close()
        cnxn.close()  


def test3():
    ts=accesshandler("C:\\Alan\\baokao\\baokao.accdb","tblEolSchoolList")
    ts.printfieldnamelist()
    test = ts.createsql({"schoolid":"123456","schoolname":"haha"})
    a="2015-1-1"
    b="2"
    c="3"
    d="/zzz2015-1-1"
    e="5"
    #ts.closedatabase()
    print("test",test)
    # ts.insertdata(test)
    kk = ts.selectdata(["schoolid"])
    idx = 0
    for k in kk:
        idx = idx +1
        # print("chardet.detect(k)",chardet.detect(k))
        print(k[0])
    print(idx)
    ts.closedatabase()

def main():
    test3()
    pass

if __name__ == '__main__':
    main()
