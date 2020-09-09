import requests
import json
import datetime
import logging
from AccessConnector import AccessConnector
from DataBaseHandler import accesshandler #这个是几年前写的，我都给忘了，搞得现在又做了个AccessConnector，不过insert太难写,就混着用吧。以后再重构。
class DateEncoder(json.JSONEncoder): #重写json序列化类以解决datetime类型的数据不能被json的错误
    def default(self, obj):
        if isinstance(obj, datetime.datetime):
            return obj.strftime('%Y-%m-%d %H:%M:%S')

        elif isinstance(obj, datetime.date):
            return obj.strftime("%Y-%m-%d")
        else:
            return json.JSONEncoder.default(self, obj)

def Sync2Server(rootURL):
    conn = AccessConnector("SELECT * FROM tblFileNumber where Sync=False")
    FileNumberRecords = conn.recordset()
    conn = AccessConnector("SELECT * FROM tblLearnNote where Sync=False")
    LearnNoteRecords = conn.recordset()
    PictureRecords=[]
    for LearnNoteRecord in LearnNoteRecords:
        conn = AccessConnector("SELECT * FROM tblPicture Where ID_Problem="+str(LearnNoteRecord['id_problem']))
        p = conn.recordset()
        PictureRecords.extend(p)

    data =json.dumps({
        "FileNumberRecords": FileNumberRecords, 
        "LearnNoteRecords": LearnNoteRecords,
        'PictureRecords':PictureRecords
        },cls=DateEncoder)
    response = requests.post(rootURL+"/sync2server/", data=data,headers={'Content-Type': 'application/json'})
    errorcode =json.loads(response.content)["errorcode"]
    if errorcode == 0:
        conn = AccessConnector('update tblLearnNote set sync=1')
        conn.update()
        conn = AccessConnector('update tblFileNumber set sync=1')
        conn.update()
        logging.debug("sync2server success.")
    elif errorcode == 1:
        conn = AccessConnector("update tblConfig set ItemStatus=True where tblConfig.ItemName='IsSyncConflict'")
        conn.update()
        logging.warning(response.content.decode('unicode_escape')) 

def Sync2Access(rootURL):
    response = requests.post(rootURL+"/sync2access/")
    errorcode =json.loads(response.content)["errorcode"]
    if errorcode == 1:
        logging.debug("无待同步数据！")
    elif errorcode == 0:
        learnnotelist= json.loads(response.content)["learnnotelist"]
        filenumberlist= json.loads(response.content)["filenumberlist"]
        logging.debug("成功得到服务器数据！")
        # print(learnnotelist)
        # print(filenumberlist)
        ac = accesshandler(r"C:\Users\Alan\OneDrive\LearnNote\learnnote3.0.accdb", "tblFileNumber")
        ac2 = accesshandler(r"C:\Users\Alan\OneDrive\LearnNote\learnnote3.0.accdb", "tblLearnNote")
        ac3 = accesshandler(r"C:\Users\Alan\OneDrive\LearnNote\learnnote3.0.accdb", "tblPicture")
        for filenumber in filenumberlist:
            ac.executesql("delete * from tblFileNumber Where FileNumber='"+filenumber["filenumber"]+"'")
            sql = ac.createsql({
                "ID_Problem": str(filenumber['id_problem']),
                "FileNumber": filenumber['filenumber'],
                "FileTitle": filenumber['filetitle'],
                "RecordTime": filenumber['recordtime'],
                "Active": str(filenumber['active']),
                "sync": '1',
                })
            ac.executesql(sql)
        for learnnote in learnnotelist:
            ac.executesql("delete * from tblPicture Where ID_Problem="+str(learnnote["id_problem"]))
            ac.executesql("delete * from tblLearnNote Where ID_Problem="+str(learnnote["id_problem"]))
            sql = ac2.createsql({
                "ID_Problem": str(learnnote['id_problem']),
                "D_Level": learnnote['d_level'],
                "Source": learnnote['source'],
                "Notes": learnnote['notes'],
                "Answer": learnnote['answer'],
                "UsedTime": learnnote['usedtime'],
                "UsedTimes": str(learnnote['usedtimes']),
                "InputDate": learnnote['inputdate'],
                "Active": str(learnnote['active']),
                "sync": '1',
                })
            ac.executesql(sql)
            for picture in learnnote['pictures']:
                sql = ac3.createsql({
                    "ID_Problem": str(picture['id_problem']),
                    "FilePath": picture['filepath'],
                    "isAnswer": str(picture['isanswer']),
                    })
                ac.executesql(sql)
        logging.debug("已把服务器数据同步到本机！")
    

if __name__ == '__main__':
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    file = open('synclog.log', encoding="utf-8", mode="w")
    logging.basicConfig(stream=file, level=logging.DEBUG, format=LOG_FORMAT)

    rootURL = "http://localhost:5000"
    conn = AccessConnector("SELECT * FROM tblFileNumber Where Sync = False")
    FileNumberRecords = conn.recordset()
    conn = AccessConnector("SELECT * FROM tblLearnNote Where Sync = False")
    LearnNoteRecords = conn.recordset()
    if FileNumberRecords or LearnNoteRecords:
        Sync2Server(rootURL)
    else:
        Sync2Access(rootURL)
