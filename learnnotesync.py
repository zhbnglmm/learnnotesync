
from AccessConnector import AccessConnector

a = AccessConnector("SELECT Top 2 * FROM tblLearnNote")
for b in a.recordset()["fields"]:
    print(b.name)