import os
import datetime
import shutil



# ---------------------------------------------------------------------------------
# Function to find workbook file
# ---------------------------------------------------------------------------------
def findWorkbookfile(name, location):
    for root, dirs, files in os.walk(location):
        if name in files:
            return os.path.join(root, name)

location = os.getcwd()
name = "OfflineCustomerWorkbook-v200330081105.xlsm"
workbookFile = findWorkbookfile(name, location)
timestampStr = datetime.datetime.now().strftime("%d-%b-%Y(%H-%M-%S.%f)")
os.mkdir('newWorkbookFile_%s' % timestampStr)
newworkbookFileLocation = location + '\\' + 'newWorkbookFile_%s' % timestampStr
newworkbookFile = newworkbookFileLocation + '\\' + "%s"  % name
shutil.copy(workbookFile, newworkbookFileLocation)
renamednewworkbookFile = os.path.splitext(newworkbookFile)[0] + "_%s" % (timestampStr + os.path.splitext(newworkbookFile)[1])
os.rename(newworkbookFile, renamednewworkbookFile)
workbookFile = renamednewworkbookFile

print(workbookFile)