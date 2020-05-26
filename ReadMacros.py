import win32com.client
import os

# pre requisites are to have the macros enabled
# crack version of the template

# TODO
# change the hardcoded the value of path of the excel


def validateMappingSheet():
    try:
        print("Validating inputs.... Please click on the excel dialog box")
        xlApp = win32com.client.DispatchEx('Excel.Application')
        xlsPath = os.path.expanduser(
            'D:\Work\Automation\Sizing\SizingTool\OfflineCustomerWorkbook-v191219102705.xlsm')
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlApp.Run('Validate')
        wb.Save()
        xlApp.Quit()
        print("Macro ran successfully!")
    except Exception as e:
        print(e)
        print("Error found while running the excel macro!")
        xlApp.Quit()


def validateDBServerSheet():
    try:
        xlApp = win32com.client.DispatchEx('Excel.Application')
        xlsPath = os.path.expanduser(
            'D:\Work\Automation\Sizing\SizingTool\OfflineCustomerWorkbook-v191219102705.xlsm')
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlApp.Run('ValidateDBServerSheet')
        wb.Save()
        xlApp.Quit()
        print("Macro ran successfully!")
    except Exception as e:
        print(e)
        print("Error found while running the excel macro!")
        xlApp.Quit()


def validateDatabaseSheet():
    try:
        xlApp = win32com.client.DispatchEx('Excel.Application')
        xlsPath = os.path.expanduser(
            'D:\Work\Automation\Sizing\SizingTool\OfflineCustomerWorkbook-v191219102705.xlsm')
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlApp.Run('ValidateDatabaseSheet')
        wb.Save()
        xlApp.Quit()
        print("Macro ran successfully!")
    except Exception as e:
        print(e)
        print("Error found while running the excel macro!")
        xlApp.Quit()
