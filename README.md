# Sizing Tool Automation

## Introduction
**Budgetary Architecture Recommendation (BAR)** provides Capacity Planning solutions for customer's IT landscape. The centralized self-service tool helps qualify, scope and budget customer opportunities. BAR addresses Greenfield and existing/consolidation deployments of Applications, Technology and Systems for On-Premise and Cloud implementations.

### Challenge
Forecast of Sizing for DB Workloads might take several hours to fill up information into BAR Tool. It may take an average of 8-10 hours and more than one technical resource to input paramters manually against Opportunity on the BAR Tool for Medium-sized workload.

### Current Procedure
- AWR Minor Files uploaded onto BAR Tool using Console against Opportunity created in bulk/indivudual file upload mode.
- Summarized CSV is generated for uploaded files which contains required parameters such as:
   - Memory (GB)
   - Memory Peak Utilization (%)
   - CPU Peak Utilization (%)
   - Database Name
   - Server Name
   - RAC/Non-RAC Database
   - Database Management System Vendor
   - Database Version
   - Database Sizer (GB)
   - Read & Write IOPS
   - SGA & PGA
- Using the summarized CSV file, create a Worklad and manually input above mentioned paramters in Console.
- Create Deployment Scenario and run for future state architecture prediction.

### Prerequisites
- Create a Stage Directory and place all AWR Minor Files in the same. The Stage Directory can have sub-folders within itself as well.
  
   ![Tree Structure](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Tree.PNG)
- Download & Install **[Python3](https://www.python.org/downloads/)**.
- Install the Python module pip:
   ```
   $ curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
   $ python get-pip.py
   ```

---

## Code Flow
- The code will go through each AWR Minor File present in the Stage Directory.
- Each host present in each of the files will be mapped to a Server Model, which will be inputted by the User.
- For each Database mentioned in the AWR Minor Files, the above mentioned values will be fetched, upon which calculations will be conducted.
- The calculations will be uploaded to the workbook and the same will be uploaded to the BAR Tool.
- The master workbook is present in the **MainWorkbook** directory. The name of the Workbook is **OfflineCustomerWorkbook-v200330081105.xlsm**.
- On execution, the code will create a copy of the workbook. All changes will be made to the newly created workbook and the same will be uploaded to the BAR Tool. The newly created workbook will be present in a newly created directory. The naming structure of the directory is **newWorkbookFile_*date(time)*** and that of the workbook is **OfflineCustomerWorkbook-v200330081105*_date(time)*.xlsm**. **OfflineCustomerWorkbook-v200330081105*_date(time)*.xlsm**
---


## STEP 1: Install Dependencies
```
$ git clone https://github.com/nitinnbbhagat/SizingTool.git
$ cd SizingTool
$ pip install openpyxl
$ pip install statistics
$ python -m pip install pywin32
```

## STEP 2: Execute
```
$ python WrapperScript.py
```

- Enter the Initial Values:
   - Stage Directory
   - Environment Name
   - Planned Data Growth Value
   - Planned CPU Growth Value

   ![Initial Values](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/InitialValues.PNG)
- Search for the Server Model Name according to a keyword. The values are fetched from the ```servers_m-values.csv``` file present in the working code directory.
- Copy the correct Server Model Name from the list shown in the terminal and proceed with the execution.
   ![Server Model Name](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/ServerModelName.png)
- Once all the values are fetched from each AWR Minor Files, Validate the sheet.
   ![Validate1](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Validate1.PNG)
   ![Validate2](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Validate2.PNG)
   ![Validate3](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Validate3.PNG)
- The Validation takes place thrice in the complete execution of the code:
   - ReadDatabaseInputs.py
   - ReadServerInputs.py
   - Upload Workbook process


## STEP 3: Upload
Provide the given fields:
- Contact email address or drop box name
- Sizing engagement
- Deployment platform (dropdown)

![Upload1](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Upload1.PNG)

Click on **Upload** once the values are filled.
The following message will be displayed once the Upload is complete.
![Upload2](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Upload2.PNG)

You will receive the following email upon the same.
![Upload3](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Upload3.PNG)

## STEP 4: Proceed with Sizing
Head over to the **[BAR Tool](http://sizingtool.us.oracle.com)** to complete the Sizing activity.
<br>
Alternatively, you can click on the links present in the email which you received.

### All Done!
