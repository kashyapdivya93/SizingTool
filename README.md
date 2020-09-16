# Sizing Tool Automation

## Introduction
Budgetary Architecture Recommendation (BAR) provides Capacity Planning solutions for customer's IT landscape. The centralized self-service tool helps qualify, scope and budget customer opportunities. BAR addresses Greenfield and existing/consolidation deployments of Applications, Technology and Systems for On-Premise and Cloud implementations.

### Challenge
Forecase of Sizing for DB Workloads might take several hours to fill up information into Bartool. It may take an average of 8-10 hours and more than one technical resource to input paramters manually against Opportunity on the BAR Tool for Medium-sized workload.

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
- Search for the Server Model Name according to the keyword. The values are fetched from the ```servers_m-values.csv``` file present in the working code directory.
- Copy the correct Server Model Name from the list shown in the terminal and proceed with the execution.
   ![Server Model Name](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/ServerModelName.PNG)

### All Done!
