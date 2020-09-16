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
  
    ![](https://github.com/nitinnbbhagat/SizingTool/blob/nitin/Images/Tree.PNG)
- Download & Install **[Python3](https://www.python.org/downloads/)**.
- Install the Python module pip:
  ```
  curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
  python get-pip.py
  ```

---

## STEP 1: Install
```
git clone https://github.com/nitinnbbhagat/SizingTool.git
cd SizingTool
pip install openpyxl
pip install statistics
python -m pip install pywin32
```
## STEP 2: Use the Data Discovery wizard to discover sensitive data

- To access the Data Discovery wizard, click the **Home** tab, and then click **Data Discovery**.

   ![](./Images/Img25.png " ")
- On the **Select Target** for Sensitive Data Discovery page, your target database is listed.
- Often, you want to perform data discovery against a production database where you can get an accurate and up-to-date picture of your sensitive data. You can discover sensitive data in the source database (a production or copy) and mask the cloned copies of this source database. Or, you can simply run a data discovery job on the actual database that you want to mask.
- Select your target database, and then click **Continue**.
- Next, the **Select Sensitive Data Model** page is displayed. On this page, you can create a new sensitive data model, select an existing one from the Library, or import a file-based sensitive data model.

   ![](./Images/Img26.png " ")
- Leave **Create** selected.
- Name the sensitive data model as **<username>SDM1**.
- Enable **Show and save sample data** so that Data Discovery retrieves sample data for each sensitive column, if it's available.
- Select your resource group.
- On the **Select Schemas for Sensitive Data Discovery** page, select the schema that you want Data Discovery to search. In this case, select the **HCM1** schema, and click **Continue**.

   ![](./Images/Img27.png " ")
- On the **Select Sensitive Types for Sensitive Data Discovery** page, you select the sensitive types that you want to discover. Data Discovery categorizes its sensitive types as Identification Information, Biographic Information, IT Information, Financial Information, Healthcare Information, Employment Information, and Academic Information. 

   ![](./Images/Img28.png " ")
Do the following:
1. **Expand all** the categories by moving the Expand All slider to the right to view each sensitive type. Notice that you can select individual sensitive types, sensitive categories, and all sensitive types.
2. Scroll down the list and review the sensitive types available.
3. Return to the top of the list and select the **Select All** check box.
4. At the bottom of the page, select the **Non-Dictionary Relationship Discovery** check box.<br>
Oracle Data Safe automatically discovers referential relationships defined in the data dictionary. The **Non-Dictionary Relationship Discovery** feature helps to identify application-level parent-child relationships that are not defined in the data dictionary. It helps discover application-level relationships so that data integrity is maintained during data masking.
5. When you are ready to start the data discovery job, click **Continue**.
6. Wait for the job to finish.

   ![](./Images/Img29.png " ")
- If the job is successful, the **Detail** column states Data discovery job finished successfully, and you can click **Continue**. Otherwise, you need to click **Back** or **Exit** and investigate the issue.

## STEP 3: Review and fine tune your sensitive data model

- On the **Non-Dictionary Referential Relationships** page, you are presented with a list of potential non-dictionary (application level) referential relationships that Data Discovery found by using column name patterns and column data patterns. 

*Do the following:*

- To view all of the columns, move the **Expand All** slider to the right. Data Discovery found some potentially sensitive columns (non-dictionary referential relationships) in the `HR` schema.

   ![](./Images/Img30.png " ")
   
- Click **Save** and **Continue**.<br>
The **Sensitive Data Discovery Result** page is displayed. On this page, you can view and manage the sensitive columns in your sensitive data model. Your sensitive data model is saved to the Library.

   ![](./Images/Img31.png " ")
   
- Notice that Data Discovery found sensitive columns in all three sensitive categories that you selected. To view the sensitive columns, move the **Expand All** slider to the right. The list includes the following:
  1. Sensitive columns discovered based on the sensitive types that you selected
  2. Dictionary-based referential relationships
  3. Non-dictionary referential relationships

*Note:*
- Take a look at how the sensitive columns are organized. Initially, they are grouped by sensitive categories and sensitive types.
- To list the sensitive columns by schema and table, select **Schema View** from the drop-down list next to the **Expand All Slider**. **Schema View** is useful for quickly finding a sensitive column in a table and for viewing the list of sensitive columns in a table. For example, in the `EMPLOYEES` table there are several sensitive columns listed.
- If needed, you can add and remove sensitive columns from your sensitive data model. Sensitive columns that have a check box are removable. To remove a sensitive column from your sensitive data model, you deselect its check box. You can use the **Add** button to add more sensitive columns.
- Notice that some of the sensitive columns do not have a check box. These are dependent columns. They have a relationship with their parent column. For example, in the `EMPLOYEES` table, `JOB_ID` is listed. It has a relationship defined in the Oracle data dictionary to the `JOBS.JOB_ID` sensitive column. If you remove a sensitive column that has a referential relationship, both the sensitive column and referential relationship are removed from the sensitive data model. Therefore, if you deselect `JOBS.JOB_ID`, then `EMPLOYEES.JOB_ID` is removed too.
- View the sample data for the `HCM1.SUPPLEMENTAL_DATA.LAST_INS_CLAIM` column.<br/>
The sensitive type is **Healthcare Provider** and the discovered sensitive column is `LAST_INS_CLAIM`, which has values such as `Cavity` and `Hair Loss`. Your value may be different. This column isn't a Healthcare Provider type of column and thus it is a false positive. You can deselect this column. Being able to remove a sensitive column is important when your sensitive data model includes false positives. To be able to recognize false positives, it helps to know your data well.

> **Tip:**<br>
To quickly locate a sensitive column, enter the name or part of the name in the search box.

## STEP 4: Modify your search for sensitive data and re-run the data discovery job

Suppose that you're missing some sensitive columns in your sensitive data model. While working in the Data Discovery wizard, you can backtrack to reconfigure and rerun the data discovery job. You can repeat the process as many times as you need until you feel that your sensitive data model is accurate. Try the following:

- Click **Back** <br>
Now you are on the **Select Sensitive Types for Sensitive Data Discovery** page. Here you can change your sensitive type selections and choose whether to include non-dictionary referential relationships in the search.
- Select all of the sensitive categories.

   ![](./Images/Img32.png " ")

- Deselect **Non-Dictionary Relationship Discovery**.

   ![](./Images/Img33.png " ")

- To rerun the data discovery job, click **Continue**.
- When the job is finished, click **Continue**.
Because you chose to not discover non-dictionary referential relationships, the wizard takes you directly to the **Sensitive Data Discovery Result** page.
- Expand all of the sensitive columns and review the results.

   ![](./Images/Img34.png " ")

- To view the newly discovered sensitive columns, click **View newly discovered sensitive columns only**. Notice that Data Discovery found additional sensitive columns.

## STEP 5: View the Sensitive Data Discovery report and analyze the report data

- Scroll down and click **Report** at the bottom right corner of the screen.<br>
The report shows you a chart that compares sensitive categories. You can also view totals of sensitive values, sensitive types, sensitive tables, and sensitive columns. The table at the bottom of the report displays individual sensitive column names, sample data for the sensitive columns, column counts based on sensitive categories, and estimated data counts.

   ![](./Images/Img35.png " ")

- Analyze the data:
  - To drill-down into a sensitive category in the chart, position your mouse over the chart, and then click the **Expand** button.

   ![](./Images/Img36.png " ")

   ![](./Images/Img37.png " ")

  - To drill-up, position your mouse over an expanded sensitive category, and then click the **Collapse** button.
  - To enlarge the chart, click the **Expand** button (double-arrows) in the bottom right corner. View the chart and click **Close**.
  - Expand the list of sensitive columns and review the information.

   ![](./Images/Img38.png " ")
  - Click **Exit**.

   ![](./Images/Img39.png " ")

- To access the report from the Reports tab, do the following:
  - Click the **Reports** tab.

   ![](./Images/Img40.png " ")

  - Scroll down, and under **Discovery Reports**, click **Data Discovery**.
  - Click your sensitive data model to open the report.
  
## STEP 6: View your sensitive data model in the Library

- Click the **Library** tab.
- Click **Sensitive Data Models**.<br>
The Sensitive Data Models page is displayed, listing the sensitive data models to which you have access.
For each sensitive data model, you can view information about when your sensitive data model was created, when it was last updated, and who owns it.

   ![](./Images/Img41.png " ")
   
- Click the name of your sensitive data model to open it.

   ![](./Images/Img42.png " ")

- To return to the **Sensitive Data Models** page, click the **Library** tab, and then click **Sensitive Data Models**.
- If you need to remove your sensitive data model from the Library, you can select the check box for it, and click **Delete**.

> **Note:**<br>
Keep your sensitive data model so that later you can use it for masking sensitive data in **[Masking Lab - Configure a Variety of Masking Formats with Oracle Data Safe](https://github.com/labmaterial/adbguides-dev/blob/master/adb-datasafe/Masking%20Lab%204.md)**.

## Step 7: Download your sensitive data model

- Select the check box for your sensitive data model.
- Click **Download**.
Your sensitive data model is downloaded to your browser.
- Open the file, and review the XML code.
- Save the file to your desktop as **<username> SDM1.xml**, and then close the file.<br>
In **[Masking Lab - Configure a Variety of Masking Formats with Oracle Data Safe](https://github.com/labmaterial/adbguides-dev/blob/master/adb-datasafe/Masking%20Lab%204.md)**, you import this file-based sensitive data model into the Data Masking wizard.

### All Done!
