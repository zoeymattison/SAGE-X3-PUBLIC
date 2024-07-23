## Sage X3 Backorder Report
### 1. Purpose
To generate an excel file containing a list of sales orders with backordered products, including several information columns, summarized colums and filterable columns.

### 2. Method
1. Generte and export the report to .xlsx with Powershell using _Invoke-SqlCmd_ and _Export-Excel_
2. Schedule the script to run via Windows tas scheduler, once a day at 7:20 AM
3. In Sage, schedule a recurring task to run once a day at 7:30 AM which will execute a custom integration to triggers the workflow
4. In the workflow, the attachment path is set to that of the generated .xlsx document
5. At the end of the task, the file is moved into an archive folder in the same directory

### 3. Setup Instructions
#### I. Initial Setup & Powershell
1. First, on your Sage X3 database server, create a new folder to house your powershell scripts, and a new folder to house your outputted .xlsx files. I've named them "Zoey's Scripts" and ZATTACH. I've placed ZATTACH in the LIVE folder of X3, and "Zoey's Scripts" in the Sage X3 installation.
2. In the supplied .ps1 file under Source, modify the SQL query to fit your needs, as well as the export path. The SQL query should remain surrounded by "@SQL QUERY HERE@". Here is an example of my output path:
    1. ```D:\Sage\X3\folders\LIVE\ZATTACH\BackOrderReport_$dateSuffix.xlsx```
    2. Note the _$dateSuffix_ variable which appends the current date. We will use this for archival purposes.
3. Always test that your powershell script executes successfully using an IDE or by executing it manually with powershell.
#### II. Creating a task in the task scheduler
1. Open Window Task Scheduler (taskschd.msc).
2. Click "Action" in the toolbar and then "Create Task".
    ##### General Tab
    1. Enter a name and description.
    2. Select "Run whether user is logged on or not.
    3. Selcet "Run with highest priveleges".
    ##### Triggers Tab
    1. Select "New".
        1. For our purposes, we need the report to genreate every morning at 0720, from Monday through Friday, so this tutorial will be based on that setup.
    2. Select "Weekly".
    3. Set the start date and time - today's date and 7:20:00 AM.
    4. Set the recurrance to every Monday through Friday.
    5. Disabe all advanced settings, and ensure "Enabled" is checked.
    ##### Actions Tab
    1. Select "New".
    2. Set the action type to "Start a program".
    3. In the Program/script path field, enter the path to powershell.exe.
        1. ```C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe```
    4. In the Add arguments (optional) field, enter the path to your .ps1 file, as well as a few arguments for the execution.
        1. ```-noprofile -executionpolicy unrestricted -noninteractive -file "D:\Sage\Zoey's Scripts\BackOrderReport.ps1"```
    5. Click OK.
    ##### Conditions, Settings and History Tabs
    1. These tabs are optional, and you can leave them as default. Under settings, you can modify the timeout of the task if your SQL query is prone to long execution times. Due to the size of our data, our execution time is about 2 minutes, so I've set the "Stop the task if it runs longer than" time to 1 hour.

3. Click "OK" to generate the action. Enter the password for the user. You can now test that the task runs successfully by right-clicking on the new task in the list and selecting "Run". Note that manually running the task will cause the Status to remain as "Running", but it's not actually doing so. Press F5 to refresh the list after seeing a successful output of your excel file.
#### Creating the workflow in Sage X3
1. In Sage X3, launch the GESAWA function (Workflow) from the Setup menu (Setup -> Workflows -> Workflow Rules)
2. Enter creation mode to create a new workflow, and give it a name and description.
    ##### General Tab
    1. Event Type: Miscellaneous.
    2. Event Code: (create a new one by selection the actin button and then "Miscellaneous Event Types") - I've called mine ZBO.
    3. Ensure "Trigger Mail" in the management block is checked.
    ##### Recipient Tab
    1. Enter your user recipients or emails, marking them as "Yes" or "Copy" under the Send Mail columns (note: "Send Mail" CC's the email).
    ##### Messages Tab
    1. Enter your subject and text for the message. Example:
        1. ![Workflow Message Setup](https://ptpimg.me/fkrd8v.png)
    2. Under the Management block, select "Any" s the sending type (unless you need to specifically select Client or Server depending on your situtation).
    3. Under the ttachments block, we're going to link directly to the exported file, while building a dynamic constructor for the previous _$dateSuffix_ variable. We need to now build it with 4GL syntax. The _$dateSuffix_ variable in our powershell script generates the date as YYYYMMDD. In 4GL, we simply need to use for _format$()_ function on the current date function, which looks like the following:
        1. ```format$("YYYYMMDD",date$)```
    4. Now, to get it into the Attached document path, we need it to generate as a string, by passing the entire line into the _num$()_ function:
        1. ```num$(format$("YYYYMMDD",date$))```
    5. We can now concatenate this onto a hard-coded document path. Since the document will always be called "BackOrderReport", it looks like this:
        1. ```"D:\Sage\X3\folders\LIVE\ZATTACH\BackOrderReport_"+num$(format$("YYYYMMDD",date$))+".xlsx"```
        2. This simple line will allow Sage to always grab the correct document generated today, even if you leave the old documents in the path folder.
    5. The workflow is complete. The Milestone and Action tabs are not needed unless you need to use them for your own purposes. Create and validate the workflow.