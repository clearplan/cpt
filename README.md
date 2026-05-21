![header](/images/header.png?raw=true)

![cpt](/images/cpt.png?raw=true)

<sub>_License: This software is provided gratis (free of charge), AS IS and without Warranty. It is free to use, and it is free to distribute **with prior written consent from the contributors/copyright holders** provided **no modifications are made**. Contributors retain their patents. All other rights reserved. Copyright &#169; 2019-2020, contributors and ClearPlan Consulting, LLC._</sub>

---

## Purpose
* The purpose of this project is to provide schedulers with a time-saving and error-free tool _to be used in support of industry best practices and **solid processes**_. (Read: don't be like the TSA conveyor belt.)
* It is **not** the purpose of this project to absolve schedulers of the duty to build, analyze, and maintain good schedules. No tool, and no automation, can replace the mind, experience, judgment, and instincts of a living, breathing human being. Therefore:
* Use only as directed. "When in doubt, read the directions."

---

## Table of Contents
1. [Prerequisites](#prerequisites)
1. [Installation](#installation)
1. [Use](#use)
    1. [View and Count Tools](#view-and-count-tools)
    1. [Text Tools](#text-tools)
    1. [Trace Tools](#trace-tools)
    1. [Status Tools](#status-tools)
        1. [Create Status Sheet](#create-status-sheet)
        1. [Update Status Sheet](#update-status-sheet)
        1. [Import Status Sheet](#import-status-sheet)
    1. [Resource Demand](#resource-demand)
    1. [QuickMetrics](#quickmetrics)
        1. [Schedule Metrics](#schedule-metrics)
        1. [EVish Metrics](#evish-metrics)
    1. [Backbone](#backbone)
    1. [Integration Tools](#integration-tools)
    1. [Data Dictionary](#data-dictionary)
1. [Help](#help)
1. [Credits](#credits)

[[top]](#table-of-contents)

---

## Prerequisites
1. Microsoft Project Standard or Professional 2013+ (32-bit or 64-bit, but must match Office version for some features to work properly)
1. Microsoft Office (Outlook, Excel, Word, PowerPoint) (32-bit or 64-bit, but must match MS Project version for some features to work properly)
1. Internet Connection preferred, but not required

[[top]](#table-of-contents)

---

## Installation
_Note: this tool is not designed to run "out of the box" on Microsoft Project Server/Online. Contact <a href="mailto:help@ClearPlanConsulting.com">help@ClearPlanConsulting.com</a> for further information or if you would like to explore a custom installation for your Server/Online environment._

1. Enable Macros
    1. Open Microsoft Project, go to File > Options > Trust Center > Trust Center Settings...
    1. In the **Macro Settings** pane, click **Enable all macros**
    ![installation-01](/images/installation-01.PNG?raw=true)
    1. In the **Legacy Formats** pane, click **Allow loading files with legacy or non-default formats** _(this will allow import/export of settings required for various features)_ 
    ![installation-02](/images/installation-02.PNG?raw=true)
    1. Click **OK** a couple of times to close the dialogs
    1. Completely exit, and then restart, Microsoft Project (this makes the settings above 'stick')
1. Download and open the [latest release](https://github.com/clearplan/cpt/releases/latest) (download the cpt_\*.mpp file).
1. Open the **Organizer** and select the **Modules** tab:
1. Copy **all** modules prefixed with `cpt` into your Global.MPT
1. On the Ribbon, click **View** > **Macros** > **View Macros** > and run the macro `cptSetup()`
    1. This macro installs necessary core modules if they are not already installed
    1. Changes will be made to your ThisProject module, but if you have existing code it will not be overwritten. cpt-related code will be inserted at the very top of the procedures **Project_Activate** and **Project_Open** and each line is appended with `'</cpt>` for reference.
1. The ClearPlan Toolbar will be added. Click **ClearPlan** > **Help** > **Help** > **Check for Upgrades** at any time to download the latest version(s). If you do not have an internet connection, please contact [help@ClearPlanConsulting.com](mailto:help@ClearPlanConsulting.com) for the latest hotfixes.
    
[[top]](#table-of-contents)

---

## Use

---

### View and Count Tools
![view](/images/view.png?raw=true)

1. _**Reset All**_ - removes all groups and filters, expands all tasks, and reorders by Task ID.
1. _**WrapItUp**_ - collapses the currently visible outline (or group) of your project starting from the lowest level up.
1. _**Count Tasks**_ - Just what it sounds like - optionally count all tasks, visible (filtered) tasks, or selected tasks. User is prompted with a tally of subtasks vs. summary tasks and includes inactive tasks separately. External and 'Nothing' Tasks are skipped.

[[up]](#use) | [[toc]](#table-of-contents)

---

### Text Tools
![text](/images/text.png?raw=true)

1. _**Advanced Text Tools**_ - a single dialog aggregating the text tools below. Presents a preview of changes and can be applied to any selected tasks. Undo works.
1. The following utilities are also available standalone in the top portion of the splitButton menu:
    1. _**Bulk Prepend**_ - does just what it says. Select some tasks, click the button, enter a prefix, done.
    1. _**Bulk Append**_ - same as above but adds a suffix instead of a prefix.
    1. _**Enumerate**_ - select some tasks, click _enumerate_. Enter how many characters your sequence should be (i.e., '3' = 000), and what number you wish to start by.
    1. _**MyReplace**_ - a better (and safer) find/replace tool that:
        1. Limits found/replaced texts to the selected tasks
        1. Includes all Text (and Outline Code) fields in the selection also.
1. Other tools:
    1. _**Trim Task Names**_ - removes all the pesky leading and trailing spaces when folks copy/paste from other tools.
    1. _**Replicate a Process**_ - WIP (future release)
    1. _**Find Duplicate Task Names**_ - finds duplicate task names and exports them into an Excel workbook.
    1. _**Reset Row Height**_ = for the type-A-er in you; when task rows auto-adjust and it gives you chest pains, click this button to restore your mental sanity and enter a limited state of ~ nirvana ~.
1. _**Dynamic Filter**_ - This 'find-as-you-type' tool is arguably the greatest addition ever made to Microsoft Project (but we are biased). It allows you to optionally 'pin' a given task while searching for others. Optionally highlight the results instead of filtering. Include Summaries, or related summaries, in your results. Press CTRL+BKSP to clear the filter and return to your originally selected task or click 'Clear All' to reset the file. This tool is particularly helpful when working with a team to link up a schedule.

[[up]](#use) | [[toc]](#table-of-contents)

---

### Trace Tools
![trace](/images/trace.png?raw=true)

1. _**Driving Path**_ - Select any target task, and then click this button. The Primary, Secondary, and Tertiary Driving Path will be displayed. _Note: as of v1.3 "`New Orleans`" this does not yet work with master/subprojects. We are targeting v1.5 "`Coronado`" for upgrade._
![drivingpath](/images/driving-path.png?raw=true)

1. _**Export to PowerPoint**_ - _With the target task still selected_, click this button to export this `VBA wizardry` into a PowerPoint slide deck. Each Driving Path (Primary, Secondary, and Tertiary) will start a new set of slides. If the task count is too high, the network will be split over multiple slides to keep it readable (anti-eye-charting). If multiple slides are required, the last task on (e.g.,) slide 1 will be repeated as the first task on the following slide to retain context and presentation flow.

[[up]](#use) | [[toc]](#table-of-contents)

---

### Status Tools
![status](/images/status.png?raw=true)

#### Create Status Sheet

![status-sheet-form](/images/status-form.png?raw=true)

1. Set a Project Status Date using Project > Project Information, or View > Status Date.
1. Open the form and select: which field EV% is stored in; which field Earned Value Technique (e.g., LOE, 50/50, etc.).
1. Choose whether to Hide Tasks Completed Before [a certain date]. E.g., set this to the previous status period to show tasks completed since the last update. If unchecked, the status sheet will include all incomplete tasks.
1. Choose whether to create: A Single Workbook (all tasks on one sheet); A Worksheet For Each (a single workbook, but with separate worksheets for each [CAM, CWBS, IPT, etc.]; or A Workbook For Each [CAM, WBS, IPT], etc.]. If you select one of the 'For Each' options, you can further refine your output by selecting one or more items in the selected custom field. E.g., if you select _A Workbook For Each_ and you designate the _CAM_ custom field, a separate workbook will be created for each selected CAM and that CAM's name will be included in the file name.
1. Optionally create an email with the created file attached. If you select _A Single Workbook_ or _A Worksheet For Each_ a single email will be created. If you select _A Workbook For Each_ a separate email will be created for each [item] in the selected custom field.
1. Optionally add custom fields between the UID and Task Name fields in the created workbook(s). Use the yellow search filter to find fields, and the right-arrow button to add it to your Export List. Use the double right-arrow to add all selected fields at once. Similarly, use the single or double left arrows to remove fields from your Export List. Use the plus and minus buttons to reorder the fields in your created workbook(s). Note that as you add/remove/reorder fields to your Export List, the current view will be updated to reflect your options--to provide you with a 'preview' of the created workbook(s). The structure of the tasks (summary levels) in the current view (whether native summary tasks or group by summaries) will be reflected in your created workbook(s).
1. When you have selected all of the above options, click **Create**. The yellow status bar will notify you of progress. Because we include many conditional formats to catch invalid data, the export may take a couple of minutes to complete.
1. Created workbooks will be saved to C:\Users\\[username]\CP_Status_Sheets\\[Project_Name]\\[Status Date]\\. All worksheets are protected; updates will only be permitted in empty _Actual Start/Actual Finish_ fields, _New Forecast Start/New Forecast Finish_ fields, the _New EV%_ field, the _New ETC_ field, and the _Reason / Action / Impact_ field, which are highlighted.
1. If you have a resource-loaded schedule, the status sheet will include editable Assignment-level cells for Remaining Work for Work (labor) resources, or Remaining Cost for Material resources.

_Notes:_ Your choices will be saved between sessions. These settings are saved locally, so are not shared between users. If a saved custom field does not exist in the currently active project, you will be prompted to remove it.

#### Update Status Sheet
![update-status-sheet](/images/status-update.png?raw=true)

1. Stakeholders should open the received Status Update workbooks and make changes to the editable fields (AS/AF, FS/FS, New EV%, New ETC, R/A/I). Built-in conditional formatting will catch common entry errors, such as: Actual Finish <> NA and EV% < 100% or ETC > 0; Actual Finish = NA and EV% = 100 or ETC = 0, etc.
1. Stakeholders may not add or remove tasks or assignments via the Status Sheet--this must (and should) only be done via the Scheduler in the live IMS.
1. If forecast dates change significantly the spreadsheet does not duplicate Total Slack calculations (that's what MS Project is for).
1. Collect returned status sheets in a new directory somewhere handy. You will be able to import one or many at a time.

_Notes:_ Your choices will be saved between sessions. These settings are saved locally, so are not shared between users. If a saved custom field does not exist in the currently active project, you will be prompted to remove it.

#### Import Status Sheet
The _Import Status Sheet_ feature allows you to automatically import updates to custom fields in your IMS for further review and analysis before being applied to the IMS. 

![status-import-form](/images/status-import-form.png?raw=true)

1. Open the Import Status Sheet Form.
1. Click _Select Files..._, navigate to the directory where returned status workbooks are stored, and make your selections. They will be listed in the ListBox on the form.
1. Select where to import Actual Start dates: any of the 10 local custom Start fields, the local custom 10 Date fields, or (careful!) choose to import directly to the Actual Start field. Likewise for Actual Finish: any of the 10 local custom Finish fields, the 10 local custom Date fields, or the Actual Finish fields are available. Do not select the same custom Date field for both Start and Finish.
1. Select where to import the _New EV%_: any of the 20 local custom Number fields are available.
1. Select where to import the _Assignment ETC_: any of the 20 local custom Number fields are available. (Note that these values will be imported to the _Assignment_ level.)
1. Optionally choose to import the _R/A/I_ values to the Task Notes. You may choose to prepend to, append to, or overwrite, existing Task Notes.
1. Save your file before importing in case there are any issues.
1. Click **Import**, change your view to Task Usage, and review changes. Once you have reviewed and approved the updates, manually update the values in the live fields (i.e., forecast start/finish, EV%, and Remaining Work), with the values in your selected local custom fields.

_Notes:_ Your import choices will be saved between sessions. These settings are saved locally, so are not shared between users. If a saved custom field does not exist in the currently active project, you will be prompted to reassign it before importing updates.

#### Smart Duration
1. "That task will be done on November 25th." "Ok...uh...let's start guessing how many days of remaining duration I need to hit that date..." Ever done that? NO MORE! Select a task, click this button, enter your finish date, and it will Figure Out The Things for you automagically. Please be sure to confirm remaining work after the magic subsides.

[[up]](#use) | [[toc]](#table-of-contents)

---

### Resource Demand
![resource](/images/resource-demand.png?raw=true)

![resource-form](/images/resource-demand-form.png?raw=true)

_**Resource Demand**_ - Export timephased remaining work to Excel (similar to the Task Usage view), with automatic PivotTable and PivotCharts. Add whatever fields you'd like to the report, and your settings will be saved for next time.

[[up]](#use) | [[toc]](#table-of-contents)

---

### QuickMetrics
![quick-metrics](/images/metrics.png?raw=true)

1. _**Schedule Metrics**_
    1. _CPLI_ - select a target task and click this to get a quick CPLI. 
    1. _BEI_ - Baseline Execution Index. Assuming your project is baselined, get a quick BEI.
    1. _Hit Task %_ - from the old Gold Card.
1. _**EVish Metrics**_ - the following "EVish" metrics assume that your schedule is resource-loaded (using MS Project Assignments), baselined, and that EV% is stored in the build-in _Physical % Complete_ field:
    1. _SPI_ - Ballpark Schedule Performance Index (in hours).
    1. _SV_ - Ballpark Schedule Variance (in hours).
    1. _BCWS_ - Budgeted Cost of Work Scheduled, a.k.a. Planned Value [PV] (in hours).
    1. _BCWP_ - Budgeted Cost of Work Performed, a.k.a. Earned Value [EV] (in hours).

[[up]](#use) | [[toc]](#table-of-contents)

---

### Backbone
This feature helps you set up and manage Outline Codes (e.g., CWBS, IMP).

![backbone](/images/backbone.png?raw=true)

![backbone-form](/images/backbone-form.png?raw=true)

1. Select the local custom Outline Code you want to work with. Existing lookup tables will be shown in the TreeView control below. If you make updates to the code while the form is open, the TreeView will refresh automatically.
1. Select the **Import Code** option and choose an _Import Source_. You can import: 
    1. A properly formatted Excel Workbook (*.xlsx). Required fields are CODE, LEVEL, and TITLE and must exist in range [A1:C1]. Be sure there are no empty rows of data in your import. To create a template for later import, click **Template**. Optionally, click _Also create task structure to automatically create summary/task structure. If unchecked, the import will only update the selected local custom Outline Code picklist.
    1. A generic MIL-STD-881D Appendix B (Electronic Systems). Note: this option will create new summaries and tasks in your active project. If this is not what you want, then run this in an empty *.mpp file, and simply import the resulting Outline Code from your empty *.mpp file into your live *.mpp file using the _Import_ button on MSP's built-in Custom Fields dialog.
    1. The existing summary/task structure in the active project. This will use the native WBS field as the code, and the Task Name as the Description.
    1. If the selected local custom Outline Code is not yet named, provide a name for it (e.g., "CWBS").
    1. Once you have made your selections, click **Import**.
1. Select the **Export Code** option to export the selected local custom Outline Code into: An Excel Workbook; "WBS Descriptive" CSV file for import to MPM; a CSV Code file for import to COBRA; or import into pre-formatted DI-MGMT-81334D for update. Note: the [DI-MGMT-81334D Template](https://github.com/clearplan/cpt/blob/develop/Templates/81334D_CWBS_TEMPLATE.xltm) is required; install it to your Microsoft Templates directory (usually located at C:\Users\\[username]\AppData\Roaming\Microsoft\Templates\).
1. Select the **Find/Replace** option to find and replace multiple instances of a word in the selected local custom Outline Code descriptions.

[[up]](#use) | [[toc]](#table-of-contents)

---

### Integration Tools
![integration](/images/integration.png?raw=true)

1. _**Export IMS to COBRA**_ - Creates the CSV-formatted import files for Baseline, Forecast, and Status for upload to COBRA. See the [Help Doc](https://github.com/clearplan/cpt/blob/develop/Trace/ClearPlan_IMS_Export_Utility_r3.1.0.docx) for more information.

[[up]](#use) | [[toc]](#table-of-contents)

---

### Data Dictionary
Store custom field descriptions and automatically generate an IMS Data Dictionary.

![data-dictionary](/images/data-dictionary.png?raw=true)

![data-dictionary](/images/data-dictionary-form.png?raw=true)

1. Local and Enterprise Custom fields will be listed in the ListBox. Use the yellow search filter to find a specific field.
1. Enter a Description in the Text Box. 
1. Optionally export the IMS Data Dictionary to an Excel workbook. If desired, update the descriptions and import them. If you wish to share your data dictionary descriptions with others, export your entries and send to another user with the CPT installed, and they can import them. CAVEAT: be sure local custom fields are aligned between your file and theirs.

_Note:_ Data Dictionary entries are stored _per project file_ and _per user_. 

[[up]](#use) | [[toc]](#table-of-contents)

---

## Help
![help](/images/help.png?raw=true)

1. _**Check for Upgrades**_ - From time to time click _Check for Upgrades_ to get the latest upgrades and hotfixes. _Note: if you do not have an internet connection available from your client computer, download on a non-client asset next time you're in the coffee-shop, then copy the modules over into your Organizer next time you're onsite._
1. _**Submit an Issue**_ - submit your bugs/issues for quick triage and (depending on availability of our developers who are themselves fully deployed) we'll push fixes ASAP.
1. _**Submit a Feature Request**_ - submit your feature requests for review! Suggestions, once reviewed, will be voted on at the next offsite!
1. _**Submit Other Feedback**_ - our entire goal is to help you _sharpen your axe_ so please let us know how we're doing, what you'd like changed, what's broken, etc.
1. _**Uninstall ClearPlan Toolbar**_ - Never do this. _Never_. (Seriously, why would you do this?) ...ok, if you must uninstall (client request, you're rolling off the gig, etc.) then click this button and follow the prompts.

[[up]](#help) | [[toc]](#table-of-contents)

---

## Credits
A special thanks to our developers and contributors, and to the ClearPlan leadership team for inspiring us to `make it better`.

[[up]](#help) | [[toc]](#table-of-contents)

---

<sub>Copyright &#169; 2019-2026 ClearPlan Consulting, LLC. All Rights Reserved.</sub>
