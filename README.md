# Agent Report Generator v0.1

This was a work project I made during an internship. The project takes inputs from the user and generates a report for specified call center agents. There is an enormous amount of data to sift through and this tool grabs the relevant parts and displays them in an easier-to-read style. It was built in an Excel spreadsheet using VBA and it would query multiple databases. 

A lot of actions are occurring in the background that are hidden to the user. All of the functions rely on a single table that contains every static information (e.g. query statements, table names, column names, etc.). This way the values just need to be changed in one place in future versions. 

The main parts are outlined in the images below. I don't really know if I can post the code here, legally, so this repo is just a snapshot of the whole project.


The form used to select criteria:
![reportform](https://user-images.githubusercontent.com/33203865/36357407-4b6b099c-14c3-11e8-90d6-e6b3f3bd98e8.PNG)



Next part of selection. The other tabs are basically the same, but shows different data.

![reportform2](https://user-images.githubusercontent.com/33203865/36357425-81aec5ca-14c3-11e8-8a98-ba34eb9e3a75.PNG)



The report layout that is generated:

![mainreport](https://user-images.githubusercontent.com/33203865/36357454-d74ddfca-14c3-11e8-9a75-dda83c9eff8e.PNG)



Individual Agent's report with their data only, which compares to goals depending on the role selected.

![agentreport](https://user-images.githubusercontent.com/33203865/36357467-00ac98a2-14c4-11e8-9dab-d913d6387ea6.PNG)



Another table that reports on the Agents' "idle" time actions.

![notreadyreport](https://user-images.githubusercontent.com/33203865/36357490-5602a3a0-14c4-11e8-84fa-24f6785bf3b9.PNG)
