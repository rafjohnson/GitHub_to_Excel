This Excel project gets data from either a Kanban Project board, or a Milestone list (Backlog list) and 
adds it to an Excel Workbook (with 1 to many sheets). 

## ---Configuration for Running---
To configure the macro, you must first generate a Personal Access Token on github, the 
following link has more information: 
https://help.github.com/en/github/authenticating-to-github/creating-a-personal-access-token-for-the-command-line
The access I used is:  admin:org, admin:repo_hook, read:packages, repo, user, but you 
can select all if you'd like.

Enter your github username in the username field, and the token you generated in the last
step in the Personal Access Token field.  The URL of the board is 
the full URL of the Kanban Board. The URL of the Milestone is the full URL of the backlog or 
milestone list. 

DO NOT CHANGE THE ORDER OF THESE VALUES. 



## ---Usage---
Download the xlsm file shown above. 
Upon opening the file, you may be prompted to enable macros for this sheet. Click Enable Content.

To run the report, select the run mode and click Run. A new sheet will open for saving. Full is 
the original mode, getting all information about the board. Issues Only mode places all issues 
on the board in a single spreadsheet, with only Status, Incident Number, Short Description, 
State, Story Points, and Card Address.  Issues will display in the order of the board. 

To run as a Milestone Listing, enter the Milestone URL, and select "Milestone" from the dropdown, and 
click Run. Has the same information as the Issues Only report, adding a column for the sprint. Milestone
items will generally display in the date order created. The order of the Milestone list has no impact on 
the order of the spreadsheet. 

!!! The Personal Access Token should be kept private as it's essentially a password.
