This Excel Project uses macros to access a github project board, and convert it to a series of WorkSheets based on the project kanban board columns (lanes). It imports data based on the datatype (note or issue) and fills in columns on the sheet.

It's 100% not perfect, and was intended for a quick project. I'm continuing to work on this in my spare time. USE AT YOUR OWN RISK.

--Installation --
To install, first import the project at https://github.com/VBA-tools/VBA-JSON. This allows for the conversion from JSON to VBA objects.

Next, copy the text of the ThisWorkbook module into your thisWorkbook module in VBA editor.

Finally, setup the worksheet, titled "Config" with three rows:

A	B
1	Username:	
2	Personal Access Token:	
3	URL of Board:	
---Configuration for Running---
To configure the macro, you must first generate a Personal Access Token on github, the following link has more information: https://help.github.com/en/github/authenticating-to-github/creating-a-personal-access-token-for-the-command-line The access I used is: admin:org, admin:repo_hook, read:packages, repo, user, but you can select all if you'd like.

Enter your github username in the username field, and the token you generated in the last step in the Personal Access Token field. The URL of the board is the full URL of the Kanban Board.

DO NOT CHANGE THE ORDER OF THESE VALUES.

---Usage---
Upon opening the file, you may be prompted to enable macros for this sheet. Click Enable Content.

To run the report, select the run mode and click Run. A new sheet will open for saving. Full is the original mode, getting all information about the board. Issues Only mode places all issues on the board in a single spreadsheet, with only Status, Incident Number, Short Description, State, Story Points, and Card Address.

To run as a Milestone Listing, enter the Milestone URL, and select "Milestone" from the dropdown, and click Run. Has the same information as the Issues Only report, adding a column for the sprint.

!!! The Personal Access Token should be kept private as it's essentially a password.
