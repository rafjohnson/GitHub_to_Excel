This Excel Project uses macros to access a github project board, and convert it to 
a series of WorkSheets based on the project kanban board columns (lanes). It imports 
data based on the datatype (note or issue) and fills in columns on the sheet. 

It's 100% not perfect, and was intended for a quick project. I'm continuing to work on this in my spare time.
USE AT YOUR OWN RISK.

## --Installation --
To install, first import the project at https://github.com/VBA-tools/VBA-JSON. This allows for the conversion from JSON
to VBA objects.

Next, copy the text of the ThisWorkbook module into your thisWorkbook module in VBA editor. 

Finally, setup the worksheet, titled "Config" with three rows:
|   |           A            |               B                |
|---|------------------------|--------------------------------|
| 1 | Username:              | <github username>              |
| 2 | Personal Access Token: | <github Personal Access Token> |
| 3 | URL of Board:          | <board URL>                    |

## ---Configuration for Running---
To configure the macro, you must first generate a Personal Access Token on github, the 
following link has more information: 
https://help.github.com/en/github/authenticating-to-github/creating-a-personal-access-token-for-the-command-line
The access I used is:  admin:org, admin:repo_hook, read:packages, repo, user, but you 
can select all if you'd like.

Enter your github username in the username field, and the token you generated in the last
step in the Personal Access Token field.  The URL of the board is 
the full URL of the Kanban Board. 

DO NOT CHANGE THE ORDER OF THESE VALUES. 



## ---Usage---
Upon opening the file, you may be prompted to enable macros for this sheet. Click Enable Content.

To kick off the macro, you can go to the developer tab in Excel, and click Macros. 
Select ThisWorkbook.OnRun macro, and click run. The sheet will populate with the board and 
card info. If you do not have the developer tab, right click on the ribbon (the area with File, Home,
Insert) and select "Customize Ribbon". Click "Developer" in the right side box and click Ok.

!!! Remember to remove the Config tab and save as XLSX before emailing out. Most orgs will not allow emailing
of XLSM (macro enabled workbooks), and the Personal Access Token should be kept private as it's 
essentially a password.

