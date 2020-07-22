
Public usernameP As String
Dim BaseURL As String
Dim HTTP_ErrorCode
Dim runMode As String

'column definitions
Const Simple_CardAddress = 6
Const Simple_Lane = 1
Const Simple_IncidentNumber = 2
Const Simple_Title = 3
Const Simple_Status = 4
Const Simple_StoryPoints = 5

Const Full_Creator = 1
Const Full_CreatedDT = 2
Const Full_UpdatedDT = 3
Const Full_Title = 4
Const Full_NoteBodyText = 5
Const Full_CardURL = 6
Const Full_Type = 7
Const Full_State = 8
Const Full_AssignedTo = 9
Const Full_Labels = 10


'2020-05-27 RAJ - Updating to allow for multiple options, for Dr. Anderson's benefit.
'Simple option will only display Incident number - extracted from short desc, story points - extracted from labels,
'Short description, and status.


 
 



'Using https://github.com/VBA-tools/VBA-JSON

Public Sub SetRunMode()
'sets the runmode
    'Debug.Print Sheet1.getRunMode
    runMode = Sheet1.getRunMode
    
End Sub


Public Function GetProjectsByOrg(strOrgName As String) As String
    'https://www.codeproject.com/Articles/1088523/Excel-Jira-Rest-API-end-to-end-example
    Dim GitHubAPI As New MSXML2.XMLHTTP
    Dim Json As Object
    Dim pageJSON As String
    Dim allJSON As String
    Dim finalJSON As String
    Dim URL As String
    Dim page As Integer
    For page = 1 To 99 'and hope we never hit 99000 boards....
        With GitHubAPI
            URL = BaseURL + "/orgs/" _
                + strOrgName _
                + "/projects" _
                + "?per_page=100" _
                + "&page=" + CStr(page)
            .Open "GET", URL, False
            
            .setRequestHeader "Accept", "application/vnd.github.inertia-preview+json"
            .setRequestHeader "Authorization", "Basic " + usernameP
            .send ""
        End With
        
        'remove leading and trailing square brackets, they'll be re-added after final page is processed. Should be first and last characters.
        pageJSON = GitHubAPI.responseText
        
        
        pageJSON = Mid(pageJSON, 2, Len(pageJSON) - 2)
        
        Debug.Print ("pageJSON" + vbCr + Right(pageJSON, 255))
        
        allJSON = allJSON + pageJSON
        
        If GitHubAPI.responseText = "[]" Then
            page = 99
        End If
        
        Debug.Print ("allJSON" + vbCr + Right(allJSON, 255))
    Next page
    
    'add leading and trailing square brackets
    finalJSON = "[" + allJSON + "]"
    
    
    GetProjectsByOrg = finalJSON
    
    
End Function

Function GetProjectsByRepo(strOrgName As String, strRepo As String) As String
    Dim GitHubAPI As New MSXML2.XMLHTTP
    Dim Json As Object
    Dim URL As String
    With GitHubAPI
        URL = BaseURL + "/repos/" _
            + strOrgName + "/" _
            + strRepo _
            + "/projects" _
            + "?per_page=100"                    '100 is max.

        .Open "GET", URL, False
        
        .setRequestHeader "Accept", "application/vnd.github.inertia-preview+json"
        .setRequestHeader "Authorization", "Basic " + usernameP
        .send ""
    End With
    
    GetProjectsByRepo = GitHubAPI.responseText
End Function

Function getColumnsByProject(projectID As String) As String
    Dim GitHubAPI As New MSXML2.XMLHTTP
    Dim Json As Object
    Dim URL As String
    With GitHubAPI
        URL = BaseURL + "/projects/" _
            + projectID _
            + "/columns"

        .Open "GET", URL, False
        
        .setRequestHeader "Accept", "application/vnd.github.inertia-preview+json"
        .setRequestHeader "Authorization", "Basic " + usernameP
        .send ""
    End With
    
    getColumnsByProject = GitHubAPI.responseText
End Function

Function getCardsByColumn(columnID As String) As String
    Dim GitHubAPI As New MSXML2.XMLHTTP
    Dim Json As Object
    Dim URL As String
    With GitHubAPI
        URL = BaseURL + "/projects/columns/" _
            + columnID _
            + "/cards"

        .Open "GET", URL, False
        
        .setRequestHeader "Accept", "application/vnd.github.inertia-preview+json"
        .setRequestHeader "Authorization", "Basic " + usernameP
        .send ""
    End With
    
    getCardsByColumn = GitHubAPI.responseText
End Function

Function getIssueByIssueURL(issueURL As String) As String
    Dim GitHubAPI As New MSXML2.XMLHTTP
    Dim Json As Object
    Dim URL As String
    With GitHubAPI
        URL = issueURL

        .Open "GET", URL, False
        
        .setRequestHeader "Accept", "application/vnd.github.inertia-preview+json"
        .setRequestHeader "Authorization", "Basic " + usernameP
        .send ""
    End With
    
    getIssueByIssueURL = GitHubAPI.responseText
End Function

Public Sub SetPasswordString()
    Dim UID, PAT, UID_PAT As String
    
    UID = ThisWorkbook.Sheets("Config").Range("B1")
    PAT = ThisWorkbook.Sheets("Config").Range("B2")
    UID_PAT = UID + ":" + PAT
    
    
    usernameP = TextBase64Encode(UID_PAT, "ASCII")
End Sub

Public Sub OnRun()
    'check run mode is valid
    SetRunMode
    If runMode <> "-----" Then
        'create password string
        SetPasswordString
        
        'set Base URL
        BaseURL = "https://api.github.com"
        
        'going to create a new sheet. Likely less confusing.
        Dim OutputWorkbook As Workbook
        Set OutputWorkbook = Workbooks.Add
        

    
        

        
        '----find project ID---------
        Dim projectID As String
        Dim projectsObj
        'determine url peices: either in a repo (no org) or not (has org in url).
        Dim BoardURL As String
        Dim sOrg As String
        Dim sRepo As String
        Dim startPos As Integer, firstSlashPos As Integer, secondSlashPos As Integer, thirdSlashPos As Integer
        BoardURL = ThisWorkbook.Sheets("Config").Range("B3")
        
        'get pos of "github.com" for startpos
        startPos = InStr(1, BoardURL, "github.com", vbTextCompare)
        
        'first slash pos (starts at 1):
        firstSlashPos = InStr(startPos, BoardURL, "/", vbTextCompare)
            
        'second slash pos:
        secondSlashPos = InStr(firstSlashPos + 1, BoardURL, "/", vbTextCompare)
            
        'third slash pos
        thirdSlashPos = InStr(secondSlashPos + 1, BoardURL, "/", vbTextCompare)
        
        If InStr(1, BoardURL, "/orgs/", vbTextCompare) = 0 Then
            '/orgs/is not found.
            'if not found, need to go by Org, and Repo name.
            'org is after first slash until the second, repo is after the second to the third.
            sOrg = Mid(BoardURL, firstSlashPos + 1, secondSlashPos - firstSlashPos - 1)
            sRepo = Mid(BoardURL, secondSlashPos + 1, thirdSlashPos - secondSlashPos - 1)
            
            'get the projectsobj
            Set projectsObj = processJSONtoJSONObject(GetProjectsByRepo(sOrg, sRepo))
            
        Else
            '/orgs/ is found. Can search
            'get org. its the bit between the first and second slash.
            
        
            'bit inbetween
            sOrg = Mid(BoardURL, secondSlashPos + 1, thirdSlashPos - secondSlashPos - 1)
            
            'get the projectsObj
            Set projectsObj = processJSONtoJSONObject(GetProjectsByOrg(sOrg))
        
        End If
        
        'get Projects by Org
        ' GetProjectsByOrg (ThisWorkbook.Sheets("Config").Range("B3"))
        'convert projects to json
        
        
        'get project ID for url specified
        'loop through the responses for the one with the specific url.
        Dim p As Integer
        
        'if https is not at the front of the url, add it
        If InStr(1, BoardURL, "https://", vbTextCompare) = 0 Then
            BoardURL = "https://" + BoardURL
        End If
        
        
        For p = 1 To projectsObj.count
            If projectsObj(p)("html_url") = BoardURL Then
                projectID = projectsObj(p)("id")
                Exit For
            End If
                
        Next p
        
        'report if project not found
        If projectID = "" Then
        MsgBox ("Project not found")
        Exit Sub
        End If
        
        'get the columns for that project.
        'each column gets its own worksheet, named with the column name.
        'we'll loop through those columns, and within those columns, create the rows for the card info.
        Dim columnsObj As Variant
        Dim colIndex As Integer
        Dim colName As String
        Dim colID As String
        
        Dim cardsObj As Variant
        
        Dim isIssue As Boolean
        Dim issueURL As String
        Dim issueObj As Variant
        
        Dim labelIndex As Integer
        '-------Get Columns for project board
        Set columnsObj = processJSONtoJSONObject(getColumnsByProject(projectID))
        
        
        '-------Determine Mode: simple or original: simple just gets cards: per Dr. Anderson's request.
        If runMode = "Issues Only" Then
            'card ID with hyperlink.
            'column name
            'INC number
            'short description
            'story points
            'status
            
            'all on one sheet
            
            
                
            'create the worksheet
            Dim simpleWS As Worksheet
            Set simpleWS = OutputWorkbook.Worksheets(1)
            
            simpleWS.Name = "Issues"
            
            Dim startRow As Integer
            startRow = 1
                
            'set sheet columns
            OutputWorkbook.Sheets("Issues").Cells(startRow, Simple_CardAddress) = "Card Address"
            OutputWorkbook.Sheets("Issues").Cells(startRow, Simple_Lane) = "Lane Name"
            OutputWorkbook.Sheets("Issues").Cells(startRow, Simple_IncidentNumber) = "Incident Number"
            OutputWorkbook.Sheets("Issues").Cells(startRow, Simple_Title) = "Short Description"
            OutputWorkbook.Sheets("Issues").Cells(startRow, Simple_Status) = "State"
            OutputWorkbook.Sheets("Issues").Cells(startRow, Simple_StoryPoints) = "Story Points"
                    
                
            'still need to go through each column to get the cards.
            
            
            For colIndex = 1 To columnsObj.count
                colID = columnsObj(colIndex)("id")
                Set cardsObj = processJSONtoJSONObject(getCardsByColumn(colID))
                
                Dim cardIndex As Integer
                
                colName = columnsObj(colIndex)("name")
                

                For cardIndex = 1 To cardsObj.count
                    
                    'check if note or issue by looking for content_url key (has key, is issue)
                    If cardsObj(cardIndex).Exists("content_url") Then
                        'is issue
                        'need to get issue info by issue ID.
                        issueURL = cardsObj(cardIndex)("content_url")
                        
                        Set issueObj = processJSONtoJSONObject(getIssueByIssueURL(issueURL))
                        
                        'get INC number
                        'INC Number is always INC followed by some number of numbers.
                        Dim INCNumber As String
                        Dim issueTitle As String
                        issueTitle = issueObj("title")
                        INCNumber = getINCNumberFromShortDescription(issueTitle)
                        
                        'get Story Points
                        Dim hasStoryPoints As Boolean
                        hasStoryPoint = False
                        
                        Dim storyPoints As String
                        Dim maxStoryPoints As String
                        
                        
                        If issueObj("labels").count <> 0 Then
                            Dim label As String
                            
                            For labelIndex = 1 To issueObj("labels").count
                                label = issueObj("labels")(labelIndex)("name")
                                hasStoryPoints = isStoryPoints(label)
                                If hasStoryPoints Then
                                    storyPoints = getStoryPoints(label)
                                    If maxStoryPoints < storyPoints Then
                                        maxStoryPoints = storyPoints
                                    End If
                                    
                                End If
                            Next labelIndex
                            
                        End If
                        
                        Dim cardLink As String
                        Dim cardNum As String
                        
                        cardLink = issueObj("html_url")
                        cardNum = issueObj("number")
                        
                        'output data
                        OutputWorkbook.Sheets("Issues").Cells(startRow + cardIndex, Simple_CardAddress) = cardLink
                        OutputWorkbook.Sheets("Issues").Cells(startRow + cardIndex, Simple_Lane) = colName
                        OutputWorkbook.Sheets("Issues").Cells(startRow + cardIndex, Simple_IncidentNumber) = INCNumber
                        OutputWorkbook.Sheets("Issues").Cells(startRow + cardIndex, Simple_Title) = issueTitle
                        OutputWorkbook.Sheets("Issues").Cells(startRow + cardIndex, Simple_Status) = issueObj("state")
                        OutputWorkbook.Sheets("Issues").Cells(startRow + cardIndex, Simple_StoryPoints) = storyPoints
                    Else
                        'is note
                        'skipping these on Issues only run.
                        
                    End If 'end if issue...
                    
                Next cardIndex
                startRow = startRow + cardIndex - 1
            Next colIndex
            
            'autosize
            Dim colRange As Range
            For Each colRange In OutputWorkbook.Sheets("Issues").UsedRange.Columns
                colRange.AutoFit
            Next
            
            'unselect all
            
            
        '---------------Isn't Dr. Anderson's mode
        Else
        'loop through the columns and create a new worksheet

            
            For colIndex = 1 To columnsObj.count
                'get the column name, removing invalid Chars up to limit =31.
                'invalid chars = \ , / , * , ? , : , [ , ]
                'replacing with _

                colName = columnsObj(colIndex)("name")
                colID = columnsObj(colIndex)("id")
                Dim InvalidChars() As Variant
                Dim InvalidChar As Variant
                InvalidChars = Array("\", "/", "*", "?", ":", "[", "]")
                For Each InvalidChar In InvalidChars
                    colName = Replace(colName, InvalidChar, "_")
                Next InvalidChar
                'length
                colName = Left(colName, 31)
                
                'add a new sheet
                Dim ws As Worksheet
                With OutputWorkbook
                    Set ws = .Sheets.Add(After:=.Sheets(.Sheets.count))
                    ws.Name = colName
                End With
                'add column headers
                'creator, created datetime, updated dt, note text, card url, type, state (open, closed), labels (comma sep)
                OutputWorkbook.Sheets(colName).Cells(1, Full_Creator) = "Creator" '1
                OutputWorkbook.Sheets(colName).Cells(1, Full_CreatedDT) = "Created" '2
                OutputWorkbook.Sheets(colName).Cells(1, Full_UpdatedDT) = "Updated" '3
                 OutputWorkbook.Sheets(colName).Cells(1, Full_Title) = "Title" '4
                OutputWorkbook.Sheets(colName).Cells(1, Full_NoteBodyText) = "Note/Body Text" '5
                OutputWorkbook.Sheets(colName).Cells(1, Full_CardURL) = "Card URL" '6
                OutputWorkbook.Sheets(colName).Cells(1, Full_Type) = "Type" '7
                OutputWorkbook.Sheets(colName).Cells(1, Full_State) = "State" '8
                OutputWorkbook.Sheets(colName).Cells(1, Full_AssignedTo) = "Assigned To" '9
                OutputWorkbook.Sheets(colName).Cells(1, Full_Labels) = "Labels" '10
                'fix width for note body text
                OutputWorkbook.Sheets(colName).Cells(1, Full_Title).ColumnWidth = 60
                OutputWorkbook.Sheets(colName).Cells(1, Full_NoteBodyText).ColumnWidth = 84
                OutputWorkbook.Sheets(colName).Cells(1, Full_NoteBodyText).WrapText = True
                
                'add the rows/cards.
                
                'get the cards by column id.
                Set cardsObj = processJSONtoJSONObject(getCardsByColumn(colID))
                
                
                
                For cardIndex = 1 To cardsObj.count
                    'check if note or issue by looking for content_url key (has key, is issue)
                    If cardsObj(cardIndex).Exists("content_url") Then
                        'is issue
                        'need to get issue info by issue ID.
                        
                        issueURL = cardsObj(cardIndex)("content_url")
                        
                        
                        Set issueObj = processJSONtoJSONObject(getIssueByIssueURL(issueURL))
                        
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 1) = issueObj("user")("login")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 2) = issueObj("created_at")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 3) = issueObj("updated_at")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 4) = issueObj("title")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 5) = issueObj("body")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 6) = cardsObj(cardIndex)("content_url")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 7) = "Issue"
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 8) = issueObj("state")
                        'possibility of multiple assignees
                        If issueObj("assignees").count <> 0 Then
                                'loop over assignees
                                Dim a As Integer
                                Dim assignees As String
                                assignees = ""
                                For a = 1 To issueObj("assignees").count
                                    assignees = assignees & issueObj("assignees")(a)("login") & vbCrLf
                                Next a
                                OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 9) = assignees
                        End If
                        'labels
                        If issueObj("labels").count <> 0 Then
                            Dim labels As String
                            labels = ""
                            For labelIndex = 1 To issueObj("labels").count
                                labels = labels & issueObj("labels")(labelIndex)("name") & vbCrLf
                            Next labelIndex
                            OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 10) = labels
                        End If
                        
                        
                    Else
                        'is note
                        'data populated
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 1) = cardsObj(cardIndex)("creator")("login")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 2) = cardsObj(cardIndex)("created_at")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 3) = cardsObj(cardIndex)("updated_at")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 5) = cardsObj(cardIndex)("note")
                        OutputWorkbook.Sheets(colName).Cells(cardIndex + 1, 7) = "Note"
                        
                        
                    End If
                    
                Next cardIndex
                
            Next colIndex
            
            'delete the remaining Sheet1 sheet
        'delete any sheets named Sheet1
        Dim tSheets As Variant
        For Each tSheets In OutputWorkbook.Sheets
            If tSheets.Name = "Sheet1" Then
                Application.DisplayAlerts = False
                tSheets.Delete
                Application.DisplayAlerts = True
            End If
        Next tSheets
        End If
    
    
    Else 'if it's not a valid run mode selection
        MsgBox "Please select a run mode."
    
    End If
    

    
    
    Debug.Print ("done")
    
    
    
End Sub
Function isStoryPoints(labelText As String) As Boolean
    'https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops.
    Dim strPattern As String: strPattern = "[0-9] pts|[0-9] pt"
    Dim regEx As New RegExp
    Dim strInput As String
    Dim Myrange As Range
    Dim matches As Object
    

    If strPattern <> "" Then
        strInput = labelText

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            'MsgBox (regEx.Replace(strInput, strReplace))
            isStoryPoints = True
            
        Else
            isStoryPoints = False
        End If
    End If
End Function

Sub testisStoryPoints()
    Dim shortDesc As String
    Dim output As Boolean
    shortDesc = "5 pt"
    output = isStoryPoints(shortDesc)
    Debug.Print output
End Sub
Sub testgetStoryPoints()
    Dim shortDesc As String
    Dim output As String
    shortDesc = "5 pt"
    output = getStoryPoints(shortDesc)
    Debug.Print output
End Sub

Function getStoryPoints(labelText As String) As String
    Dim strPattern As String: strPattern = "[0-9]"
    Dim regEx As New RegExp
    Dim strInput As String
    Dim Myrange As Range
    Dim matches As Object
    

    If strPattern <> "" Then
        strInput = labelText

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            'MsgBox (regEx.Replace(strInput, strReplace))
            Set matches = regEx.Execute(strInput)
            'Debug.Print matches.Item(0)
            getStoryPoints = matches.Item(0)
        Else
            
            Debug.Print "Not matched"
            getStoryPoints = ""
        End If
    End If
End Function

Function getINCNumberFromShortDescription(shortDescription As String) As String
    'https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops.
    Dim strPattern As String: strPattern = "INC[0-9]*"
    Dim regEx As New RegExp
    Dim strInput As String
    Dim Myrange As Range
    Dim matches As Object
    

    If strPattern <> "" Then
        strInput = shortDescription

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            'MsgBox (regEx.Replace(strInput, strReplace))
            Set matches = regEx.Execute(strInput)
            Debug.Print matches.Item(0)
            getINCNumberFromShortDescription = matches.Item(0)
        Else
            
            Debug.Print "Not matched"
            getINCNumberFromShortDescription = "No INC # found"
        End If
    End If
End Function

Sub testGetINCNumber()
    Dim shortDesc As String
    Dim output As String
    shortDesc = "INC10786438 - User Story - VDIF CDA Documents: Add Discrete Lab Details"
    output = getINCNumberFromShortDescription(shortDesc)
    Debug.Print output
End Sub

Function processJSONtoJSONObject(jsonTxt As String) As Variant
    'converts json text to a json object.

    Dim parsedJson As Variant
    Set parsedJson = JsonConverter.ParseJson(jsonTxt)
    
    Set processJSONtoJSONObject = parsedJson
End Function

Function TextBase64Encode(strText, strCharset)

    Dim arrBytes

    With CreateObject("ADODB.Stream")
        .Type = 2                                ' adTypeText
        .Open
        .Charset = strCharset
        .WriteText strText
        .Position = 0
        .Type = 1                                ' adTypeBinary
        arrBytes = .Read
        .Close
    End With

    With CreateObject("MSXML2.DOMDocument").createElement("tmp")
        .DataType = "bin.base64"
        .nodeTypedValue = arrBytes
        TextBase64Encode = Replace(Replace(.Text, vbCr, ""), vbLf, "")
    End With

End Function


Private Sub Worksheet_FollowHyperlink(ByVal Target As hyperlink)
    Debug.Print Target
End Sub



