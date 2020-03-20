
Public usernameP As String
Dim BaseURL As String
Dim HTTP_ErrorCode
 
 



'Using https://github.com/VBA-tools/VBA-JSON

Public Function GetProjectsByOrg(strOrgName As String) As String
    'https://www.codeproject.com/Articles/1088523/Excel-Jira-Rest-API-end-to-end-example
    Dim GitHubAPI As New MSXML2.XMLHTTP
    Dim Json As Object
    Dim URL As String
    With GitHubAPI
        URL = BaseURL + "/orgs/" _
            + strOrgName _
            + "/projects" _
            + "?per_page=100"                    '100 is max.

        .Open "GET", URL, False
        
        .setRequestHeader "Accept", "application/vnd.github.inertia-preview+json"
        .setRequestHeader "Authorization", "Basic " + usernameP
        .send ""
    End With
    
    GetProjectsByOrg = GitHubAPI.responseText
    
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
    'create password string
    SetPasswordString
    
    'set Base URL
    BaseURL = "https://api.github.com"
    
    'delete any sheets not named Config
    Dim tSheets As Variant
    For Each tSheets In ThisWorkbook.Sheets
        If Not tSheets.Name = "Config" Then
            Application.DisplayAlerts = False
            tSheets.Delete
            Application.DisplayAlerts = True
        End If
    Next tSheets
    
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
    MsgBox ("Projet not found")
    Exit Sub
    End If
    
    '-------Get Columns for project board
    
    'next, get the columns for that project.
    'each column gets its own worksheet, named with the column name.
    'we'll loop through those columns, and within those columns, create the rows for the card info.
    Dim columnsObj As Variant
    Set columnsObj = processJSONtoJSONObject(getColumnsByProject(projectID))
    
    'loop through the columns and create a new worksheet
    Dim col As Integer
    Dim colName As String
    Dim colID As String
    For col = 1 To columnsObj.count
        'get the column name, removing invalid Chars up to limit =31.
        'invalid chars = \ , / , * , ? , : , [ , ]
        'replacing with _

        colName = columnsObj(col)("name")
        colID = columnsObj(col)("id")
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
        With ThisWorkbook
            Set ws = .Sheets.Add(After:=.Sheets(.Sheets.count))
            ws.Name = colName
        End With
        'add column headers
        'creator, created datetime, updated dt, note text, card url, type, state (open, closed), labels (comma sep)
        ThisWorkbook.Sheets(colName).Range("A1") = "Creator" '1
        ThisWorkbook.Sheets(colName).Range("B1") = "Created" '2
        ThisWorkbook.Sheets(colName).Range("C1") = "Updated" '3
         ThisWorkbook.Sheets(colName).Range("D1") = "Title" '4
        ThisWorkbook.Sheets(colName).Range("E1") = "Note/Body Text" '5
        ThisWorkbook.Sheets(colName).Range("F1") = "Card URL" '6
        ThisWorkbook.Sheets(colName).Range("G1") = "Type" '7
        ThisWorkbook.Sheets(colName).Range("H1") = "State" '8
        ThisWorkbook.Sheets(colName).Range("I1") = "Assigned To" '9
        ThisWorkbook.Sheets(colName).Range("J1") = "Labels" '10
        'fix width for note body text
        ThisWorkbook.Sheets(colName).Range("D1").ColumnWidth = 60
        ThisWorkbook.Sheets(colName).Range("E1").ColumnWidth = 84
        ThisWorkbook.Sheets(colName).Range("E1").WrapText = True
        
        'add the rows/cards.
        
        'get the cards by column id.
        Dim cardsObj As Variant
        Set cardsObj = processJSONtoJSONObject(getCardsByColumn(colID))
        
        
        Dim card As Integer
        Dim isIssue As Boolean
        
        For card = 1 To cardsObj.count
            'check if note or issue by looking for content_url key (has key, is issue)
            If cardsObj(card).Exists("content_url") Then
                'is issue
                'need to get issue info by issue ID.
                Dim issueURL As String
                issueURL = cardsObj(card)("content_url")
                
                Dim issueObj As Variant
                Set issueObj = processJSONtoJSONObject(getIssueByIssueURL(issueURL))
                
                ThisWorkbook.Sheets(colName).Cells(card + 1, 1) = issueObj("user")("login")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 2) = issueObj("created_at")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 3) = issueObj("updated_at")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 4) = issueObj("title")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 5) = issueObj("body")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 6) = cardsObj(card)("content_url")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 7) = "Issue"
                ThisWorkbook.Sheets(colName).Cells(card + 1, 8) = issueObj("state")
                'possibility of multiple assignees
                If issueObj("assignees").count <> 0 Then
                        'loop over assignees
                        Dim a As Integer
                        Dim assignees As String
                        assignees = ""
                        For a = 1 To issueObj("assignees").count
                            assignees = assignees & issueObj("assignees")(a)("login") & vbCrLf
                        Next a
                        ThisWorkbook.Sheets(colName).Cells(card + 1, 9) = assignees
                End If
                'labels
                If issueObj("labels").count <> 0 Then
                    Dim l As Integer
                    Dim labels As String
                    labels = ""
                    For l = 1 To issueObj("labels").count
                        labels = labels & issueObj("labels")(l)("name") & vbCrLf
                    Next l
                    ThisWorkbook.Sheets(colName).Cells(card + 1, 10) = labels
                End If
                
                
            Else
                'is note
                'data populated
                ThisWorkbook.Sheets(colName).Cells(card + 1, 1) = cardsObj(card)("creator")("login")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 2) = cardsObj(card)("created_at")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 3) = cardsObj(card)("updated_at")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 5) = cardsObj(card)("note")
                ThisWorkbook.Sheets(colName).Cells(card + 1, 7) = "Note"
                
                
            End If
            
        Next card
        
    Next col
    
    Debug.Print ("done")
    
    
    
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



