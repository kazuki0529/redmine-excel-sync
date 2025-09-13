VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Issue クラス
'------------------------------------------------------------------------------
' Redmine のチケット情報を表すクラス。
'==============================================================================
Private pData As Dictionary
Private pId As Long
Private pProject As String
Private pTracker As String
Private pStatus As String
Private pPriority As String
Private pAuthor As String
Private pAssignedTo As String
Private pFixedVersion As String
Private pParentId As Long
Private pSubject As String
Private pDescription As String
Private pStartDate As Date
Private pDueDate As Date
Private pDoneRatio As Long
Private pCustomFields As Dictionary
Private pCreatedOn As Date
Private pUpdatedOn As Date

Property Get Data() As Dictionary
    Set Data = pData
End Property
Property Get Id() As Long
    Id = pId
End Property
Property Get Project() As String
    Project = pProject
End Property
Property Get Tracker() As String
    Tracker = pTracker
End Property
Property Get Status() As String
    Status = pStatus
End Property
Property Get Priority() As String
    Priority = pPriority
End Property
Property Get Author() As String
    Author = pAuthor
End Property
Property Get AssignedTo() As String
    AssignedTo = pAssignedTo
End Property
Property Get FixedVersion() As String
    FixedVersion = pFixedVersion
End Property
Property Get ParentId() As Long
    ParentId = pParentId
End Property
Property Get Subject() As String
    Subject = pSubject
End Property
Property Get Description() As String
    Description = pDescription
End Property
Property Get StartDate() As Date
    StartDate = pStartDate
End Property
Property Get DueDate() As Date
    DueDate = pDueDate
End Property
Property Get DoneRatio() As Long
    DoneRatio = pDoneRatio
End Property
Property Get CustomFields() As Dictionary
    Set CustomFields = pCustomFields
End Property
Property Get CreatedOn() As Date
    CreatedOn = pCreatedOn
End Property
Property Get UpdatedOn() As Date
    UpdatedOn = pUpdatedOn
End Property

Private Sub Class_Initialize()
    ' 初期化処理 (必要に応じて追加)
    Set pData = New Dictionary
    pId = 0
    pProject = ""
    pTracker = ""
    pStatus = ""
    pPriority = ""
    pAuthor = ""
    pAssignedTo = ""
    pFixedVersion = ""
    pParentId = 0
    pSubject = ""
    pDescription = ""
    pStartDate = 0
    pDueDate = 0
    pDoneRatio = 0
    Set pCustomFields = Nothing
    pCreatedOn = 0
    pUpdatedOn = 0
End Sub
Private Sub Class_Terminate()
    ' 終了処理 (必要に応じて追加)
    Set pData = Nothing
    Set pCustomFields = Nothing
End Sub

'------------------------------------------------------------------------------
' Issue オブジェクトを初期化します。
' @param a_data Redmine API から取得したチケット情報の Dictionary
'------------------------------------------------------------------------------
Public Sub Init(ByVal a_data As Dictionary)
    On Error GoTo ErrorHandler

    Set pData = a_data

    If pData.Exists("id") Then pId = CLng(pData("id"))
    If pData.Exists("project") Then
        If TypeName(pData("project")) = "Dictionary" Then
            Dim projDict As Dictionary
            Set projDict = pData("project")
            If projDict.Exists("name") Then pProject = CStr(projDict("name"))
        End If
    End If
    If pData.Exists("tracker") Then
        If TypeName(pData("tracker")) = "Dictionary" Then
            Dim trackDict As Dictionary
            Set trackDict = pData("tracker")
            If trackDict.Exists("name") Then pTracker = CStr(trackDict("name"))
        End If
    End If
    If pData.Exists("status") Then
        If TypeName(pData("status")) = "Dictionary" Then
            Dim statusDict As Dictionary
            Set statusDict = pData("status")
            If statusDict.Exists("name") Then pStatus = CStr(statusDict("name"))
        End If
    End If
    If pData.Exists("priority") Then
        If TypeName(pData("priority")) = "Dictionary" Then
            Dim priorityDict As Dictionary
            Set priorityDict = pData("priority")
            If priorityDict.Exists("name") Then pPriority = CStr(priorityDict("name"))
        End If
    End If
    If pData.Exists("author") Then
        If TypeName(pData("author")) = "Dictionary" Then
            Dim authorDict As Dictionary
            Set authorDict = pData("author")
            If authorDict.Exists("name") Then pAuthor = CStr(authorDict("name"))
        End If
    End If
    If pData.Exists("assigned_to") Then
        If TypeName(pData("assigned_to")) = "Dictionary" Then
            Dim assignedDict As Dictionary
            Set assignedDict = pData("assigned_to")
            If assignedDict.Exists("name") Then pAssignedTo = CStr(assignedDict("name"))
        End If
    End If
    If pData.Exists("fixed_version") Then
        If TypeName(pData("fixed_version")) = "Dictionary" Then
            Dim versionDict As Dictionary
            Set versionDict = pData("fixed_version")
            If versionDict.Exists("name") Then pFixedVersion = CStr(versionDict("name"))
        End If
    End If
    If pData.Exists("parent") Then
        If TypeName(pData("parent")) = "Dictionary" Then
            Dim parentDict As Dictionary
            Set parentDict = pData("parent")
            If parentDict.Exists("id") Then pParentId = CLng(parentDict("id"))
        End If
    End If
    If pData.Exists("subject") Then pSubject = CStr(pData("subject"))
    If pData.Exists("description") Then pDescription = CStr(pData("description"))
    If pData.Exists("start_date") Then
        If IsDate(pData("start_date")) Then pStartDate = CDate(pData("start_date"))
    End If
    If pData.Exists("due_date") Then
        If IsDate(pData("due_date")) Then pDueDate = CDate(pData("due_date"))
    End If
    If pData.Exists("done_ratio") Then pDoneRatio = CLng(pData("done_ratio"))
    If pData.Exists("custom_fields") Then
        If TypeName(pData("custom_fields")) = "Collection" Then
            Dim cfCollection As Collection
            Set cfCollection = pData("custom_fields")
            Dim cfDict As Dictionary
            Set pCustomFields = New Dictionary
            Dim item As Variant
            For Each item In cfCollection
                If TypeName(item) = "Dictionary" Then
                    Set cfDict = item
                    If cfDict.Exists("name") And cfDict.Exists("value") Then
                        pCustomFields(cfDict("name")) = cfDict("value")
                    End If
                End If
            Next item
        End If
    End If
    If pData.Exists("created_on") Then
        If IsDate(pData("created_on")) Then pCreatedOn = CDate(pData("created_on"))
    End If
    If pData.Exists("updated_on") Then
        If IsDate(pData("updated_on")) Then pUpdatedOn = CDate(pData("updated_on"))
    End If

    Exit Sub
ErrorHandler:
    Err.Description = "Issue.Init(" & Erl & "): " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Sub

'------------------------------------------------------------------------------
' Issue オブジェクトの内容を文字列として取得します。
' @return Issue オブジェクトの内容を表す文字列
'------------------------------------------------------------------------------
Public Function ToString() As String
    On Error GoTo ErrorHandler

    Dim result As String
    result = "Issue {"
    result = result & "Id: " & pId & ", "
    result = result & "Project: " & pProject & ", "
    result = result & "Tracker: " & pTracker & ", "
    result = result & "Status: " & pStatus & ", "
    result = result & "Priority: " & pPriority & ", "
    result = result & "Author: " & pAuthor & ", "
    result = result & "AssignedTo: " & pAssignedTo & ", "
    result = result & "FixedVersion: " & pFixedVersion & ", "
    result = result & "ParentId: " & pParentId & ", "
    result = result & "Subject: " & pSubject & ", "
    result = result & "Description: " & Left(pDescription, 30) & IIf(Len(pDescription) > 30, "...", "") & ", "
    result = result & "StartDate: " & IIf(pStartDate <> 0, Format(pStartDate, "yyyy-mm-dd"), "N/A") & ", "
    result = result & "DueDate: " & IIf(pDueDate <> 0, Format(pDueDate, "yyyy-mm-dd"), "N/A") & ", "
    result = result & "DoneRatio: " & pDoneRatio & "%, "
    If Not pCustomFields Is Nothing Then
        Dim key As Variant
        Dim cfParts As Collection
        Set cfParts = New Collection
        For Each key In pCustomFields.Keys
            cfParts.Add key & ": " & CStr(pCustomFields(key))
        Next key
        result = result & "CustomFields: {" & Join(CollectionToArray(cfParts), ", ") & "}, "
    Else
        result = result & "CustomFields: {}, "
    End If
    result = result & "CreatedOn: " & IIf(pCreatedOn <> 0, Format(pCreatedOn, "yyyy-mm-dd HH:MM:SS"), "N/A") & ", "
    result = result & "UpdatedOn: " & IIf(pUpdatedOn <> 0, Format(pUpdatedOn, "yyyy-mm-dd HH:MM:SS"), "N/A")
    result = result & "}"

    ToString = result
    Exit Function
ErrorHandler:
    Err.Description = "Issue.ToString(" & Erl & "): " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Function

Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    CollectionToArray = arr
End Function

'------------------------------------------------------------------------------
' Issue オブジェクトの内容をDictionaryとして取得します。
' @return Issue オブジェクトの内容を表す Dictionary
'------------------------------------------------------------------------------
Public Function ToDictionary() As Dictionary
    On Error GoTo ErrorHandler

    Dim dict As Dictionary
    Set dict = New Dictionary

    dict("Id") = pId
    dict("Project") = pProject
    dict("Tracker") = pTracker
    dict("Status") = pStatus
    dict("Priority") = pPriority
    dict("Author") = pAuthor
    dict("AssignedTo") = pAssignedTo
    dict("FixedVersion") = pFixedVersion
    dict("ParentId") = pParentId
    dict("Subject") = pSubject
    dict("Description") = pDescription
    If pStartDate <> 0 Then
        dict("StartDate") = Format(pStartDate, "yyyy-mm-dd")
    Else
        dict("StartDate") = Null
    End If
    If pDueDate <> 0 Then
        dict("DueDate") = Format(pDueDate, "yyyy-mm-dd")
    Else
        dict("DueDate") = Null
    End If
    dict("DoneRatio") = pDoneRatio
    If Not pCustomFields Is Nothing Then
        Dim cfDict As Dictionary
        Set cfDict = New Dictionary
        Dim key As Variant
        For Each key In pCustomFields.Keys
            cfDict(key) = pCustomFields(key)
        Next key
        Set dict("CustomFields") = cfDict
    Else
        Set dict("CustomFields") = New Dictionary
    End If
    If pCreatedOn <> 0 Then
        dict("CreatedOn") = Format(pCreatedOn, "yyyy-mm-dd HH:MM:SS")
    Else
        dict("CreatedOn") = Null
    End If
    If pUpdatedOn <> 0 Then
        dict("UpdatedOn") = Format(pUpdatedOn, "yyyy-mm-dd HH:MM:SS")
    Else
        dict("UpdatedOn") = Null
    End If

    Set ToDictionary = dict
    Exit Function
ErrorHandler:
    Err.Description = "Issue.ToDictionary(" & Erl & "): " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Function
