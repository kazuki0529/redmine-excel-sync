VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineApiClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IItsApiClient

'==============================================================================
' RedmineApiClient クラス
'------------------------------------------------------------------------------
' Redmine API とやり取りするためのクラス。
' このクラスを使用して Redmine のチケット情報を取得できます。
'
' 使用例:
'   Dim client As RedmineApiClient
'   Set client = RedmineApiClient.of("https://redmine.example.com", "your-api-key")
'   Dim issues As Object
'   Set issues = client.getIssueByQuery(123)
'==============================================================================

Private baseUrl As String
Private apiKey As String
Private xInit As Boolean
Private xLength As Long

'------------------------------------------------------------------------------
' クラスの初期化時に呼び出されます。
' 直接インスタンス化せず、Ariawase の Init 関数を使用してください。
'
' @private
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub

'------------------------------------------------------------------------------
' クラスの終了時に呼び出されます。
'
' @private
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub

'------------------------------------------------------------------------------
' RedmineApiClient のインスタンスを作成します。
'
' @param a_baseUrl Redmine サーバーの URL (例: "https://redmine.example.com")
' @param a_apiKey Redmine API アクセス用の API キー
'------------------------------------------------------------------------------
Public Sub Init(ParamArray a_args() As Variant)
    On Error GoTo ErrorHandler

    If xInit Then Err.Raise 5 ' 5: Invalid procedure call or argument
    xLength = UBound(a_args) + 1

    If Not xLength = 2 Then Err.Raise 5
    xInit = True

    baseUrl = a_args(0)
    apiKey = a_args(1)

    Exit Sub
ErrorHandler:
    Err.Description = "RedmineClient.Init(" & Erl & ")" & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Sub

'------------------------------------------------------------------------------
' 指定したクエリ ID に基づいてチケット情報を取得します。
'
' @param a_queryId 取得したいチケットのクエリ ID
' @return チケット情報を含むオブジェクト
'------------------------------------------------------------------------------
Public Function IItsApiClient_getIssueByQuery( _
    ByVal a_queryId As Long) As Dictionary
    On Error GoTo ErrorHandler

    Debug.Print "RedmineApiClient.IItsApiClient_getIssueByQuery: Starting request for query ID: " & a_queryId

    Dim params As New ArrayEx
    Dim queryId As Tuple
    Set queryId = Core.Init(New Tuple, "query_id", a_queryId)
    params.AddObj queryId

    Set IItsApiClient_getIssueByQuery = fetchAllIssues(baseUrl & "issues.json", params)

    Exit Function
ErrorHandler:
    Err.Description = "RedmineClient.IItsApiClient_getIssueByQuery(" & Erl & "): " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Function

'------------------------------------------------------------------------------
' 指定したプロジェクト ID とクエリ ID に基づいてチケット情報を取得します。
'
' @param a_projectId プロジェクト ID
' @param a_queryId クエリ ID
' @return チケット情報を含むオブジェクト
'------------------------------------------------------------------------------
Public Function IItsApiClient_getIssuesByProjectAndQuery( _
    ByVal a_projectId As String, _
    ByVal a_queryId As Long) As Dictionary
    On Error GoTo ErrorHandler

    Debug.Print "RedmineApiClient.IItsApiClient_getIssuesByProjectAndQuery: Starting request for project ID: " & a_projectId & ", query ID: " & a_queryId

    Dim params As New ArrayEx
    Dim queryId As Tuple
    Set queryId = Core.Init(New Tuple, "query_id", a_queryId)
    params.AddObj queryId

    Set IItsApiClient_getIssuesByProjectAndQuery = fetchAllIssues(baseUrl & "projects/" & a_projectId & "/issues.json", params)

    Exit Function
ErrorHandler:
    Err.Description = "RedmineClient.IItsApiClient_getIssuesByProjectAndQuery(" & Erl & "): " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Function

'------------------------------------------------------------------------------
' 指定したベースクエリで全チケット情報を取得します (ページネーション対応)。
'
' @param a_baseQuery Tuple の Array (Collection) でクエリパラメータを指定
' @return チケット情報を含むオブジェクト
' @private
'------------------------------------------------------------------------------
Private Function fetchAllIssues( _
    ByVal a_baseUrl As String, _
    ByVal a_params As ArrayEx) As Dictionary
    On Error GoTo ErrorHandler

    ' Tuple の Array からクエリ文字列を生成
    Dim baseQuery As String
    baseQuery = ""
    Dim Tuple As Variant
    For Each Tuple In a_params.ToArray()
        If baseQuery <> "" Then baseQuery = baseQuery & "&"
        baseQuery = baseQuery & Tuple.Item1 & "=" & Tuple.Item2
    Next Tuple

    Debug.Print "RedmineApiClient.fetchAllIssues: Starting request for base query: " & baseQuery

    Dim allIssues As Collection
    Set allIssues = New Collection

    Dim offset As Long
    offset = 0

    Dim limit As Long
    limit = 100 ' 初期値 (後でレスポンスから取得)

    Do
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.SetProxy 2, "127.0.0.1:8888"

        ' URLを構築 (offset と limit を追加)
        Dim url As String
        url = a_baseUrl & "?" & baseQuery & "&offset=" & offset & "&limit=" & limit & "&include=[%27relations%27]"
        Debug.Print "RedmineApiClient.fetchAllIssues: Request URL: " & url

        ' リクエスト設定
        http.Open "GET", url, False
        http.setRequestHeader "X-Redmine-API-Key", apiKey
        Debug.Print "RedmineApiClient.fetchAllIssues: Sending HTTP request..."
        http.Send

        ' レスポンスチェック
        Debug.Print "RedmineApiClient.fetchAllIssues: Response status: " & http.Status
        If http.Status = 200 Then
            Dim json As String
            json = http.responseText
            Debug.Print "RedmineApiClient.fetchAllIssues: Parsing JSON response..."
            ' JsonConverter.bas を使用してJSONを解析
            Dim response As Object
            Set response = JsonConverter.ParseJson(json)
            Debug.Print "RedmineApiClient.fetchAllIssues: JSON parsing completed successfully."

            ' レスポンスから limit を取得
            limit = response("limit")
            Debug.Print "RedmineApiClient.fetchAllIssues: Retrieved limit from response: " & limit

            ' issues を収集
            Dim issues As Collection
            Set issues = response("issues")
            Dim issue As Variant
            For Each issue In issues
                allIssues.Add issue
            Next issue

            ' total_count と offset をチェック
            Dim totalCount As Long
            totalCount = response("total_count")
            Debug.Print "RedmineApiClient.fetchAllIssues: Retrieved " & issues.Count & " issues. Total so far: " & allIssues.Count & " / " & totalCount
            offset = offset + limit

            If offset >= totalCount Then Exit Do
        Else
            Debug.Print "RedmineApiClient.fetchAllIssues: HTTP Error occurred: " & http.Status & " - " & http.statusText
            Err.Raise vbObjectError + 1, "fetchAllIssues", "HTTP Error: " & http.Status & " - " & http.statusText
        End If
    Loop

    ' すべての issues を含むオブジェクトを返す
    Dim result As Dictionary
    Set result = New Dictionary
    result.Add "issues", allIssues
    result.Add "total_count", allIssues.Count
    Set fetchAllIssues = result

    Exit Function
ErrorHandler:
    Err.Description = "RedmineClient.fetchAllIssues(" & Erl & "): " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Function
