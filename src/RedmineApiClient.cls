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
Public Sub Init(ParamArray args() As Variant)
    On Error GoTo ErrorHandler

    If xInit Then Err.Raise 5 ' 5: Invalid procedure call or argument
    xLength = UBound(args) + 1

    If Not xLength = 2 Then Err.Raise 5
    xInit = True

    baseUrl = args(0)
    apiKey = args(1)

    Exit Sub
ErrorHandler:
    Err.Description = "RedmineClient.Init > " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Sub

'------------------------------------------------------------------------------
' 指定したクエリ ID に基づいてチケット情報を取得します。
'
' @param a_queryId 取得したいチケットのクエリ ID
' @return チケット情報を含むオブジェクト
'------------------------------------------------------------------------------
Public Function getIssueByQuery( _
    ByVal a_queryId As Long) As Object
    On Error GoTo ErrorHandler

    Debug.Print "RedmineApiClient.getIssueByQuery: Starting request for query ID: " & a_queryId

    Dim allIssues As Collection
    Set allIssues = New Collection

    Dim offset As Long
    offset = 0

    Dim limit As Long
    limit = 100 ' 初期値 (最初のレスポンスから取得)

    Do
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.SetProxy 2, "127.0.0.1:8888"

        ' URLを構築 (offset を追加)
        Dim url As String
        url = baseUrl & "issues.json?query_id=" & a_queryId & "&offset=" & offset & "&limit=" & limit
        Debug.Print "RedmineApiClient.getIssueByQuery: Request URL: " & url

        ' リクエスト設定
        http.Open "GET", url, False
        http.setRequestHeader "X-Redmine-API-Key", apiKey
        Debug.Print "RedmineApiClient.getIssueByQuery: Sending HTTP request..."
        http.Send

        ' レスポンスチェック
        Debug.Print "RedmineApiClient.getIssueByQuery: Response status: " & http.Status
        If http.Status = 200 Then
            Dim json As String
            json = http.responseText
            Debug.Print "RedmineApiClient.getIssueByQuery: Parsing JSON response..."
            ' JsonConverter.bas を使用してJSONを解析
            Dim response As Object
            Set response = JsonConverter.ParseJson(json)
            Debug.Print "RedmineApiClient.getIssueByQuery: JSON parsing completed successfully."

            ' 最初のレスポンスから limit を取得
            If offset = 0 Then
                limit = response("limit")
                Debug.Print "RedmineApiClient.getIssueByQuery: Retrieved limit from response: " & limit
            End If

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
            Debug.Print "RedmineApiClient.getIssueByQuery: Retrieved " & issues.Count & " issues. Total so far: " & allIssues.Count & " / " & totalCount
            offset = offset + limit

            If offset >= totalCount Then Exit Do
        Else
            Debug.Print "RedmineApiClient.getIssueByQuery: HTTP Error occurred: " & http.Status & " - " & http.statusText
            Err.Raise vbObjectError + 1, "getIssueByQuery", "HTTP Error: " & http.Status & " - " & http.statusText
        End If
    Loop

    ' すべての issues を含むオブジェクトを返す
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "issues", allIssues
    result.Add "total_count", allIssues.Count
    Set getIssueByQuery = result

    Exit Function
ErrorHandler:
    Err.Description = "RedmineClient.getIssueByQuery > " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Function




