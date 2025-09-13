VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineIssueRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IIssueRepository

'==============================================================================
' RedmineIssueRepository クラス
'------------------------------------------------------------------------------
' RedmineIssueRepository クラスは、Redmine API を使用して課題情報を取得するための実装を提供します。
'==============================================================================
Private apiClient As IItsApiClient

'------------------------------------------------------------------------------
' クラスの初期化時に呼び出されます。
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub

'------------------------------------------------------------------------------
' クラスの終了時に呼び出されます。
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set apiClient = Nothing
End Sub

'------------------------------------------------------------------------------
' RedmineIssueRepository のインスタンスを初期化します。
'
' @param a_apiClient IItsApiClient インターフェースを実装するオブジェクト
'------------------------------------------------------------------------------
Public Sub Init(ByRef a_apiClient As IItsApiClient)
    Set apiClient = a_apiClient

    Exit Sub
ErrorHandler:
    Err.Description = "RedmineIssueRepository.Init(" & Erl & "): " & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Sub

'------------------------------------------------------------------------------
' IIssueRepository インターフェースの getIssuesByProjectAndQuery メソッドの実装。
' 指定されたプロジェクトIDとクエリIDに基づいて課題情報を取得します。
'
' @param a_projectId プロジェクトID
' @param a_queryId クエリID
' @return 課題情報の Dictionary オブジェクト
'------------------------------------------------------------------------------
Public Function IIssueRepository_getIssuesByProjectAndQuery(ByVal a_projectId As String, ByVal a_queryId As Long) As Collection
    On Error GoTo ErrorHandler

    If apiClient Is Nothing Then
        Err.Raise 5 ' 5: Invalid procedure call or argument
    End If

    Dim issues As Dictionary
    Set issues = apiClient.getIssuesByProjectAndQuery(a_projectId, a_queryId)

    ' issues("issues") を Issueクラスのインスタンスに変換する処理を実施
    Dim result As New Collection
    Dim item As Dictionary
    For Each item In issues("issues")
        Dim i As Issue
        Set i = Core.Init(New Issue, item)
        result.Add i
    Next item
    Set IIssueRepository_getIssuesByProjectAndQuery = result

    Exit Function
ErrorHandler:
    Err.Description = "RedmineIssueRepository.getIssuesByProjectAndQuery(" & Erl & ")" & Err.Description
    Debug.Print Err.Description
    Err.Raise Err.Number
End Function
