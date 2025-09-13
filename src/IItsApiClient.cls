VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IItsApiClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' IItsApiClient インターフェース
'------------------------------------------------------------------------------
' ITsApiClient を実装するクラスが準拠すべきインターフェースを定義します。
'==============================================================================
Public Function getIssueByQuery(ByVal a_queryId As Long) As Dictionary
    ' インターフェースのメソッドは実装しません。
    Err.Raise 5 ' 5: Invalid procedure call or argument
End Function

Public Function getIssuesByProjectAndQuery(ByVal a_projectId As String, ByVal a_queryId As Long) As Dictionary
    ' インターフェースのメソッドは実装しません。
    Err.Raise 5 ' 5: Invalid procedure call or argument
End Function
