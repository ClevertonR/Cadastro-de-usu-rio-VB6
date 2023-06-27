Attribute VB_Name = "mdlGlobal"
Option Explicit

' Made By Michael Ciurescu (CVMichael from vbforums.com)
Public NomeUsuario As String
Public DBConn As ADODB.Connection
Public LogInUserID As Long, LogInUserName As String

Public Function LoadDatabase(ByVal DatabaseName As String, Optional ByVal UserID As String, Optional ByVal Password As String) As ADODB.Connection
    Dim conData As ADODB.Connection
    
    Set conData = New ADODB.Connection
    
    conData.Provider = "Microsoft.Jet.OLEDB.4.0"
    conData.ConnectionString = "Data Source = " & DatabaseName
    conData.CursorLocation = adUseClient
    conData.Open , UserID, Password
    
    Set LoadDatabase = conData
End Function

Public Function SelectNewID(Cn As ADODB.Connection, ByVal TableName As String, Optional ByVal IDFieldName As String = "ID") As Long
    Dim Request As String, RS As ADODB.Recordset
    Dim NewID As Long
    
    Request = "SELECT MAX(" & IDFieldName & ") FROM " & TableName
    Set RS = Cn.Execute(Request)
    
    If RS Is Nothing Then
        NewID = 1
    Else
        If RS.RecordCount = 0 Then
            NewID = 1
        Else
            RS.MoveFirst
            
            If IsNull(RS.Fields(0).value) Then
                NewID = 1
            Else
                NewID = CLng(RS.Fields(0).value) + 1
            End If
        End If
    End If
    
    SelectNewID = NewID
End Function

