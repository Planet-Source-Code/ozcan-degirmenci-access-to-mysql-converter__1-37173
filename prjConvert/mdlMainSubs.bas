Attribute VB_Name = "mdlMainSubs"


Public Sub Main()
    frmmain.Show
End Sub

Public Function NonNull(g) As String
Dim strTemp As String
    If IsNull(g) Then
        strTemp = ""
    Else
        strTemp = CStr(g)
    End If
    NonNull = strTemp
End Function

Public Sub EndProgram()
    CloseAccesConnection
    CloseMySQLConnection
End Sub

Public Sub CloseAccesConnection()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

Public Sub CloseMySQLConnection()
    If Not connMySQL Is Nothing Then
        If connMySQL.State = adStateOpen Then
            connMySQL.Close
        End If
        Set connMySQL = Nothing
    End If
End Sub

