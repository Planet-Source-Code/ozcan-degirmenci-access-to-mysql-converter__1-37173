Attribute VB_Name = "mdlMySQL"
Public DBNameMySQL As String
Public ServerMySQL As String
Public UserNameMySQL As String
Public PasswordMySQL As String
Public DB2Tables() As String
Public connMySQL As ADODB.Connection
Dim rs As ADODB.Recordset
Public TypeMySQL() As String

Public Function ConnectDBMySQL() As Boolean
  Dim sTablename As String
  On Error GoTo ErrorTrap
  DBNameMySQL = Trim(frmMAin.txtDBNameMySQL.Text)
  ServerNameMySQL = Trim(frmMAin.txtServerMySQL.Text)
  UserNameMySQL = Trim(frmMAin.txtUserNameMySQL.Text)
  PasswordMySQL = Trim(frmMAin.txtPasswordMySQL.Text)

  Set connMySQL = New ADODB.Connection
    connMySQL.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & _
    ServerNameMySQL & "; DATABASE=" & DBNameMySQL & "; UID=" & UserNameMySQL & _
    ";PWD=" & PasswordMySQL
    connMySQL.CursorLocation = adUseClient
  connMySQL.Open

        Set rs = connMySQL.Execute("SHOW TABLES;")
        ReDim Preserve DB2Tables(0)
        
       For i = 1 To rs.RecordCount
            sTablename = rs(0)
                ReDim Preserve DB2Tables(UBound(DB2Tables) + 1)
                DB2Tables(UBound(DB2Tables)) = sTablename
            rs.MoveNext
       Next
        
        rs.Close

        DoEvents
        ConnectDBMySQL = True
        
    Exit Function
ErrorTrap:
    If Not connMySQL Is Nothing Then
        If connMySQL.State = adStateOpen Then
            connMySQL.Close
        End If
        Set connMySQL = Nothing
    End If
    ConnectDBMySQL = Flase
    MsgBox "An error occured during connect to db. Error no: " & Err.Number & " err.desc = " & Err.Description
End Function

Public Sub MakeFieldListMySQL(FileName)
On Error GoTo ErrorTrap
    If frmMAin.cmbMySQLTables & "" = "" Then Exit Sub

    
    Dim rsfield As ADODB.Field
    Screen.MousePointer = vbHourglass
    

    Dim i As Integer
        For i = frmMAin.lstMySQLTable.ListCount - 1 To 0 Step -1
            frmMAin.lstMySQLTable.RemoveItem i
        Next i

    Set rs = New ADODB.Recordset
    rs.ActiveConnection = connMySQL

        rs.Open "SELECT * FROM " & frmMAin.cmbMySQLTables & ";"
        
    ReDim Preserve TypeMySQL(0)
    With rs
        For Each rsfield In .Fields
                frmMAin.lstMySQLTable.AddItem rsfield.Name
                'MsgBox rsfield.Name & " --- " & rsfield.Type
                ReDim Preserve TypeMySQL(UBound(TypeMySQL) + 1)
                TypeMySQL(UBound(TypeMySQL)) = rsfield.Type 'TypeName(rsField.Type)
        Next
    End With
    Screen.MousePointer = vbArrow
    rs.Close
    Set rs = Nothing
    Exit Sub
ErrorTrap:
    rs.Close: Set rs = Nothing
    MsgBox "During the process an error occured. err.nu : " & Err.Number & " err.des : " & Err.Description
End Sub

Function TypeNameMySQL(FieldType As Integer) As String

Select Case FieldType
    Case 2
        TypeNameMySQL = "SmallInt"
    Case 3
        TypeNameMySQL = "Int"
    Case 4
        TypeNameMySQL = "Float"
    Case 5
        TypeNameMySQL = "Double"
    Case 201
        TypeNameMySQL = "VarChar(255)"
    Case 135
        TypeNameMySQL = "Date/Time"
    Case 129
        TypeNameMySQL = "Char(1) binary"
    Case 133
        TypeNameMySQL = "Date"
    Case 134
        TypeNameMySQL = "Time"
    Case 202
        TypeNameMySQL = "Text"
    Case 203
        TypeNameMySQL = "Text"
    Case 200
        TypeNameMySQL = "Text"
    Case Else
        TypeNameMySQL = "Not Defined"
End Select

End Function
