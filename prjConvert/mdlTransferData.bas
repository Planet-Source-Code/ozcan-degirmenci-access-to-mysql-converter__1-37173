Attribute VB_Name = "mdlTransferData"
Option Explicit
Public TableMySQLName As String


Function CreateTable() As Boolean
On Error GoTo ErrorTrap
    Dim i As Integer
    Dim sql_str As String
    Dim rs As ADODB.Recordset
    Dim a As Integer

    
    'frmCopying.Show ' Durum Bildiriliyor ...
        
    TableMySQLName = MySQLName(frmMAin.cmbAccesTables)
        
    '    frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
     '       "Starting Create Table " & TableMySQLName & vbCrLf
    
    sql_str = "Create Table " & frmMAin.txtDBNameMySQL & "." & MySQLName(TableMySQLName) & " ("
    
      '  frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
       '     "Making Columns " & vbCrLf & "-----------------------" & vbCrLf
            
    For i = 0 To frmMAin.lstAccesTable.ListCount - 1
        sql_str = sql_str & MySQLName(frmMAin.lstAccesTable.List(i)) & " " & MySQLType(CInt(TypeAcces(i + 1))) '& " " & FindNullable(i + 1)
        If i = frmMAin.lstAccesTable.ListCount - 1 Then
            sql_str = sql_str
        Else
            sql_str = sql_str & ", "
        End If
        'frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
         '   "Create " & frmMAin.lstAccesTable.List(i) & " column with type -- > " & _
         '   MySQLType(CInt(TypeAcces(i + 1))) & vbCrLf
    Next
    sql_str = sql_str & ");"
    'Debug.Print sql_str
    Set rs = connMySQL.Execute(sql_str)
        frmMAin.cmbMySQLTables.AddItem TableMySQLName
    CreateTable = True
        'frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
         '   "Table " & TableMySQLName & " was created succesfully..." & vbCrLf
    Exit Function
ErrorTrap:
          '  frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
           ' "Becouse of an error creating " & TableMySQLName & " was cancelled." & vbCrLf
           MsgBox "During Created Table an error occured. Err.Nu : " & Err.Number & " err.desc : " & Err.Description
    CreateTable = False
End Function


Sub CopyTable()
Dim sql_delete As String
Dim sql_select As String
Dim sql_insert As String
Dim sql_insert_part2 As String
Dim rsfield As ADODB.Field
Dim rs As ADODB.Recordset
Dim rsinsert As ADODB.Recordset
Dim a As Long
Dim i As Long
Dim counter As Long
Dim k As String

On Error GoTo ErrorTrap
       ' frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
        '    "Starting Copying Rows from " & frmMAin.cmbAccesTables & " (Access) Table to " & _
         '   TableMySQLName & " (MySQL) Table" & vbCrLf & "------------------------" & vbCrLf
    sql_select = "Select * from [" & frmMAin.cmbAccesTables & "]"
    Set rs = conn.Execute(sql_select) ' Acces Tabloyabaðlandýk
    If Not rs.EOF Then
        With rs
        counter = 1
            
            Do
            
       ' frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
            "Prepare SQL string for " & counter & ". record" & vbCrLf & "------------" & vbCrLf

                a = 0
                sql_insert = "Insert into " & TableMySQLName '& " ("
                sql_insert_part2 = " values ('"
                    For Each rsfield In .Fields

                        
                      If IsNull(.Fields(a).Value) Then
                        k = ""
                        a = a + 1
                      Else
                        k = NonNull(.Fields(a).Value)
                        k = Replace(k, "'", "_")
                        
                        a = a + 1
                            If TypeAcces(a) = 121 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "yyyy-mm-dd")
                            ElseIf TypeAcces(a) = 122 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "yyyy-mm-dd hh:mm:ss")
                            ElseIf TypeAcces(a) = 123 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "hh:mm:ss AMPM")
                            ElseIf TypeAcces(a) = 124 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, vbGeneralDate)
                            ElseIf TypeAcces(a) = 125 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "hh:mm:ss")
                            End If
                        End If
                         
                        sql_insert_part2 = sql_insert_part2 & k
                        
                        If a = .Fields.Count Then
                            'sql_insert = sql_insert & ")"
                            sql_insert_part2 = sql_insert_part2 & "')"
                        Else
                            'sql_insert = sql_insert & ","
                            sql_insert_part2 = sql_insert_part2 & "','"
                        End If
                        
        '                frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
                            "   " & a & ".Column prepared." & vbCrLf
                    Next
                    
         '           frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
                        "SQL String Ok!" & vbCrLf & "Inserting..." & vbCrLf
                        
                  sql_insert = sql_insert & sql_insert_part2
                  Debug.Print sql_insert
                  Set rsinsert = connMySQL.Execute(sql_insert)
          '        frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
                        "Insert Ok!" & vbCrLf
                        
                rs.MoveNext
                counter = counter + 1
                frmMAin.Label11.Caption = counter & " record copied."
            Loop Until rs.EOF
           ' frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
                "Copying Table Finished." & vbCrLf & vbCrLf & "Total : " & counter & " record copied.& vbCrLf"
        End With
    End If
    If frmMAin.optDelete = True Then
        'frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
            "Starting Delete" & vbCrLf & "-------------" & vbCrLf
        sql_delete = "Delete * from " & frmMAin.cmbAccesTables
        Set rs = conn.Execute(sql_delete)
        'frmCopying.txtStatus.Text = frmCopying.txtStatus.Text & _
            "All Records were deleted" & vbCrLf & "-------------" & vbCrLf
    End If
Exit Sub
ErrorTrap:
    MsgBox "During operation there was a problem occured. Program will end the process.   err.no : " & _
        Err.Number & " err.desc : " & Err.Description
End Sub


Function MySQLType(iTypeAcces As Integer) As String
Dim TypeOfColumn As String
    Select Case iTypeAcces
        Case 17 ' Byte için burasý defined size = 1 byte
            TypeOfColumn = "CHAR(255)" 'EMÝN DEÐÝLÝM
        Case 2  ' Integer için burasý defined size = 2 byte
            TypeOfColumn = "SMALLINT"
        Case 3  ' Long int için bura defined size = 4 byte
            TypeOfColumn = "INT"
        Case 4  ' Single(fload) için bura defined size = 4 byte
            TypeOfColumn = "FLOAT"
        Case 5  ' double için burasý defined size = 8 byte
            TypeOfColumn = "DOUBLE"
        Case 72 ' replication id için burasý defined size = 16
            TypeOfColumn = "INT" 'valla bulamadým abi
        Case 131    ' decimal için bura defined size = 19
            TypeOfColumn = "DOUBLE"
        Case 202 ' Text için burasý valla 255 karektere kadar yolu var haberin ola
            TypeOfColumn = "VARCHAR(255)"
        Case 203    ' memo için bura 65 535 karektee kadar yolu var
            TypeOfColumn = "LONGTEXT"
        Case 7  ' Date /Time valla tüm çeþitleri için burasý var sdece
            TypeOfColumn = "DATETIME"
            'TypeOfColumn = MySQLDateTime(iTypeAcces)
        Case 6  ' currency burasý ve tek olyor tüm çeþitleri için eðer bulursam bir yolunu iyi olur
            TypeOfColumn = "VARCHAR(255)" 'EMÝN DEÐÝLÝM
        Case 11 ' yes/no, true/false, on/off
            TypeOfColumn = "CHAR(1) BINARY" 'EMIN DEGILIM
        Case 205    ' ole aman boþ ver o kadar önemli deðil
            TypeOfColumn = "LONGTEXT"
        Case 121
            TypeOfColumn = "DATE"
        Case 122
            TypeOfColumn = "DATETIME"
        Case 123
            TypeOfColumn = "CHAR(50)"
        Case 124
            TypeOfColumn = "DATETIME"
        Case 125
            TypeOfColumn = "CHAR(50)"
        Case Else
            TypeOfColumn = "LONGTEXT"
    End Select
MySQLType = TypeOfColumn
End Function


Function FindNullable(NumOfEl As Integer) As String
Dim strTemp
    If TypeAccesNullable(NumOfEl) Then
        strTemp = "NOT NULL"
    Else
        strTemp = "default NULL"
    End If
    MsgBox strTemp
    FindNullable = strTemp
End Function

Function MySQLName(tmp)
  tmp = Replace(tmp, " ", "_")
  tmp = Replace(tmp, "-", "_")
  tmp = Replace(tmp, "(", "_")
  If Not ReservedWords(tmp) Then
    tmp = tmp & "_c"
  End If
  MySQLName = Replace(tmp, ")", "_")
End Function

Function ReservedWords(strTemp) As Boolean 'Tablo column isimlerinde hata lmasýn
Dim a As Boolean
a = True
    Select Case LCase(strTemp)
        Case "int": a = False
        Case "update": a = False
        Case "null": a = False
        Case "not null": a = False
        Case "set": a = False
        Case "select": a = False
        Case "float": a = False
        Case "double": a = False
        Case "date": a = False
        Case "datetime": a = False
        Case "default": a = False
    End Select
    ReservedWords = a
End Function

Function CreateDB() As Boolean
    Dim NewDBMySQLName As String
    Dim rs As ADODB.Recordset
    Dim sql_str As String
    On Error GoTo ErrorTrap
    NewDBMySQLName = InputBox("Please Enter the name of the new Database", "New Database's Name")
    If Not Trim(NewDBMySQLName) = "" Then
        sql_str = "CREATE DATABASE " & NewDBMySQLName
        Set rs = connMySQL.Execute(sql_str)
                
        CreateDB = True
        frmMAin.cmdDisconnect_Click
        frmMAin.txtDBNameMySQL.Text = NewDBMySQLName
        frmMAin.cmdConnectDBmysql_Click
        Exit Function
    Else
        CreateDB = False
        Exit Function
    End If
ErrorTrap:
    CreateDB = False
    MsgBox "During the operation there was an error occured. err.no : " & Err.Number & " err.desc : " & Err.Description
End Function

Sub CopyTables()
Dim i As Integer
On Error GoTo ErrorTrap
    
    For i = 0 To frmMAin.cmbAccesTables.ListCount - 1
        frmMAin.cmbAccesTables.ListIndex = i
        If CreateTable Then
            CopyTable
        End If
    Next
    
    Exit Sub
ErrorTrap:
    MsgBox "Error Occured err.desc : " & Err.Description
End Sub

Function ControlTables() As Boolean
Dim i As Integer
Dim IsOk As Boolean
On Error GoTo ErrorTrap
IsOk = True
    If frmMAin.lstAccesTable.ListCount = frmMAin.lstMySQLTable.ListCount Then ' Eðer column sayýlarý ayný ise
        For i = 0 To frmMAin.lstAccesTable.ListCount - 1
            If Not ConvertType(CInt(TypeAcces(i + 1)), CInt(TypeMySQL(i + 1))) Then ' her elemanýn type ý kontrol ediliyor
                IsOk = False
            End If
        Next
    Else
        ControlTables = False
        Exit Function
    End If
    
    If IsOk Then
        ControlTables = True
        Exit Function
    Else
        ControlTables = False
        Exit Function
    End If
ErrorTrap:
    ControlTables = False
    MsgBox "During Control the tables an error occured. Err.No: " & Err.Number & " Err.Desc : " & Err.Description
End Function


Function ConvertType(i As Integer, a As Integer) As Boolean
Dim IsOk As Boolean
IsOk = False
    If a = 201 Then
        Select Case i
            Case 17: IsOk = True
            Case 202: IsOk = True
            Case 6: IsOk = True
            Case 205: IsOk = True
            Case 203: IsOk = True
            Case 123: IsOk = True
            Case 125: IsOk = True
            Case Else: IsOk = False
        End Select
    ElseIf a = 200 Then
        Select Case i
            Case 17: IsOk = True
            Case 202: IsOk = True
            Case 6: IsOk = True
            Case 205: IsOk = True
            Case 203: IsOk = True
            Case 123: IsOk = True
            Case 125: IsOk = True
            Case Else: IsOk = False
        End Select
    ElseIf a = 202 Then
        Select Case i
            Case 17: IsOk = True
            Case 202: IsOk = True
            Case 6: IsOk = True
            Case 205: IsOk = True
            Case 203: IsOk = True
            Case 123: IsOk = True
            Case 125: IsOk = True
            Case Else: IsOk = False
        End Select
    ElseIf a = 203 Then
        Select Case i
            Case 17: IsOk = True
            Case 202: IsOk = True
            Case 6: IsOk = True
            Case 205: IsOk = True
            Case 203: IsOk = True
            Case 123: IsOk = True
            Case 125: IsOk = True
            Case Else: IsOk = False
        End Select
    ElseIf a = 2 Then
        If i = 2 Then IsOk = True
    ElseIf a = 3 Then
        If i = 3 Then IsOk = True
    ElseIf a = 4 Then
        If i = 4 Then IsOk = True
    ElseIf a = 5 Then
        If i = 5 Then IsOk = True
    ElseIf a = 135 Then
        Select Case i
            Case 7: IsOk = True
            Case 122: IsOk = True
            Case 124: IsOk = True
            Case Else: IsOk = False
        End Select
    ElseIf a = 129 Then
        If i = 11 Then IsOk = True
    ElseIf a = 133 Then
        If i = 121 Then IsOk = True
    ElseIf a = 134 Then
        If i = 7 Then IsOk = True
        If i = 124 Then IsOk = True
    Else
        IsOk = False
    End If
ConvertType = IsOk
End Function

Function CopyRow() As Boolean
On Error GoTo ErrorTrap
Dim sql_str As String
Dim sql_insert_part2 As String
Dim k
Dim i As Long
Dim a As Long
Dim rsinsert As ADODB.Recordset
TableMySQLName = MySQLName(frmMAin.cmbAccesTables)
                a = 0
                sql_str = "Insert into " & TableMySQLName '& " ("
                sql_insert_part2 = " values ('"
                    For i = 1 To frmMAin.gridData.Columns.Count
                      If IsNull(frmMAin.gridData.Columns(a).Text) Then
                        k = ""
                        a = a + 1
                      Else
                        k = NonNull(frmMAin.gridData.Columns(a).Text)
                        k = Replace(k, "'", "_")
                        
                        a = a + 1
                            If TypeAcces(a) = 121 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "yyyy-mm-dd")
                            ElseIf TypeAcces(a) = 122 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "yyyy-mm-dd hh:mm:ss")
                            ElseIf TypeAcces(a) = 123 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "hh:mm:ss AMPM")
                            ElseIf TypeAcces(a) = 124 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, vbGeneralDate)
                            ElseIf TypeAcces(a) = 125 Then
                                k = Replace(k, "/", "-")
                                k = Format(k, "hh:mm:ss")
                            End If
                        End If
                         
                        sql_insert_part2 = sql_insert_part2 & k
                        
                        If a = frmMAin.gridData.Columns.Count Then
                            'sql_insert = sql_insert & ")"
                            sql_insert_part2 = sql_insert_part2 & "')"
                        Else
                            'sql_insert = sql_insert & ","
                            sql_insert_part2 = sql_insert_part2 & "','"
                        End If
                        Next
                  sql_str = sql_str & sql_insert_part2 & ";"
                  'Debug.Print sql_str
                  Set rsinsert = connMySQL.Execute(sql_str)

CopyRow = True
Exit Function
ErrorTrap:
    CopyRow = False
    MsgBox "During the copy the data's an error occured. err.nu :" & Err.Number & " err.desc : " & Err.Description, , "NOT COPY"
End Function
