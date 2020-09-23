Attribute VB_Name = "mdlAcces"
Public conn As ADODB.Connection
Public rs1 As ADODB.Recordset
Public DB1Tables() As String
Public TypeAcces() As String
Public TypeAccesNullable() As String

Public Function GetDBTables(FileName) As Boolean
    
    Dim sTablename As String
    On Error GoTo ErrorTrap

    Set conn = New ADODB.Connection
    With conn
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";"
        .CursorLocation = adUseClient
        .Open

  
        Set rs1 = .OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
        
        ReDim Preserve DB1Tables(0)
        
        Do While Not rs1.EOF
            sTablename = rs1!TABLE_NAME

                ReDim Preserve DB1Tables(UBound(DB1Tables) + 1)
                DB1Tables(UBound(DB1Tables)) = sTablename
            rs1.MoveNext
        Loop
        
        rs1.Close
        DoEvents
    End With
    GetDBTables = True
    
exitGetDBTables:
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorTrap:
    GetDBTables = False
    Resume exitGetDBTables
    
End Function

Public Sub MakeFieldList(FileName)
Dim ok As Boolean
Dim FieldName As String
On Error GoTo ErrorTrap
    If frmMAin.cmbAccesTables & "" = "" Then Exit Sub

If Left(frmMAin.cmbAccesTables, 4) <> "MSys" Then
    Dim rsfield As ADODB.Field
    Screen.MousePointer = vbHourglass
    

    Dim i As Integer
        For i = frmMAin.lstAccesTable.ListCount - 1 To 0 Step -1
            frmMAin.lstAccesTable.RemoveItem i
        Next i

    Set rs1 = New ADODB.Recordset
    rs1.ActiveConnection = conn

        rs1.Open "SELECT * FROM [" & frmMAin.cmbAccesTables & "]"
    ReDim Preserve TypeAcces(0)
     ReDim Preserve TypeAccesNullable(0)
    With rs1
        For Each rsfield In .Fields
        FieldName = rsfield.Name
                frmMAin.lstAccesTable.AddItem rsfield.Name
                
                ReDim Preserve TypeAcces(UBound(TypeAcces) + 1)
                ReDim Preserve TypeAccesNullable(UBound(TypeAccesNullable) + 1)
                
                If rsfield.Type = 7 Then
                    If rs1.EOF Then
                        TypeAcces(UBound(TypeAcces)) = rsfield.Type
                    Else
                    ok = False
                    rs1.MoveFirst
                        For i = 1 To rs1.RecordCount
                            If rs1.EOF Then Exit For
                                If IsNull(rsfield.Value) Then
                                    rs1.MoveNext
                                Else
                            
                                    If FormatDateTime(rs1(FieldName).Value, vbLongTime) = rs1(FieldName).Value Then
                                    TypeAcces(UBound(TypeAcces)) = 123
                                    ok = True
                                    Exit For
                                    ElseIf FormatDateTime(rs1(FieldName).Value, vbGeneralDate) = rs1(FieldName).Value Then
                                    TypeAcces(UBound(TypeAcces)) = 122
                                    ok = True
                                    Exit For
                                    ElseIf Format(rs1(FieldName).Value, vbShortDate) = rs1(FieldName).Value Then
                                    TypeAcces(UBound(TypeAcces)) = 121
                                    ok = True
                                    Exit For
                                    ElseIf FormatDateTime(rs1(FieldName).Value, vbLongDate) = rs1(FieldName).Value Then
                                    TypeAcces(UBound(TypeAcces)) = 124
                                    ok = True
                                    Exit For
                                    ElseIf FormatDateTime(rs1(FieldName).Value, vbShortTime) = rs1(FieldName).Value Then
                                    TypeAcces(UBound(TypeAcces)) = 125
                                    ok = True
                                    Exit For
                                    Else
                                    TypeAcces(UBound(TypeAcces)) = rsfield.Type
                                    ok = True
                                    Exit For
                                    End If

                                End If
                        Next
                    End If
                                If Not ok Then
                                    TypeAcces(UBound(TypeAcces)) = rsfield.Type
                                End If
                Else
                    TypeAcces(UBound(TypeAcces)) = rsfield.Type
                End If
                      
        Next
    End With
    Screen.MousePointer = vbArrow
    rs1.Close
    Set rs1 = Nothing
ErrorTrap:
   If Not rs1 Is Nothing Then
        If rs1.State = adOpenState Then
            rs1.Close
        End If
        Set rs1 = Nothing
   End If
End If
End Sub

Function TypeName(FieldType As Integer) As String

Select Case FieldType
    Case 17
        TypeName = "Byte"
    Case 2
        TypeName = "Integer"
    Case 3
        TypeName = "Long Integer"
    Case 4
        TypeName = "Single"
    Case 5
        TypeName = "Double"
    Case 72
        TypeName = "Replication ID"
    Case 131
        TypeName = "Decimal"
    Case 202
        TypeName = "Text"
    Case 203
        TypeName = "Memo"
    Case 7
        TypeName = "Date/Time"
    Case 6
        TypeName = "Currency"
    Case 11
        TypeName = "Yes/No"
    Case 205
        TypeName = "OLE Object"
    Case 121
        TypeName = "Short Date"
    Case 122
        TypeName = "General Date"
    Case 123
        TypeName = "Long Time"
    Case 124
        TypeName = "Long Date"
    Case 125
        TypeName = "Short Time"
    Case Else
        TypeName = "Type Is Not Defined"
End Select

End Function

