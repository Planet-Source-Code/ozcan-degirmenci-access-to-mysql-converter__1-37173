VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   Caption         =   "Convert Access To MySQL"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "  Status  "
      Height          =   855
      Left            =   5400
      TabIndex        =   53
      Top             =   3120
      Width           =   1695
      Begin VB.Label Label11 
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc adoDataMySQL 
      Height          =   375
      Left            =   5520
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   4695
      Left            =   5400
      TabIndex        =   14
      Top             =   4080
      Width           =   1695
      Begin VB.CommandButton cmdCopyDB 
         Caption         =   "Copy Full DB"
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton optProtec 
         Caption         =   "Protec Data"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDelete 
         Caption         =   "Delete Data"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCopyRow 
         Caption         =   "Copy Row"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopyStructure 
         Caption         =   "Copy Structure"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopyTable 
         Caption         =   "Copy Table"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   1560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   1560
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   1560
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin MSComDlg.CommonDialog cmmDialog 
      Left            =   6840
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Ms Acces File"
      Filter          =   "*.mdb"
      InitDir         =   "c:\Windows\Desktop"
   End
   Begin VB.Frame Frame2 
      Caption         =   " MySQL Database "
      Height          =   8775
      Left            =   7200
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   2520
         TabIndex        =   45
         Top             =   2200
         Width           =   2535
         Begin VB.Label Label20 
            Caption         =   "Data Type:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   50
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label19 
            Height          =   255
            Left            =   1200
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Total Record:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   45
            TabIndex        =   49
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblRecordCountMySQL 
            Height          =   255
            Left            =   1200
            TabIndex        =   48
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Total Column:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   45
            TabIndex        =   47
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblColumnMySQL 
            Height          =   255
            Left            =   1200
            TabIndex        =   46
            Top             =   720
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid gridDataMySQL 
         Bindings        =   "frmMain.frx":030A
         Height          =   4455
         Left            =   120
         TabIndex        =   33
         Top             =   3360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   0   'False
         ForeColor       =   -2147483642
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtSQLMysql 
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   8160
         Width           =   4095
      End
      Begin VB.CommandButton cmdExecuteMySQL 
         Caption         =   "Execute"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4320
         TabIndex        =   35
         Top             =   8160
         Width           =   825
      End
      Begin VB.CheckBox editModeMySQL 
         Caption         =   "Edit Mode"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   7920
         Width           =   1335
      End
      Begin VB.ListBox lstMySQLTable 
         Enabled         =   0   'False
         Height          =   1035
         ItemData        =   "frmMain.frx":0325
         Left            =   120
         List            =   "frmMain.frx":0327
         TabIndex        =   32
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox cmbMySQLTables 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   30
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "DisConn.."
         Enabled         =   0   'False
         Height          =   735
         Left            =   4200
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdConnectDBmysql 
         Caption         =   "Connect"
         Height          =   735
         Left            =   4200
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtPasswordMySQL 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtUserNameMySQL 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtDBNameMySQL 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   "Lynn"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtServerMySQL 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Text            =   "localhost"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   5040
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label7 
         Caption         =   "Tables :"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Password :"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "User Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Database Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Server Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " MS Acces Database "
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdDisconnectAcces 
         Caption         =   "DisConnect"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3000
         TabIndex        =   5
         Top             =   650
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   2760
         TabIndex        =   37
         Top             =   1200
         Width           =   2415
         Begin VB.Label lblColumnAcces 
            Height          =   255
            Left            =   1200
            TabIndex        =   43
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Total Column:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   45
            TabIndex        =   42
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblRecordCountAcces 
            Height          =   255
            Left            =   1200
            TabIndex        =   41
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Total Record:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   45
            TabIndex        =   40
            Top             =   480
            Width           =   975
         End
         Begin VB.Label txtFieldTypeAcces 
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Data Type:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   50
            TabIndex        =   38
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox editModeAcces 
         Caption         =   "Edit Mode"
         Enabled         =   0   'False
         Height          =   230
         Left            =   120
         TabIndex        =   13
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton cmdExecuteAcces 
         Caption         =   "Execute"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4380
         TabIndex        =   11
         Top             =   8160
         Width           =   825
      End
      Begin VB.TextBox txtSQLAcces 
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   8160
         Width           =   4215
      End
      Begin VB.ListBox lstAccesTable 
         Enabled         =   0   'False
         Height          =   840
         ItemData        =   "frmMain.frx":0329
         Left            =   120
         List            =   "frmMain.frx":032B
         TabIndex        =   7
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cmbAccesTables 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   650
         Width           =   2175
      End
      Begin VB.CommandButton cmdShowDB 
         Caption         =   "Connect"
         Default         =   -1  'True
         Height          =   350
         Left            =   4320
         TabIndex        =   4
         Top             =   260
         Width           =   855
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   350
         Left            =   3000
         TabIndex        =   3
         Top             =   260
         Width           =   1215
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   300
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid gridData 
         Bindings        =   "frmMain.frx":032D
         Height          =   5295
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   0   'False
         ForeColor       =   -2147483642
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "DB :"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   120
         X2              =   5160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Fields:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tables:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc adoData 
      Height          =   330
      Left            =   4800
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoData"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMAin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileNameAcces As String
Public CopyDB As Boolean

Private Sub txtDBNameMySQL_Click()
    cmdConnectDBmysql.Default = True
End Sub

Private Sub txtDBNameMySQL_LostFocus()
    cmdConnectDBmysql.Default = False
End Sub

Private Sub txtServerMySQL_Click()
    cmdConnectDBmysql.Default = True
End Sub

Private Sub txtServerMySQL_LostFocus()
    cmdConnectDBmysql.Default = False
End Sub

Private Sub txtSQLAcces_Click()
    cmdExecuteAcces.Default = True
End Sub

Private Sub txtSQLAcces_LostFocus()
    cmdDisconnectAcces.Default = True
End Sub

Private Sub txtSQLmysql_Click()
    cmdExecuteMySQL.Default = True
End Sub

Private Sub txtSQLMysql_LostFocus()
    cmdDisconnect.Default = True
End Sub
Private Sub cmbAccesTables_click()
    If Not cmbAccesTables.Text = "Select Table >>" Or CopyDB = True Then
        MakeFieldList (frmMAin.txtFileName)
        LoadData (txtFileName)
    End If
End Sub
Private Sub cmbMySQLTables_click()
    If Not cmbMySQLTables.Text = "Select Table >>" Then
        MakeFieldListMySQL (frmMAin.cmbMySQLTables)
        LoadDataMySQL
    End If
End Sub

Private Sub cmdBrowse_Click()
Dim FileName
On Error GoTo ErrorTrap
    cmmDialog.ShowOpen
    FileName = cmmDialog.FileName
    FileNameAcces = FileName
    txtFileName.Text = FileName
    
    For i = cmbAccesTables.ListCount - 1 To 0 Step -1
        cmbAccesTable.RemoveItem i
    Next i

    If GetDBTables(FileName) = True Then
        MsgBox "Connecting to the Acces database succesfully..."
        MakeEnabledAcces (True)
            cmbAccesTables.Text = "Select Table >>"
        For i = 1 To UBound(DB1Tables)
            cmbAccesTables.AddItem DB1Tables(i)
        Next i
    End If

Exit Sub
If Err = 32755 Then
    Exit Sub
End If
ErrorTrap:
    MsgBox "An error occured during connect to acces db. error no : " & Err.Number & " err.desc : " & Err.Description
End Sub

Public Sub cmdConnectDBmysql_Click()
    If ConnectDBMySQL Then
        MsgBox "Connecting to the MySQL database succesfully..."
        MakeEnabled (True)
        cmbMySQLTables.Text = "Select Table >>"
        For i = 1 To UBound(DB2Tables)
            cmbMySQLTables.AddItem DB2Tables(i)
        Next i
    End If
    For i = lstMySQLTable.ListCount - 1 To 0 Step -1
        lstMySQLTable.RemoveItem i
    Next i
End Sub
Sub MakeEnabled(a As Boolean)
    txtServerMySQL.Enabled = Not a
    txtPasswordMySQL.Enabled = Not a
    txtUserNameMySQL.Enabled = Not a
    txtDBNameMySQL.Enabled = Not a
    cmdConnectDBmysql.Enabled = Not a
    Label3.Enabled = Not a
    Label4.Enabled = Not a
    Label5.Enabled = Not a
    Label6.Enabled = Not a
    Label7.Enabled = a
    cmdDisconnect.Enabled = a
    cmbMySQLTables.Enabled = a
    lstMySQLTable.Enabled = a
    gridDataMySQL.Enabled = a
    editModeMySQL.Enabled = a
    txtSQLMysql.Enabled = a
    cmdExecuteMySQL.Enabled = a
    Label18.Enabled = a
    Label20.Enabled = a
    Label16.Enabled = a
    Frame5.Enabled = a
    cmdDisconnect.Default = a
    cmdConnectDBmysql.Default = Not a
End Sub
Sub MakeEnabledAcces(a As Boolean)
    cmbAccesTables.Enabled = a
    lstAccesTable.Enabled = a
    gridData.Enabled = a
    txtSQLAcces.Enabled = a
    editModeAcces.Enabled = a
    cmdExecuteAcces.Enabled = a
    Frame3.Enabled = a
    Label8.Enabled = a
    Label9.Enabled = a
    Label10.Enabled = a
    Label1.Enabled = a
    Label2.Enabled = a
    txtFileName.Enabled = Not a
    cmdBrowse.Enabled = Not a
    cmdShowDB.Enabled = Not a
    cmdDisconnectAcces.Enabled = a
    Label12.Enabled = Not a
    cmdDisconnectAcces.Default = a
    cmdShowDB.Default = Not a
End Sub

Private Sub cmdCopyDB_Click()
    If Not connMySQL Is Nothing Then ' Eðer MysqL Servera baðlý isek
        If Not conn Is Nothing Then ' Ve Acces baðlantý varsa
            If optDelete = True Then
                If MsgBox("You select the DELETE DATA from Access table" & _
                        " during copying the datas. Are you sure you want to delete these" & _
                        "datas?", vbYesNo, "DELETE DATA CONFIRMATION") = vbYes Then
                        CopyTable
                Else
                    MsgBox "Copying DB was cancelled", , "CANCEL"
                End If
            Else
                If CreateDB Then  ' ve son olarak da bu tabloyo yaratbilmilþ isek devam et...
                    CopyDB = True
                            ' Buraya Tablolar için döngü kurmam gerekir dur hele bir bakalým lo
                    CopyTables  ' Tablolarý kopyalama fonksiyonum ve bunu nereye koysam heh buldum mdlTransfer e koyuyum
                    MsgBox "Copying Database was finished", , "FINISHED"
                    CopyDB = False
                End If
            End If
        Else
            MsgBox "There is no selected Access Database. Please Select a Access Database and connect it", , "No Selected Tables"
        End If
    Else
        MsgBox "There is no connection to the MysqL Sever. To Copy Table please first connect to the MySQL Sever and than try again", , "No Connection To The MySQL Sever"
    End If
End Sub


Private Sub cmdCopyRow_Click()
     If cmbAccesTables = "Select Table >>" Then
        MsgBox "There is no selected acces table", , "Select Table"
     Else
        If cmbMySQLTables = "Select Table >>" Then
            MsgBox "There is no selected MySQL Table", , "Select Table"
        Else
            If ControlTables Then
                If CopyRow Then
                    MsgBox "Selected Row Copied succesfully"
                    LoadDataMySQL
                End If
            Else
                MsgBox "The selected tables ar not in the same structure. This will cause error during process." & _
                    " Please select another table or please use" & _
                    " Copy Structure to create a new table with this structure.", , "STRUCTURE ERROR"
            End If
        End If
    End If
End Sub

Private Sub cmdCopyStructure_Click()
    If Not connMySQL Is Nothing Then ' Eðer MysqL Servera baðlý isek
        If lstAccesTable.ListCount <> 0 Then ' Ve bir tablo seçmiþ isek
            If Not CreateTable Then  ' ve son olarak da bu tabloyo yaratbilmilþ isek devam et...
                MsgBox "During created the Table thare was an error. The program will end this operation.", , "Error During Cretaing Table"
            Else
                MsgBox "Creating table was finished", , "FINISHED"
            End If
        Else
            MsgBox "There is no selected Table to copy. Please Select a table to copy", , "No Selected Tables"
        End If
    Else
        MsgBox "There is no connection to the MysqL Sever. To Copy Table please first connect to the MySQL Sever and than try again", , "No Connection To The MySQL Sever"
    End If
End Sub

Private Sub cmdCopyTable_Click()
    If Not connMySQL Is Nothing Then ' Eðer MysqL Servera baðlý isek
        If lstAccesTable.ListCount <> 0 Then ' Ve bir tablo seçmiþ isek
            If CreateTable Then  ' ve son olarak da bu tabloyo yaratbilmilþ isek devam et...
                If optDelete = True Then
                    If MsgBox("You select the DELETE DATA from Access table" & _
                            " during copying the datas. Are you sure you want to delete these" & _
                            "datas?", vbYesNo, "DELETE DATA CONFIRMATION") = vbYes Then
                            CopyTable
                            MsgBox "Copying table was finished", , "FINISHED"
                    Else
                        MsgBox "Copying Table was cancelled", , "CANCEL"
                    End If
                Else
                    CopyTable
                End If
            Else
                MsgBox "During created the Table thare was an error. The program will end this operation.", , "Error During Cretaing Table"
            End If
        Else
            MsgBox "There is no selected Table to copy. Please Select a table to copy", , "No Selected Tables"
        End If
    Else
        MsgBox "There is no connection to the MysqL Sever. To Copy Table please first connect to the MySQL Sever and than try again", , "No Connection To The MySQL Sever"
    End If
End Sub

Public Sub cmdDisconnect_Click()

    CloseMySQLConnection
    
    For i = lstMySQLTable.ListCount - 1 To 0 Step -1
        lstMySQLTable.RemoveItem (i)
    Next
    
    For i = cmbMySQLTables.ListCount - 1 To 0 Step -1
        cmbMySQLTables.RemoveItem (i)
    Next
    lblRecordCountMySQL.Caption = ""
    lblColumnMySQL.Caption = ""
    gridDataMySQL.ClearFields
    MakeEnabled (False)
End Sub

Private Sub cmdDisconnectAcces_Click()
    CloseAccesConnection
    For i = lstAccesTable.ListCount - 1 To 0 Step -1
        lstAccesTable.RemoveItem (i)
    Next
    
    For i = cmbAccesTables.ListCount - 1 To 0 Step -1
        cmbAccesTables.RemoveItem (i)
    Next
    txtFieldTypeAcces.Caption = ""
    lblRecordCountAcces.Caption = ""
    lblColumnAcces.Caption = ""
    gridData.ClearFields
    MakeEnabledAcces (False)
End Sub

Private Sub cmdExecuteAcces_Click()
    
    On Error GoTo ErrorTrap
    If Trim(txtSQLAcces.Text) = "" Then Exit Sub
    
    Set gridData.DataSource = Nothing
    adoData.RecordSource = ""

    If adoData.ConnectionString = "" Then
        adoData.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & frmMAin.FileNameAcces & ";"
    End If

    adoData.RecordSource = Trim(txtSQLAcces.Text)
    adoData.Refresh
    Dim i As Integer
        For i = frmMAin.lstAccesTable.ListCount - 1 To 0 Step -1
            frmMAin.lstAccesTable.RemoveItem i
        Next i
    If adoData.Recordset.Fields.Count = 0 Then
        gridData.ClearFields
    Else
        Set gridData.DataSource = adoData.Recordset
        gridData.ClearFields
        gridData.ReBind
    End If
    ReDim Preserve TypeAcces(0)
    With adoData.Recordset
        For Each rsfield In .Fields
                frmMAin.lstAccesTable.AddItem rsfield.Name
                ReDim Preserve TypeAcces(UBound(TypeAcces) + 1)
                TypeAcces(UBound(TypeAcces)) = rsfield.Type 'TypeName(rsField.Type)
        Next
    End With
    lblRecordCountAcces.Caption = adoData.Recordset.RecordCount
    lblColumnAcces.Caption = lstAccesTable.ListCount
    lblDatabaseNameAccess.Caption = txtFileName.Text
    lstAccesTable.ListIndex = 0
    Exit Sub
ErrorTrap:
    MsgBox "An Error occured during connecting Database. err.no : " & Err.Number & " err.desc : " & Err.Description

End Sub

Private Sub cmdExecuteMySQL_Click()
    
  On Error GoTo ErrorTrap
    If Trim(txtSQLMysql.Text) = "" Then Exit Sub
    
    Set gridDataMySQL.DataSource = Nothing
    adoDataMySQL.RecordSource = ""

    If adoDataMySQL.ConnectionString = "" Then
    
    adoDataMySQL.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & _
    ServerNameMySQL & "; DATABASE=" & DBNameMySQL & "; UID=" & UserNameMySQL & _
    ";PWD=" & PasswordMySQL
    
    End If

    adoDataMySQL.RecordSource = Trim(txtSQLMysql.Text)
    adoDataMySQL.Refresh
    Dim i As Integer
        For i = frmMAin.lstMySQLTable.ListCount - 1 To 0 Step -1
            frmMAin.lstMySQLTable.RemoveItem i
        Next i
    If adoDataMySQL.Recordset.Fields.Count = 0 Then
        gridDataMySQL.ClearFields
    Else
        Set gridDataMySQL.DataSource = adoDataMySQL.Recordset
        gridDataMySQL.ClearFields
        gridDataMySQL.ReBind
    End If
    ReDim Preserve TypeMySQL(0)
    With adoDataMySQL.Recordset
        For Each rsfield In .Fields
                frmMAin.lstMySQLTable.AddItem rsfield.Name
                ReDim Preserve TypeMySQL(UBound(TypeMySQL) + 1)
                TypeMySQL(UBound(TypeMySQL)) = rsfield.Type 'TypeName(rsField.Type)
        Next
    End With
    lblRecordCountMySQL.Caption = adoDataMySQL.Recordset.RecordCount
    lblColumnMySQL.Caption = lstMySQLTable.ListCount
    'lblDatabaseNameAccess.Caption = txtFileName.Text
    lstMySQLTable.ListIndex = 0
    Exit Sub
ErrorTrap:
    MsgBox "An Error occured during connecting Database. err.no : " & Err.Number & " err.desc : " & Err.Description

End Sub

Private Sub cmdShowDB_Click()
    For i = cmbAccesTables.ListCount - 1 To 0 Step -1
        cmbAccesTables.RemoveItem i
    Next i

    If GetDBTables(Trim(frmMAin.txtFileName.Text)) = True Then
        MsgBox "Connecting to the Access Database succesfully"
        MakeEnabledAcces (True)
        cmbAccesTables.Text = "Select Table >>"
        For i = 1 To UBound(DB1Tables)
            cmbAccesTables.AddItem DB1Tables(i)
        Next i
    End If
    
    For i = lstAccesTable.ListCount - 1 To 0 Step -1
        lstAccesTable.RemoveItem i
    Next i
    Exit Sub
End Sub


Private Sub editModeAcces_Click()
    If editModeAcces.Value = 1 Then
        gridData.AllowAddNew = True
        gridData.AllowDelete = True
        gridData.AllowUpdate = True
    Else
        gridData.AllowAddNew = False
        gridData.AllowDelete = False
        gridData.AllowUpdate = False
    End If
End Sub

Private Sub editModeMySQL_Click()
    If editModeMySQL.Value = 1 Then
        gridDataMySQL.AllowAddNew = True
        gridDataMySQL.AllowDelete = True
        gridDataMySQL.AllowUpdate = True
    Else
        gridDataMySQL.AllowAddNew = False
        gridDataMySQL.AllowDelete = False
        gridDataMySQL.AllowUpdate = False
    End If
End Sub

Sub Form_Load()
CopyDB = False
    Dim TableName1
        TableName1 = GetSetting("prjConvert", "TableNameAcces", "TableNameAcces", "")
        'TableName2 = GetSetting("prjConvert", "TableNameMySQL", "TableNameMySQL", "")
        txtServerMySQL.Text = GetSetting("prjConvert", "TableNameMySQL", "ServerNameMySQL", "")
        txtDBNameMySQL.Text = GetSetting("prjConvert", "TableNameMySQL", "DBNameMySQL", "")
        txtUserNameMySQL.Text = GetSetting("prjConvert", "TableNameMySQL", "UserNameMySQL", "")
        If Not TableName1 = "" Then
            txtFileName.Text = TableName1
        '        For i = cmbAccesTables.ListCount - 1 To 0 Step -1  ' BURASI ÞÝMDÝLÝK YOK
        '            cmbAccesTable.RemoveItem i
        '        Next i

        '    If GetDBTables(Table1) = True Then
        '        For i = 1 To UBound(DB1Tables)
        '            cmbAccesTables.AddItem DB1Tables(i)
        '        Next i
        '    End If
    
        '    For i = lstAccesTable.ListCount - 1 To 0 Step -1
        '        lstAccesTable.RemoveItem i
        '    Next i
        'End If
        
        'If Not Table2 = "" Then
        '    txtFieldTypeMySQL.Text = Table1
        '    If ConnectDBMySQL Then
        '        For i = 1 To UBound(DB2Tables)
        '            cmbMySQLTables.AddItem DB2Tables(i)
        '        Next i
        '    End If
        '    For i = lstMySQLTable.ListCount - 1 To 0 Step -1
        '        lstMySQLTable.RemoveItem i
        '    Next i
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "prjConvert", "TableNameAcces", "TableNameAcces", txtFileName.Text
    SaveSetting "prjConvert", "TableNameMySQL", "ServerNameMySQL", txtServerMySQL.Text
    SaveSetting "prjConvert", "TableNameMySQL", "DBNameMySQL", frmMAin.txtDBNameMySQL
    SaveSetting "prjConvert", "TableNameMySQL", "UserNameMySQL", frmMAin.txtUserNameMySQL
    EndProgram
End Sub


Private Sub lstAccesTable_Click()
    If lstAccesTable.ListIndex >= 0 Then
        txtFieldTypeAcces.Caption = TypeName(CInt(TypeAcces(lstAccesTable.ListIndex + 1)))
    End If
End Sub

Private Sub lstMySQLTable_Click()
    If lstMySQLTable.ListIndex >= 0 Then
        Label19.Caption = TypeNameMySQL(CInt(TypeMySQL(lstMySQLTable.ListIndex + 1)))
    End If
End Sub

'Eger Focus kaybederse istersek bilgileri yok edelim diye ama gerek olmayabilir...

'Private Sub lstAccesTable_LostFocus()
'    txtFieldTypeAcces.Caption = ""
'    lstAccesTable.ListIndex = -1
'End Sub

Private Sub mnuFileExit_Click()
    EndProgram
    Unload Me
End Sub
Public Sub LoadData(FileName)
    
    'On Error GoTo ErrorTrap

    Set gridData.DataSource = Nothing
    adoData.RecordSource = ""
    adoData.ConnectionString = ""
    adoData.CommandType = adCmdText
    adoData.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";"

    adoData.RecordSource = "SELECT * FROM [" & cmbAccesTables & "]"
    adoData.Refresh
    lblRecordCountAcces.Caption = adoData.Recordset.RecordCount
    lblColumnAcces.Caption = lstAccesTable.ListCount
    lstAccesTable.ListIndex = 0
    If adoData.Recordset.Fields.Count = 0 Then
        gridData.ClearFields
    Else
        Set gridData.DataSource = adoData.Recordset
        gridData.ClearFields
        gridData.ReBind
    End If
    gridData.Caption = "..." & Right(FileName, 20)
    Exit Sub
    
ErrorTrap:
    MsgBox "An Error occured during connecting db. err.no : " & Err.Number & " err.desc : " & Err.Description

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub
Public Sub LoadDataMySQL()
    
    On Error GoTo ErrorTrap

    Set gridDataMySQL.DataSource = Nothing
    adoDataMySQL.RecordSource = ""
    adoDataMySQL.ConnectionString = ""
    adoDataMySQL.CommandType = adCmdText
    adoDataMySQL.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & _
    ServerMySQL & "; DATABASE=" & DBNameMySQL & "; UID=" & UserNameMySQL & _
    ";PWD=" & PasswordMySQL
    
    adoDataMySQL.RecordSource = "SELECT * FROM " & cmbMySQLTables
    adoDataMySQL.Refresh
    lblRecordCountMySQL.Caption = adoDataMySQL.Recordset.RecordCount
    lblColumnMySQL.Caption = lstMySQLTable.ListCount
    If adoDataMySQL.Recordset.Fields.Count = 0 Then
        gridDataMySQL.ClearFields
    Else
        Set gridDataMySQL.DataSource = adoDataMySQL.Recordset
        gridDataMySQL.ClearFields
        gridDataMySQL.ReBind
    End If
    gridDataMySQL.Caption = txtServerMySQL.Text & "-> " & _
        txtDBNameMySQL.Text & "-> " & cmbMySQLTables.Text
    Exit Sub
    
ErrorTrap:
    MsgBox "An Error occured during connecting db. err.no : " & Err.Number & " err.desc : " & Err.Description
End Sub



