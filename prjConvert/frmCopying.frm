VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCopying 
   Caption         =   "Copying Database"
   ClientHeight    =   4140
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCopying.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Report"
      Filter          =   "*.txt"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Report"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Copying Files ..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmCopying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    CommonDialog1.ShowSave
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
