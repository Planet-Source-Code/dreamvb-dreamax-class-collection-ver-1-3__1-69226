VERSION 5.00
Begin VB.Form FrmEx8 
   Caption         =   "dDataMerge - Example"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2475
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Ex8.frx":0000
      Top             =   195
      Width           =   5535
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2805
      Width           =   1770
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "Merge"
      Height          =   375
      Left            =   645
      TabIndex        =   1
      Top             =   2805
      Width           =   1770
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2805
      Width           =   1215
   End
End
Attribute VB_Name = "FrmEx8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Merge As dDataMerge
Attribute Merge.VB_VarHelpID = -1

Private Sub cmdExtract_Click()
    Merge.ExtractTo "C:\demo.exe", "C:\Picture.gif"
End Sub

Private Sub cmdMerge_Click()
    Merge.MergeTo "C:\demo.exe", "C:\Picture.gif"
End Sub

Private Sub Form_Load()
    Set Merge = New dDataMerge
End Sub

Private Sub Merge_Error()
    Select Case Err.Number
        Case 53
            MsgBox Err.Description & vbCrLf & Err.Source, vbInformation, "Error_" & Err.Number
        Case 54
            MsgBox Err.Description & vbCrLf & Err.Source, vbInformation, "Error_" & Err.Number
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload FrmEx8
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmEx8 = Nothing
End Sub

