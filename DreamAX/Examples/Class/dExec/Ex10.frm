VERSION 5.00
Begin VB.Form FrnEx10 
   Caption         =   "dExec - Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtParm 
      Height          =   375
      Left            =   2265
      TabIndex        =   11
      Text            =   "- Test123"
      Top             =   1230
      Width           =   3015
   End
   Begin VB.TextBox txtDir 
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Text            =   "C:\"
      Top             =   1215
      Width           =   1845
   End
   Begin VB.ComboBox cboWinShow 
      Height          =   315
      Left            =   2265
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2040
      Width           =   2400
   End
   Begin VB.ComboBox cboOp 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   1470
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Text            =   "Example.exe"
      Top             =   480
      Width           =   5310
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   375
      Left            =   3090
      TabIndex        =   1
      Top             =   2715
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4500
      TabIndex        =   0
      Top             =   2700
      Width           =   990
   End
   Begin VB.Label lblRet 
      AutoSize        =   -1  'True
      Caption         =   "Return Code:"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   2505
      Width           =   945
   End
   Begin VB.Label lblParm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parameters:"
      Height          =   195
      Left            =   2265
      TabIndex        =   10
      Top             =   975
      Width           =   840
   End
   Begin VB.Label lblDir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directory:"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   975
      Width           =   675
   End
   Begin VB.Label lblWinShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window State:"
      Height          =   195
      Left            =   2265
      TabIndex        =   6
      Top             =   1770
      Width           =   1050
   End
   Begin VB.Label lblOp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operation:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1770
      Width           =   735
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File / URL, Path"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   210
      Width           =   1155
   End
End
Attribute VB_Name = "FrnEx10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrnEx10
End Sub

Private Sub Command1_Click()
Dim cExec As New dExec
    
    With cExec
        .Hwnd = FrnEx10.Hwnd
        .ShowCmd = cboWinShow.ListIndex
        .Exec cboOp.ListIndex, txtFile.Text, txtParm.Text, txtDir.Text
        lblRet.Caption = "Return code: " & .ReturnCode
    End With
    
End Sub

Private Sub Form_Load()
    'Operations
    cboOp.AddItem "None"
    cboOp.AddItem "Edit"
    cboOp.AddItem "Explore"
    cboOp.AddItem "Find"
    cboOp.AddItem "Open"
    cboOp.AddItem "Print"
    cboOp.ListIndex = 4
    'Window State
    cboWinShow.AddItem "SW_HIDE"
    cboWinShow.AddItem "SW_SHOWNORMAL"
    cboWinShow.AddItem "SW_SHOWMINIMIZED"
    cboWinShow.AddItem "SW_SHOWMAXIMIZED"
    cboWinShow.AddItem "SW_SHOWNOACTIVATE"
    cboWinShow.AddItem "SW_SHOW"
    cboWinShow.AddItem "SW_MINIMIZE"
    cboWinShow.AddItem "SW_SHOWMINNOACTIVE"
    cboWinShow.AddItem "SW_SHOWNA"
    cboWinShow.AddItem "SW_RESTORE"
    cboWinShow.AddItem "SW_SHOWDEFAULT"
    cboWinShow.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrnEx10 = Nothing
End Sub
