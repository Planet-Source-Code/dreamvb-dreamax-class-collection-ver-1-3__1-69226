VERSION 5.00
Begin VB.Form FrmEx1 
   Caption         =   "INI Class - Example"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   3135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1875
      TabIndex        =   6
      Top             =   3045
      Width           =   1215
   End
   Begin VB.TextBox TxtKeyVal 
      Height          =   360
      Left            =   195
      TabIndex        =   5
      Top             =   3075
      Width           =   1500
   End
   Begin VB.ListBox LstKeys 
      Height          =   1425
      ItemData        =   "Ex1.frx":0000
      Left            =   195
      List            =   "Ex1.frx":0002
      TabIndex        =   3
      Top             =   1260
      Width           =   1530
   End
   Begin VB.ComboBox CboSel 
      Height          =   315
      ItemData        =   "Ex1.frx":0004
      Left            =   195
      List            =   "Ex1.frx":0006
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   465
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Key Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   4
      Top             =   2790
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selection Keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   930
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Selections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   900
   End
End
Attribute VB_Name = "FrmEx1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cINI As dINIFile

Private Function FixPath(lPath As String) As String
    If Right(lPath, 1) <> "\" Then
        FixPath = lPath & "\"
    Else
        FixPath = lPath
    End If
End Function

Private Sub CboSel_Click()
Dim col As Collection
Dim Item

    LstKeys.Clear
    'Fills a listbox with all the keys from a selection.
    Set col = cINI.GetIniSelection(CboSel.Text)
    For Each Item In col
        LstKeys.AddItem Item
    Next Item
    Set col = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload FrmEx1
End Sub

Private Sub Form_Load()
Dim col As Collection
Dim Item

    Set cINI = New dINIFile
    'Set the ini filename
    cINI.FileName = FixPath(App.Path) & "demo.ini"
    'Check if the file is found
    If Not cINI.IniFound Then
        MsgBox "demo.ini not found"
        Unload FrmEx1
    Else
        Set col = cINI.GetSelectionNames
        For Each Item In col
            CboSel.AddItem Item
        Next Item
        Set col = Nothing
        CboSel.ListIndex = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmEx1 = Nothing
End Sub

Private Sub LstKeys_Click()
    TxtKeyVal.Text = cINI.INIReadKey(CboSel.Text, LstKeys.Text, "ERROR")
End Sub
