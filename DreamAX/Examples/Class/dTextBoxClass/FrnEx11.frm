VERSION 5.00
Begin VB.Form FrnEx11 
   Caption         =   "dTextBoxClass - Example"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1890
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "FrnEx11.frx":0000
      Top             =   90
      Width           =   5610
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   4920
      TabIndex        =   0
      Top             =   2295
      Width           =   990
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnucpy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselall 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnugoto 
         Caption         =   "&Goto"
      End
      Begin VB.Menu mnuDelLine 
         Caption         =   "Delete Line"
      End
   End
   Begin VB.Menu mnumore 
      Caption         =   "&More"
      Begin VB.Menu mnulinecnt 
         Caption         =   "&Line Count"
      End
      Begin VB.Menu CurLine 
         Caption         =   "CurrentLine"
      End
      Begin VB.Menu LineLen 
         Caption         =   "Line Length"
      End
      Begin VB.Menu mnuseltext 
         Caption         =   "Selected Text"
      End
      Begin VB.Menu mnuappend 
         Caption         =   "&Append"
      End
   End
End
Attribute VB_Name = "FrnEx11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TxtCls As dTextBoxClass

Private Sub cmdExit_Click()
    Unload FrnEx11
End Sub

Private Sub CurLine_Click()
    MsgBox TxtCls.LineIndex
End Sub

Private Sub Form_Load()
    Set TxtCls = New dTextBoxClass
    TxtCls.TextBox = Text1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrnEx11 = Nothing
End Sub

Private Sub LineLen_Click()
    MsgBox TxtCls.LineLength(TxtCls.LineIndex)
End Sub

Private Sub mnuappend_Click()
    TxtCls.AppendText vbCrLf & "This is a new line of text..."
    TxtCls.SelectionStart = Len(Text1.Text)
End Sub

Private Sub mnucpy_Click()
    TxtCls.Copy
End Sub

Private Sub mnucut_Click()
    TxtCls.Cut
End Sub

Private Sub mnuDel_Click()
    TxtCls.Delete
End Sub

Private Sub mnuDelLine_Click()
    TxtCls.LineDelete 1
    MsgBox "Line 1 was deleted"
End Sub

Private Sub mnugoto_Click()
    TxtCls.GotoLine 2
    MsgBox "Now at Line 2"
End Sub

Private Sub mnulinecnt_Click()
    MsgBox TxtCls.LinesCount
End Sub

Private Sub mnupaste_Click()
    TxtCls.Paste
End Sub

Private Sub mnuselall_Click()
    TxtCls.SelectAll
End Sub

Private Sub mnuseltext_Click()
    MsgBox TxtCls.SelectedText
End Sub
