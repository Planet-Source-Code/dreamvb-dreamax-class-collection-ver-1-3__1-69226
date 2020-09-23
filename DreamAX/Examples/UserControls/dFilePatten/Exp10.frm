VERSION 5.00
Begin VB.Form FrmExp10 
   Caption         =   "File Patten - ActiveX - Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin Exp10.dFilePatten dFilePatten1 
      Height          =   315
      Left            =   270
      TabIndex        =   2
      Top             =   225
      Width           =   2055
      _extentx        =   3625
      _extenty        =   556
      font            =   "Exp10.frx":0000
      font            =   "Exp10.frx":002C
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   240
      TabIndex        =   1
      Top             =   570
      Width           =   2100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   405
      Left            =   2490
      TabIndex        =   0
      Top             =   2370
      Width           =   1215
   End
End
Attribute VB_Name = "FrmExp10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload FrmExp10
End Sub

Private Sub Form_Load()
    'Add File Pattens
    dFilePatten1.Patten = "All Files(*.*)|TextFiles(*.txt)|" _
    & "Bitmaps(*.bmp)"
    File1.Path = Environ("WinDir")
    'Set the file Listbox to use the file patten control
    dFilePatten1.vFileListBox = File1
    'Select the first patten as default All Files *.*
    dFilePatten1.PattenIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp10 = Nothing
End Sub
