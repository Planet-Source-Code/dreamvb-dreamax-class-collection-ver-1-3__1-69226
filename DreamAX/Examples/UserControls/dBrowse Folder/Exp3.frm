VERSION 5.00
Begin VB.Form FrmExp3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse For Folder ActiveX - Example"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd3 
      Caption         =   "Example3"
      Height          =   350
      Left            =   2400
      TabIndex        =   4
      Top             =   105
      Width           =   1005
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Example2"
      Height          =   350
      Left            =   1260
      TabIndex        =   3
      Top             =   105
      Width           =   1005
   End
   Begin Exp3.dBrowseFolder dBrowseFolder1 
      Left            =   120
      Top             =   900
      _ExtentX        =   794
      _ExtentY        =   714
      DialogTitle     =   "Select Folder:"
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Example1"
      Height          =   350
      Left            =   150
      TabIndex        =   1
      Top             =   105
      Width           =   1005
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   3540
      TabIndex        =   0
      Top             =   105
      Width           =   1005
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      Height          =   735
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   4890
   End
End
Attribute VB_Name = "FrmExp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click()
    'Example 1
    With dBrowseFolder1
        .HwndOwner = FrmExp3.hwnd
        .Flags = bRETURNONLYFSDIRS Or bNEWDIALOGSTYLE
        .RootFolder = Desktop
        .ShowBrowseForFolder
        lbl1.Caption = .Directory
    End With
End Sub

Private Sub cmd2_Click()
    'Example 2
    MsgBox "Using Special Folder Location"
    With dBrowseFolder1
        .HwndOwner = FrmExp3.hwnd
        .Flags = bRETURNONLYFSDIRS Or bNEWDIALOGSTYLE
        .RootFolder = StartMenu
        Call .ShowBrowseForFolder
        lbl1.Caption = .Directory
    End With
End Sub

Private Sub cmd3_Click()
    'Example 3
    MsgBox "Using Call backs start folder ProgramFiles"
    With dBrowseFolder1
        .HwndOwner = FrmExp3.hwnd
        .Flags = 0
        .StartDirectory = Environ("ProgramFiles")
        .RootFolder = NoSpecialFolder
        Call .ShowBrowseForFolder
        lbl1.Caption = .Directory
    End With
End Sub

Private Sub cmdexit_Click()
    Unload FrmExp3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp3 = Nothing
End Sub

