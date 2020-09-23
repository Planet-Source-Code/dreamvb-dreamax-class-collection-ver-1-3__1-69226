VERSION 5.00
Begin VB.Form FrmExp35 
   BackColor       =   &H00C0E0FF&
   Caption         =   "dTxtBoxBorder ActiveX - Example"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin Exp35.dTxtBoxBorder dTxtBoxBorder2 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   570
      Width           =   1965
      _extentx        =   3466
      _extenty        =   450
      font            =   "Exp35.frx":0000
      text            =   ""
      bordercolor     =   8438015
   End
   Begin Exp35.dTxtBoxBorder dTxtBoxBorder1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2355
      _extentx        =   4154
      _extenty        =   450
      backcolor       =   16761024
      font            =   "Exp35.frx":002C
      text            =   "Textbox with color border"
   End
   Begin VB.CommandButton mnuexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3735
      TabIndex        =   0
      Top             =   945
      Width           =   900
   End
   Begin Exp35.dTxtBoxBorder dTxtBoxBorder3 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   885
      Width           =   2985
      _extentx        =   5265
      _extenty        =   450
      font            =   "Exp35.frx":0058
      text            =   "Border shows when control has focus"
      bordercolor     =   33023
      showborder      =   0   'False
   End
End
Attribute VB_Name = "FrmExp35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dTxtBoxBorder3_GotFocus()
    dTxtBoxBorder3.ShowBorder = True
End Sub

Private Sub dTxtBoxBorder3_LostFocus()
    dTxtBoxBorder3.ShowBorder = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp35 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp35
End Sub
