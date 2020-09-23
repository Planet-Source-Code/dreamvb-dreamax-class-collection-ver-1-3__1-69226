VERSION 5.00
Begin VB.Form FrnEx3 
   Caption         =   "dDateLimit - Example"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   1695
      Width           =   1215
   End
End
Attribute VB_Name = "FrnEx3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cTrial As dDateLimit
Attribute cTrial.VB_VarHelpID = -1

Private Sub Form_Load()
    Set cTrial = New dDateLimit
   ' cTrial.Active = False 'turns of the date checking
    cTrial.MaxDate = "2/20/2008"
End Sub

Sub cTrial_DateUp()
    MsgBox "This program's Trial time has ended.", vbExclamation, FrnEx3.Caption
    cmdExit_Click
End Sub

Private Sub cmdExit_Click()
    Unload FrnEx3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrnEx3 = Nothing
End Sub
