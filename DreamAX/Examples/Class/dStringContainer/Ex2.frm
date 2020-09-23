VERSION 5.00
Begin VB.Form FrnEx2 
   Caption         =   "StringContainer - Example"
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
Attribute VB_Name = "FrnEx2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrnEx2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrnEx2 = Nothing
End Sub
