VERSION 5.00
Begin VB.Form FrnEx4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dChecksum- Example"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
Attribute VB_Name = "FrnEx4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrnEx4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrnEx4 = Nothing
End Sub
