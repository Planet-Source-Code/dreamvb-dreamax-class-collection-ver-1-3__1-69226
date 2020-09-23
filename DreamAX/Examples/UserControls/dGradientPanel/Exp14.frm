VERSION 5.00
Begin VB.Form FrmExp14 
   Caption         =   "GradientPanel - ActiveX Example"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin Exp14.dGradientPanel dGradientPanel1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4515
      _extentx        =   7964
      _extenty        =   1058
      startcolor      =   13665116
      endcolor        =   9452297
      direction       =   1
      labeltext       =   "Visual Basic .NET"
      labeltop        =   8
      labelleft       =   8
      labelfont       =   "Exp14.frx":0000
      labelforecolor  =   16777215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3435
      TabIndex        =   0
      Top             =   1515
      Width           =   960
   End
End
Attribute VB_Name = "FrmExp14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrmExp14
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp14 = Nothing
End Sub
