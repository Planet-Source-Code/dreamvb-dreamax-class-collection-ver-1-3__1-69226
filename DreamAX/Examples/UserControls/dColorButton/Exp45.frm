VERSION 5.00
Begin VB.Form FrmExp45 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "dColorButton ActiveX - Example"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Exp45.dColorButton dColorButton2 
      Height          =   255
      Left            =   555
      TabIndex        =   2
      Top             =   195
      Width           =   270
      _extentx        =   476
      _extenty        =   450
      buttoncolor     =   255
   End
   Begin Exp45.dColorButton dColorButton1 
      Height          =   255
      Left            =   195
      TabIndex        =   1
      Top             =   195
      Width           =   270
      _extentx        =   476
      _extenty        =   450
      buttoncolor     =   0
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   1035
      Width           =   885
   End
End
Attribute VB_Name = "FrmExp45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload FrmExp45
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp45 = Nothing
End Sub
