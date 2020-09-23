VERSION 5.00
Begin VB.Form FrmExp20 
   Caption         =   "OutLine Border ActiveX - Example"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin Exp20.dOutline dOutline2 
      Height          =   780
      Left            =   1095
      TabIndex        =   2
      Top             =   195
      Width           =   960
      _extentx        =   1693
      _extenty        =   1429
      topline         =   8421631
      leftline        =   8421631
      rightline       =   8421631
      bottomline      =   8421631
      backcolor       =   12632319
      backstyle       =   1
   End
   Begin Exp20.dOutline dOutline1 
      Height          =   810
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   900
      _extentx        =   1588
      _extenty        =   1429
      topline         =   16777215
      leftline        =   16777215
      rightline       =   -2147483636
      bottomline      =   -2147483636
      backcolor       =   16777215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   990
      Width           =   1215
   End
End
Attribute VB_Name = "FrmExp20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload FrmExp20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp20 = Nothing
End Sub
