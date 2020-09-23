VERSION 5.00
Begin VB.Form FrmExp6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Button ActiveX - Example"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Exp6.dCheckButton cmdExit 
      Height          =   375
      Left            =   3075
      TabIndex        =   3
      Top             =   1545
      Width           =   1155
      _extentx        =   2037
      _extenty        =   661
      caption         =   "E&xit"
      captionalign    =   1
      font            =   "Exp6.frx":0000
   End
   Begin Exp6.dCheckButton dCheckButton3 
      Height          =   375
      Left            =   2850
      TabIndex        =   2
      Top             =   840
      Width           =   1425
      _extentx        =   2514
      _extenty        =   661
      caption         =   "Normal Button"
      captionalign    =   1
      font            =   "Exp6.frx":002C
      forecolor       =   16711680
      backcolor       =   16770764
   End
   Begin Exp6.dCheckButton dCheckButton2 
      Height          =   375
      Left            =   1575
      TabIndex        =   1
      Top             =   525
      Width           =   1155
      _extentx        =   2037
      _extenty        =   661
      caption         =   "CheckButton"
      captionalign    =   1
      font            =   "Exp6.frx":0058
      checked         =   -1
   End
   Begin Exp6.dCheckButton dCheckButton1 
      Height          =   390
      Left            =   150
      TabIndex        =   0
      Top             =   135
      Width           =   1305
      _extentx        =   2302
      _extenty        =   688
      picture         =   "Exp6.frx":0084
      caption         =   "CheckButton"
      font            =   "Exp6.frx":03D6
   End
End
Attribute VB_Name = "FrmExp6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrmExp6
End Sub

Private Sub dCheckButton1_Click()
    FrmExp6.Caption = dCheckButton1.Checked
End Sub

Private Sub dCheckButton3_Click()
    dCheckButton3.Checked = Not dCheckButton3.Checked
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp6 = Nothing
End Sub
