VERSION 5.00
Begin VB.Form FrmExp31 
   Caption         =   "dShape ActiveX - Example"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin Exp31.dShape dShape6 
      Height          =   480
      Left            =   3045
      TabIndex        =   6
      Top             =   1920
      Width           =   1485
      _extentx        =   2619
      _extenty        =   847
      backcolor       =   12648447
      backstyle       =   1
      shape           =   2
      parentbackcolor =   0   'False
      caption         =   "Exit"
      alignment       =   1
      font            =   "Exp31.frx":0000
   End
   Begin Exp31.dShape dShape5 
      Height          =   420
      Left            =   255
      TabIndex        =   5
      Top             =   1905
      Width           =   1620
      _extentx        =   2858
      _extenty        =   741
      backcolor       =   16777215
      backstyle       =   1
      caption         =   ""
      alignment       =   1
      font            =   "Exp31.frx":0028
   End
   Begin Exp31.dShape dShape4 
      Height          =   495
      Left            =   165
      TabIndex        =   3
      Top             =   975
      Width           =   2055
      _extentx        =   3625
      _extenty        =   873
      backcolor       =   16777215
      backstyle       =   1
      shape           =   4
      parentbackcolor =   0   'False
      caption         =   "Shape with Caption"
      alignment       =   1
      font            =   "Exp31.frx":0050
   End
   Begin Exp31.dShape dShape3 
      Height          =   495
      Left            =   2055
      TabIndex        =   2
      Top             =   225
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      backcolor       =   16777215
      backstyle       =   1
      bordercolor     =   33023
      shape           =   3
      caption         =   ""
      font            =   "Exp31.frx":007C
   End
   Begin Exp31.dShape dShape2 
      Height          =   495
      Left            =   1005
      TabIndex        =   1
      Top             =   225
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      backcolor       =   33023
      backstyle       =   1
      shape           =   5
      parentbackcolor =   0   'False
      caption         =   ""
      font            =   "Exp31.frx":00A8
   End
   Begin Exp31.dShape dShape1 
      Height          =   495
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   645
      _extentx        =   1138
      _extenty        =   873
      backcolor       =   8421504
      backstyle       =   1
      bordercolor     =   16777215
      borderstyle     =   3
      parentbackcolor =   0   'False
      caption         =   ""
      font            =   "Exp31.frx":00D4
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shape with hover effects"
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   1605
      Width           =   1770
   End
End
Attribute VB_Name = "FrmExp31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dShape5_HoverIn()
    dShape5.Caption = "HoverIn"
End Sub

Private Sub dShape5_HoverOut()
    dShape5.Caption = "HoverOut"
End Sub

Private Sub dShape6_HoverIn()
    dShape6.BorderWidth = 2
    dShape6.ForeColor = vbBlue
    dShape6.BorderColor = vb3DShadow
    dShape6.BackColor = &HFFFF&
End Sub

Private Sub dShape6_HoverOut()
    dShape6.BorderWidth = 1
    dShape6.ForeColor = vbBlack
    dShape6.BorderColor = vbBlack
    dShape6.BackColor = &HC0FFFF
End Sub

Private Sub dShape6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Unload FrmExp31
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp31 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp31
End Sub
