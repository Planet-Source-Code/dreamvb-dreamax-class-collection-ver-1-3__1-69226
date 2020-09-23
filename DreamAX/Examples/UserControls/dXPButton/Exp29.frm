VERSION 5.00
Begin VB.Form FrmExp29 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ActiveX XP Button - Example"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin Exp29.dXPButton dXPButton8 
      Height          =   450
      Left            =   3960
      TabIndex        =   7
      Top             =   1920
      Width           =   1185
      _extentx        =   2090
      _extenty        =   794
      style           =   2
      font            =   "Exp29.frx":0000
      forecolor       =   255
      caption         =   "Exit"
   End
   Begin Exp29.dXPButton dXPButton5 
      Height          =   450
      Left            =   1590
      TabIndex        =   4
      Top             =   750
      Width           =   765
      _extentx        =   1349
      _extenty        =   794
      alignment       =   0
      font            =   "Exp29.frx":0024
      caption         =   "Left"
   End
   Begin Exp29.dXPButton dXPButton1 
      Height          =   450
      Left            =   270
      TabIndex        =   0
      Top             =   165
      Width           =   1200
      _extentx        =   2117
      _extenty        =   794
      font            =   "Exp29.frx":0050
      caption         =   "Cool"
   End
   Begin Exp29.dXPButton dXPButton2 
      Height          =   450
      Left            =   1530
      TabIndex        =   1
      Top             =   165
      Width           =   1200
      _extentx        =   2117
      _extenty        =   794
      style           =   2
      font            =   "Exp29.frx":007C
      caption         =   "XP"
   End
   Begin Exp29.dXPButton dXPButton3 
      Height          =   450
      Left            =   2820
      TabIndex        =   2
      Top             =   165
      Width           =   1200
      _extentx        =   2117
      _extenty        =   794
      style           =   3
      font            =   "Exp29.frx":00A8
      caption         =   "Button"
   End
   Begin Exp29.dXPButton dXPButton4 
      Height          =   450
      Left            =   270
      TabIndex        =   3
      Top             =   735
      Width           =   1200
      _extentx        =   2117
      _extenty        =   794
      default         =   -1  'True
      style           =   2
      font            =   "Exp29.frx":00D4
      caption         =   "Default"
   End
   Begin Exp29.dXPButton dXPButton6 
      Height          =   450
      Left            =   2475
      TabIndex        =   5
      Top             =   750
      Width           =   765
      _extentx        =   1349
      _extenty        =   794
      font            =   "Exp29.frx":0100
      caption         =   "Center"
   End
   Begin Exp29.dXPButton dXPButton7 
      Height          =   450
      Left            =   3330
      TabIndex        =   6
      Top             =   750
      Width           =   765
      _extentx        =   1349
      _extenty        =   794
      alignment       =   0
      font            =   "Exp29.frx":012C
      caption         =   "Right"
   End
End
Attribute VB_Name = "FrmExp29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dXPButton8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload FrmExp29
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp29 = Nothing
End Sub
