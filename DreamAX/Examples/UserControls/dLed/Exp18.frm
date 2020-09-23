VERSION 5.00
Begin VB.Form FrmExp18 
   Caption         =   "Led - ActiveX Example"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin Exp18.dLed dLed2 
      Height          =   255
      Left            =   1410
      Top             =   180
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   450
      Enabled         =   0   'False
   End
   Begin Exp18.dLed dLed1 
      Height          =   225
      Index           =   0
      Left            =   270
      Top             =   180
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
   End
   Begin VB.Timer Tmr1 
      Interval        =   300
      Left            =   450
      Top             =   660
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   2085
      TabIndex        =   0
      Top             =   870
      Width           =   960
   End
   Begin Exp18.dLed dLed1 
      Height          =   225
      Index           =   1
      Left            =   510
      Top             =   180
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      OnColor         =   255
      OffColor        =   192
   End
   Begin Exp18.dLed dLed1 
      Height          =   225
      Index           =   2
      Left            =   765
      Top             =   180
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      OnColor         =   12648384
      OffColor        =   32768
   End
   Begin Exp18.dLed dLed1 
      Height          =   225
      Index           =   3
      Left            =   1020
      Top             =   180
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      OnColor         =   12648447
      OffColor        =   32896
   End
End
Attribute VB_Name = "FrmExp18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrmExp18
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (Tmr1.Enabled) Then Tmr1.Enabled = False
    Set FrmExp18 = Nothing
End Sub

Private Sub Tmr1_Timer()
    dLed1(0).Value = Not dLed1(0).Value
    dLed1(1).Value = Not dLed1(0).Value
    dLed1(2).Value = Not dLed1(1).Value
    dLed1(3).Value = Not dLed1(2).Value
End Sub
