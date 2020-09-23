VERSION 5.00
Begin VB.Form FrmExp37 
   Caption         =   "dTrafficLight ActiveX - Example"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin Exp37.dTrafficLight dTrafficLight1 
      Height          =   930
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   810
      _extentx        =   1429
      _extenty        =   1640
   End
   Begin VB.CommandButton mnuexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2805
      TabIndex        =   0
      Top             =   1515
      Width           =   900
   End
   Begin Exp37.dTrafficLight dTrafficLight2 
      Height          =   930
      Left            =   900
      TabIndex        =   2
      Top             =   180
      Width           =   810
      _extentx        =   1429
      _extenty        =   1640
      showlight       =   2
   End
   Begin Exp37.dTrafficLight dTrafficLight3 
      Height          =   930
      Left            =   1755
      TabIndex        =   3
      Top             =   180
      Width           =   810
      _extentx        =   1429
      _extenty        =   1640
      showlight       =   3
   End
End
Attribute VB_Name = "FrmExp37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp37 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp37
End Sub
