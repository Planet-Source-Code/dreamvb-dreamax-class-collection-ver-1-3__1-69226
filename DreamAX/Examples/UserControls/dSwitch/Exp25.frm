VERSION 5.00
Begin VB.Form FrmExp25 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Switch ActiveX -  Example"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Exp25.dSwitch dSwitch3 
      Height          =   705
      Left            =   315
      TabIndex        =   5
      Top             =   1845
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1244
      ButtonState     =   -1  'True
   End
   Begin Exp25.dSwitch dSwitch2 
      CausesValidation=   0   'False
      Height          =   705
      Left            =   2160
      TabIndex        =   3
      Top             =   915
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1244
      BackColor       =   12648447
      OnColorState    =   16711680
      OffColorState   =   33023
      ButtonState     =   -1  'True
      AllowToggleSupport=   0   'False
   End
   Begin Exp25.dSwitch dSwitch1 
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1244
      AllowToggleSupport=   0   'False
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3885
      TabIndex        =   0
      Top             =   3210
      Width           =   1035
   End
   Begin Exp25.dSwitch dSwitch4 
      Height          =   705
      Left            =   330
      TabIndex        =   7
      Top             =   2745
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1244
      OnColorState    =   8454143
      ButtonState     =   -1  'True
      Enabled         =   0   'False
      AllowToggleSupport=   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "This button is disabled"
      Height          =   225
      Left            =   885
      TabIndex        =   8
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Button with Toggle key support"
      Height          =   240
      Left            =   960
      TabIndex        =   6
      Top             =   2100
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Switch background and on color and off colors"
      Height          =   405
      Left            =   210
      TabIndex        =   4
      Top             =   1125
      Width           =   1860
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Switch Value : Off"
      Height          =   195
      Left            =   675
      TabIndex        =   2
      Top             =   405
      Width           =   1275
   End
End
Attribute VB_Name = "FrmExp25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrmExp25
End Sub

Private Sub dSwitch1_StateChange(ButtonState As TState)
    If (ButtonState = bOn) Then
        lblVal.Caption = "Switch Value : On"
    Else
        lblVal.Caption = "Switch Value : Off"
    End If
End Sub

Private Sub dSwitch2_StateChange(ButtonState As TState)
    If (ButtonState = bOn) Then
        lblVal.Caption = "Switch Value : On"
    Else
        lblVal.Caption = "Switch Value : Off"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp25 = Nothing
End Sub

