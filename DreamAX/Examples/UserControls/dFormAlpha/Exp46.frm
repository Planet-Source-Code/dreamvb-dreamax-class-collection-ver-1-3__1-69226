VERSION 5.00
Begin VB.Form FrmExp46 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "dFormAlpha ActiveX - Example"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Exp46.dFormAlpha dFormAlpha1 
      Left            =   2100
      Top             =   825
      _ExtentX        =   794
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Disable"
      Height          =   375
      Left            =   2625
      TabIndex        =   3
      Top             =   855
      Width           =   885
   End
   Begin VB.HScrollBar HsbValue 
      Height          =   255
      Left            =   300
      Max             =   255
      TabIndex        =   2
      Top             =   450
      Value           =   128
      Width           =   4215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3630
      TabIndex        =   0
      Top             =   855
      Width           =   885
   End
   Begin VB.Label lblAlphaLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha Lavel: 128"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   195
      Width           =   1200
   End
End
Attribute VB_Name = "FrmExp46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnable_Click()
    If (cmdEnable.Caption = "Enable") Then
        cmdEnable.Caption = "Disable"
        dFormAlpha1.Enabled = True
    Else
        cmdEnable.Caption = "Enable"
        dFormAlpha1.Enabled = False
    End If
End Sub

Private Sub cmdexit_Click()
    Unload FrmExp46
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp46 = Nothing
End Sub

Private Sub HsbValue_Change()
    lblAlphaLevel.Caption = "Alpha Lavel: " & HsbValue.Value
    dFormAlpha1.Alpha = HsbValue.Value
End Sub

Private Sub HsbValue_Scroll()
    Call HsbValue_Change
End Sub
