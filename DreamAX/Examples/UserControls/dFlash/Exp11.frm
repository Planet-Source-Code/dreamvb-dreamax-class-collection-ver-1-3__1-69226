VERSION 5.00
Begin VB.Form FrmExp11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flash Window - ActiveX Example"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Exp11.dFlash dFlash1 
      Left            =   2445
      Top             =   210
      _extentx        =   582
      _extenty        =   582
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1305
      TabIndex        =   1
      Top             =   195
      Width           =   1000
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Start Flash"
      Height          =   375
      Left            =   165
      TabIndex        =   0
      Top             =   195
      Width           =   1000
   End
End
Attribute VB_Name = "FrmExp11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click()
Static iStop As Boolean
    iStop = (Not iStop)
    
    If (iStop) Then
        cmd1.Caption = "Stop Flash"
        dFlash1.FlashWindowwStart
    Else
        cmd1.Caption = "Start Flash"
        dFlash1.FlashWindowStop
    End If
End Sub

Private Sub cmd3_Click()
    If (dFlash1.Busy) Then dFlash1.FlashWindowStop
    'Unload the program
    Unload FrmExp11
End Sub

Private Sub Form_Load()
    dFlash1.FormObject FrmExp11
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp11 = Nothing
End Sub
