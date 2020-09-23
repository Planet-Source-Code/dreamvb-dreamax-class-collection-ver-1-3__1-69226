VERSION 5.00
Begin VB.Form FrmExp1 
   Caption         =   "Animated Cursor ActiveX - Example"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin Exp1.dAniCursor dAniCursor1 
      Height          =   555
      Left            =   135
      TabIndex        =   2
      Top             =   195
      Width           =   555
      _extentx        =   979
      _extenty        =   979
      panelborder     =   2
      backcolor       =   12640511
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Top             =   1005
      Width           =   840
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1005
      Width           =   795
   End
   Begin Exp1.dAniCursor dAniCursor2 
      Height          =   555
      Left            =   825
      TabIndex        =   3
      Top             =   195
      Width           =   2010
      _extentx        =   3545
      _extenty        =   979
      panelborder     =   3
   End
End
Attribute VB_Name = "FrmExp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    If (dAniCursor1.IsBusy) Then dAniCursor1.cStop
    Unload FrmExp1
End Sub

Private Sub cmdPlay_Click()
    If (cmdPlay.Caption = "&Play") Then
        dAniCursor1.cPlay
        dAniCursor2.cPlay
        cmdPlay.Caption = "&Stop"
    Else
        cmdPlay.Caption = "&Play"
        dAniCursor1.cStop
        dAniCursor2.cStop
    End If
End Sub

Private Sub Form_Load()
    dAniCursor1.Filename = App.Path & "\resources\example.ani"
    dAniCursor2.Filename = App.Path & "\resources\drum.ani"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp1 = Nothing
End Sub
