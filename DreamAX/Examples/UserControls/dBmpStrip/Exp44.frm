VERSION 5.00
Begin VB.Form FrmExp44 
   Caption         =   "dBmpStrip - ActiveX Example"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd2 
      Caption         =   "Start"
      Height          =   375
      Left            =   1380
      TabIndex        =   3
      Top             =   1140
      Width           =   990
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Start"
      Height          =   375
      Left            =   195
      TabIndex        =   2
      Top             =   1140
      Width           =   990
   End
   Begin VB.Timer Tmr2 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   3525
      Top             =   240
   End
   Begin VB.Timer Tmr1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   255
   End
   Begin Exp44.dBmpStrip dBmpStrip1 
      Height          =   870
      Left            =   195
      TabIndex        =   1
      Top             =   135
      Width           =   990
      _extentx        =   1746
      _extenty        =   1535
      srcpicture      =   "Exp44.frx":0000
      frames          =   4
      framewidth      =   32
      frameheight     =   32
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2565
      TabIndex        =   0
      Top             =   1140
      Width           =   900
   End
   Begin Exp44.dBmpStrip dBmpStrip2 
      Height          =   870
      Left            =   1380
      TabIndex        =   4
      Top             =   135
      Width           =   990
      _extentx        =   1746
      _extenty        =   1535
      srcpicture      =   "Exp44.frx":0894
      frames          =   4
      framewidth      =   32
      frameheight     =   32
      framedirection  =   1
      borderstyle     =   0
   End
End
Attribute VB_Name = "FrmExp44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click()
    If (cmd1.Caption = "Start") Then
        cmd1.Caption = "Stop"
        Tmr1.Enabled = True
    Else
        cmd1.Caption = "Start"
        Tmr1.Enabled = False
    End If
End Sub

Private Sub Cmd2_Click()
    If (Cmd2.Caption = "Start") Then
        Cmd2.Caption = "Stop"
        Tmr2.Enabled = True
    Else
        Cmd2.Caption = "Start"
        Tmr2.Enabled = False
    End If
End Sub

Private Sub cmdexit_Click()
    Tmr1.Enabled = False
    Tmr2.Enabled = False
    Unload FrmExp44
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Tmr1.Enabled = False
    Tmr2.Enabled = False
    Set FrmExp44 = Nothing
End Sub

Private Sub Tmr1_Timer()
Static x As Integer
    
    If (x > dBmpStrip1.Frames) Then
        x = 1
    End If

    dBmpStrip1.FrameIndex = x
    
    x = (x + 1)
End Sub

Private Sub Tmr2_Timer()
Static x As Integer
    
    If (x > dBmpStrip2.Frames) Then
        x = 1
    End If

    dBmpStrip2.FrameIndex = x
    
    x = (x + 1)
End Sub
