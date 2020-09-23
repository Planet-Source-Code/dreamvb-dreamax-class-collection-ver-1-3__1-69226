VERSION 5.00
Begin VB.Form FrmExp21 
   Caption         =   "ProgressBar ActiveX - Example"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Show Value"
      Height          =   255
      Left            =   195
      TabIndex        =   2
      Top             =   555
      Width           =   1455
   End
   Begin Exp21.dProgressBar dProgressBar1 
      Height          =   270
      Left            =   195
      Top             =   90
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   476
      BarColor        =   8438015
   End
   Begin VB.Timer Tmr1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3630
      Top             =   330
   End
   Begin VB.CommandButton cmdBut 
      Caption         =   "Start"
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   975
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1485
      TabIndex        =   0
      Top             =   975
      Width           =   1215
   End
End
Attribute VB_Name = "FrmExp21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    dProgressBar1.ShowValue = Check1
End Sub

Private Sub cmdBut_Click()
    If (cmdBut.Caption = "Start") Then
        cmdBut.Caption = "Stop"
        Tmr1.Enabled = True
    Else
        cmdBut.Caption = "Start"
        Tmr1.Enabled = False
    End If
End Sub

Private Sub cmdexit_Click()
    If (Tmr1.Enabled) Then Tmr1.Enabled = False
    Unload FrmExp21
End Sub

Private Sub dProgressBar1_ProgressChange()
    If (dProgressBar1.Value >= dProgressBar1.Max) Then dProgressBar1.Value = 0
End Sub

Private Sub Form_Load()
    dProgressBar1.Value = 0
    dProgressBar1.Max = 100
    dProgressBar1.Min = 0
    Check1.Value = Abs(dProgressBar1.ShowValue)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp21 = Nothing
End Sub

Private Sub Tmr1_Timer()
    dProgressBar1.Value = dProgressBar1.Value + 1
End Sub
