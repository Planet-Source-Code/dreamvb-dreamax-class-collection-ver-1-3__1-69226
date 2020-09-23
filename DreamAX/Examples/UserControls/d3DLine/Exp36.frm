VERSION 5.00
Begin VB.Form FrmExp36 
   Caption         =   "d3DLine ActiveX - Example"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin Exp36.d3DLine d3DLine4 
      Height          =   1380
      Left            =   1785
      TabIndex        =   3
      Top             =   540
      Width           =   150
      _extentx        =   265
      _extenty        =   2434
      direction       =   1
   End
   Begin Exp36.d3DLine d3DLine2 
      Height          =   90
      Left            =   345
      TabIndex        =   2
      Top             =   465
      Width           =   1215
      _extentx        =   2143
      _extenty        =   159
      color1          =   12648447
      color2          =   32896
   End
   Begin Exp36.d3DLine d3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _extentx        =   2143
      _extenty        =   53
   End
   Begin VB.CommandButton mnuexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3270
      TabIndex        =   0
      Top             =   1515
      Width           =   900
   End
End
Attribute VB_Name = "FrmExp36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    d3DLine1.Width = FrmExp36.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp36 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp36
End Sub
