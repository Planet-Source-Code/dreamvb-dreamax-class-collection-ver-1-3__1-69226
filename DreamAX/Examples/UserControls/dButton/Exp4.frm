VERSION 5.00
Begin VB.Form FrmExp4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Button ActiveX - Example"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   390
      Left            =   3285
      TabIndex        =   9
      Top             =   2565
      Width           =   1215
   End
   Begin Exp4.dmButton dmButton6 
      Height          =   375
      Left            =   855
      TabIndex        =   7
      Top             =   2400
      Width           =   1155
      _extentx        =   2037
      _extenty        =   661
      picture         =   "Exp4.frx":0000
      font            =   "Exp4.frx":0352
   End
   Begin Exp4.dmButton dmButton4 
      Height          =   480
      Left            =   165
      TabIndex        =   3
      Top             =   1470
      Width           =   1155
      _extentx        =   2037
      _extenty        =   847
      forecolor       =   255
      font            =   "Exp4.frx":037E
   End
   Begin Exp4.dmButton dmButton3 
      Height          =   375
      Left            =   2955
      TabIndex        =   2
      Top             =   405
      Width           =   1155
      _extentx        =   2037
      _extenty        =   661
      font            =   "Exp4.frx":03AA
   End
   Begin Exp4.dmButton dmButton2 
      Height          =   375
      Left            =   1530
      TabIndex        =   1
      Top             =   405
      Width           =   1155
      _extentx        =   2037
      _extenty        =   661
      font            =   "Exp4.frx":03D6
   End
   Begin Exp4.dmButton dmButton1 
      Height          =   375
      Left            =   210
      TabIndex        =   0
      Top             =   405
      Width           =   1155
      _extentx        =   2037
      _extenty        =   661
      font            =   "Exp4.frx":0402
   End
   Begin Exp4.dmButton dmButton5 
      Height          =   480
      Left            =   1440
      TabIndex        =   6
      Top             =   1470
      Width           =   1155
      _extentx        =   2037
      _extenty        =   847
      forecolor       =   192
      font            =   "Exp4.frx":042E
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "With Picture"
      Height          =   195
      Left            =   330
      TabIndex        =   8
      Top             =   2130
      Width           =   870
   End
   Begin VB.Label lblstyle 
      BackStyle       =   0  'Transparent
      Caption         =   "Font Styles, Colors"
      Height          =   165
      Left            =   255
      TabIndex        =   5
      Top             =   1020
      Width           =   2475
   End
   Begin VB.Label lblAlign 
      AutoSize        =   -1  'True
      Caption         =   "Alignments"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   75
      Width           =   765
   End
End
Attribute VB_Name = "FrmExp4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrmExp4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp4 = Nothing
End Sub
