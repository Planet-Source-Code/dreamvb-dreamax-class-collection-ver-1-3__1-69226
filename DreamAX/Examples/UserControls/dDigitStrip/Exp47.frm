VERSION 5.00
Begin VB.Form FrmExp47 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "dDigitStrip ActiveX - Example"
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
   Begin Exp47.dDigitStrip dDigitStrip3 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   870
      Width           =   1170
      _extentx        =   2064
      _extenty        =   556
      value           =   256
      digitcolor      =   65535
      digitdimmedcolor=   4194304
      autosize        =   0   'False
      borderstyle     =   0
   End
   Begin Exp47.dDigitStrip dDigitStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   2220
      _extentx        =   3916
      _extenty        =   661
      value           =   123456789
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3855
      TabIndex        =   0
      Top             =   960
      Width           =   885
   End
   Begin Exp47.dDigitStrip dDigitStrip2 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   1920
      _extentx        =   3387
      _extenty        =   556
      value           =   11110000
      backcolor       =   16777215
      digitcolor      =   0
      digitdimmedcolor=   14737632
      borderstyle     =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supports Autosize"
      Height          =   195
      Left            =   1410
      TabIndex        =   6
      Top             =   945
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "< Different colors"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   585
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "< Borderstyle bFixedSingle"
      Height          =   195
      Left            =   2505
      TabIndex        =   4
      Top             =   150
      Width           =   1860
   End
End
Attribute VB_Name = "FrmExp47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload FrmExp47
End Sub

Private Sub Command1_Click()
    dDigitStrip1.Value = 256
    dDigitStrip1.AutoSize = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp47 = Nothing
End Sub

Private Sub Timer1_Timer()
Dim lVal As Long
    lVal = dDigitStrip1.Value + 1
    
    dDigitStrip1.Value = lVal
End Sub
