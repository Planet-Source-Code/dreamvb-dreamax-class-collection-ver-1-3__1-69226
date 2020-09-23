VERSION 5.00
Begin VB.Form FrmExp12 
   Caption         =   "Flat Button ActiveX - Example"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin Exp12.dFlatButton dFlatButton5 
      Height          =   375
      Left            =   165
      TabIndex        =   7
      Top             =   1920
      Width           =   1185
      _extentx        =   2090
      _extenty        =   661
      backcolor       =   12640511
      font            =   "Exp12.frx":0000
      forecolor       =   33023
      showfocusrect   =   0   'False
   End
   Begin Exp12.dFlatButton dFlatButton2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1140
      _extentx        =   2011
      _extenty        =   661
      captionalignment=   0
      font            =   "Exp12.frx":0028
      caption         =   "Left"
   End
   Begin Exp12.dFlatButton dFlatButton1 
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Top             =   255
      Width           =   2220
      _extentx        =   3916
      _extenty        =   661
      font            =   "Exp12.frx":0054
      caption         =   "No Focus Rect"
      showfocusrect   =   0   'False
   End
   Begin Exp12.dFlatButton cmdExit 
      Height          =   375
      Left            =   2910
      TabIndex        =   0
      Top             =   2715
      Width           =   1185
      _extentx        =   2090
      _extenty        =   661
      font            =   "Exp12.frx":0080
      caption         =   "Exit"
   End
   Begin Exp12.dFlatButton dFlatButton3 
      Height          =   375
      Left            =   1545
      TabIndex        =   4
      Top             =   1080
      Width           =   1140
      _extentx        =   2011
      _extenty        =   661
      font            =   "Exp12.frx":00AC
      caption         =   "Center"
   End
   Begin Exp12.dFlatButton dFlatButton4 
      Height          =   375
      Left            =   2850
      TabIndex        =   5
      Top             =   1080
      Width           =   1140
      _extentx        =   2011
      _extenty        =   661
      captionalignment=   2
      font            =   "Exp12.frx":00D8
      caption         =   "Right"
   End
   Begin Exp12.dFlatButton dFlatButton6 
      Height          =   375
      Left            =   1485
      TabIndex        =   8
      Top             =   1935
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      font            =   "Exp12.frx":0104
      caption         =   "With Picture"
      picture         =   "Exp12.frx":0130
   End
   Begin Exp12.dFlatButton dFlatButton7 
      Height          =   375
      Left            =   3210
      TabIndex        =   9
      Top             =   1920
      Width           =   780
      _extentx        =   1376
      _extenty        =   661
      font            =   "Exp12.frx":0482
      caption         =   "Disabled"
      enabled         =   0   'False
   End
   Begin Exp12.dFlatButton dFlatButton8 
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   2385
      Width           =   1140
      _extentx        =   2011
      _extenty        =   661
      font            =   "Exp12.frx":04AE
      caption         =   "Old Button"
      buttonstyle     =   0
   End
   Begin Exp12.dFlatButton dFlatButton9 
      Height          =   375
      Left            =   1425
      TabIndex        =   11
      Top             =   2400
      Width           =   1140
      _extentx        =   2011
      _extenty        =   661
      font            =   "Exp12.frx":04DA
      caption         =   "Frame"
      buttonstyle     =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colors, Fonts"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alignments"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   750
      Width           =   765
   End
End
Attribute VB_Name = "FrmExp12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload FrmExp12
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp12 = Nothing
End Sub
