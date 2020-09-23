VERSION 5.00
Begin VB.Form FrmExp9 
   Caption         =   "DotNet Button ActiveX - Example"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin Exp9.dDotNetButton dFlatButton5 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1965
      Width           =   1185
      _extentx        =   2090
      _extenty        =   873
      backcolor       =   12640511
      font            =   "Exp9.frx":0000
      forecolor       =   33023
      caption         =   "FlatButton"
      showfocusrect   =   0
   End
   Begin Exp9.dDotNetButton dFlatButton2 
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1140
      _extentx        =   2011
      _extenty        =   609
      captionalignment=   0
      font            =   "Exp9.frx":0028
      caption         =   "Left"
   End
   Begin Exp9.dDotNetButton dFlatButton1 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   225
      Width           =   1485
      _extentx        =   2619
      _extenty        =   609
      font            =   "Exp9.frx":0054
      caption         =   "No Focus Rect"
      showfocusrect   =   0
   End
   Begin Exp9.dDotNetButton cmdExit 
      Height          =   360
      Left            =   2835
      TabIndex        =   0
      Top             =   2310
      Width           =   1185
      _extentx        =   2090
      _extenty        =   635
      font            =   "Exp9.frx":0080
      caption         =   "E&xit"
   End
   Begin Exp9.dDotNetButton dFlatButton3 
      Height          =   345
      Left            =   1545
      TabIndex        =   4
      Top             =   1080
      Width           =   1140
      _extentx        =   2011
      _extenty        =   609
      font            =   "Exp9.frx":00AC
      caption         =   "Center"
   End
   Begin Exp9.dDotNetButton dFlatButton4 
      Height          =   345
      Left            =   2850
      TabIndex        =   5
      Top             =   1080
      Width           =   1140
      _extentx        =   2011
      _extenty        =   609
      captionalignment=   2
      font            =   "Exp9.frx":00D8
      caption         =   "Right"
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
Attribute VB_Name = "FrmExp9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload FrmExp9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp9 = Nothing
End Sub
