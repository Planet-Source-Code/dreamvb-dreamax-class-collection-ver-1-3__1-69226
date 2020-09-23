VERSION 5.00
Begin VB.Form FrmExp24 
   Caption         =   "Statusbar - ActiveX Example"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboGripStyle 
      Height          =   315
      Left            =   1785
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   585
      Width           =   1365
   End
   Begin Exp24.dStatusbar dStatusbar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   1545
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   585
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4815
      Top             =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4050
      TabIndex        =   0
      Top             =   915
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grip Styles:"
      Height          =   195
      Left            =   1815
      TabIndex        =   4
      Top             =   360
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Styles:"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   360
      Width           =   465
   End
End
Attribute VB_Name = "FrmExp24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboGripStyle_Click()
    dStatusbar1.GripStyle = cboGripStyle.ListIndex
End Sub

Private Sub cboStyle_Click()
    dStatusbar1.BarStyle = cboStyle.ListIndex
End Sub

Private Sub cmdExit_Click()
    Unload FrmExp24
End Sub

Private Sub Form_Load()
    'Styles
    cboStyle.AddItem "bNone"
    cboStyle.AddItem "bRaised"
    cboStyle.AddItem "bLowered"
    cboStyle.AddItem "bFrame"
    cboStyle.ListIndex = 2
    'Grip style
    cboGripStyle.AddItem "Default"
    cboGripStyle.AddItem "gsNewLook"
    cboGripStyle.ListIndex = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp24 = Nothing
End Sub

Private Sub Timer1_Timer()
    dStatusbar1.SimpleText = "Simple Statusbar | " & "Date " & Date & " | Time " & Time
End Sub
