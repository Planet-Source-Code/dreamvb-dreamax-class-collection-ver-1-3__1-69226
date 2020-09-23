VERSION 5.00
Begin VB.Form FrmExp42 
   Caption         =   "dCheckPanel"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAlign 
      Caption         =   "Align Right"
      Height          =   340
      Left            =   3330
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin Exp42.dCheckPanel dColors 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   225
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   529
      Caption         =   "Colors Options"
      BackColor       =   4210752
      CaptionForeColor=   16777215
      CaptionBackColor=   8421504
      CaptionLineColor=   0
      CaptionAlignment=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   340
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label lblcount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#0"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   990
      Width           =   195
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "#1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3885
      TabIndex        =   2
      Top             =   150
      Width           =   1800
   End
End
Attribute VB_Name = "FrmExp42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAlign_Click()
    If (dColors.CheckAlignment = CLeft) Then
        dColors.CheckAlignment = CRight
        cmdAlign.Caption = "Align Left"
    Else
        dColors.CheckAlignment = CLeft
        cmdAlign.Caption = "Align Right"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload FrmExp42
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dColors_Change(Index As Integer, Key As Variant, Value As Integer)
    lblInfo.Caption = "Index: " & Index _
    & vbCrLf & "Key: " & Key _
    & vbCrLf & "Value: " & Value
End Sub

Private Sub Form_Load()
    dColors.AddCheck "Red", "Red_Key"
    dColors.AddCheck "Green", "G"
    dColors.AddCheck "Blue", "B", True
    dColors.AddCheck "Yellow", "Y"
    '
    dColors.CheckBold(1) = True
    dColors.CheckColor(1) = vbRed
    dColors.CheckColor(2) = vbGreen
    dColors.CheckColor(3) = vbBlue
    dColors.CheckColor(4) = vbYellow
    
    lblcount.Top = (dColors.Height + lblcount.Height) + 120
    lblcount.Caption = "CheckCount: " & dColors.CheckCount
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp42 = Nothing
End Sub

