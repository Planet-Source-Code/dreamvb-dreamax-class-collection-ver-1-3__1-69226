VERSION 5.00
Begin VB.Form FrmExp38 
   Caption         =   "dTabStrip ActiveX - Example"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin Exp38.dTabStrip dTabStrip1 
      Height          =   345
      Left            =   540
      TabIndex        =   3
      Top             =   315
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Make Selection Bold"
      Height          =   270
      Left            =   375
      TabIndex        =   2
      Top             =   1605
      Width           =   1890
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   3330
      TabIndex        =   0
      Top             =   2190
      Width           =   1005
   End
   Begin VB.Label lblDis 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   720
      Left            =   375
      TabIndex        =   1
      Top             =   900
      Width           =   1920
   End
End
Attribute VB_Name = "FrmExp38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBold_Click()
    dTabStrip1.BoldSelected = chkBold
End Sub

Private Sub cmdExit_Click()
    Unload FrmExp38
End Sub

Private Sub dTabStrip1_TabChange(Index As Integer, Key As String, Caption As String)
    'Here we find out what tab was pressed and other tab info.
    lblDis.Caption = "Index :" & Index _
    & vbCrLf & "Caption :" & Caption _
    & vbCrLf & "Key :" & Key
End Sub

Private Sub Form_Load()
    With dTabStrip1
        'Add some Text Items
        .AddTab "Games", "G"
        .AddTab "Software", "S"
        .AddTab "Test", "Test"
        'Select the second tab
        .TabSelect = 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp38 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp38
End Sub
