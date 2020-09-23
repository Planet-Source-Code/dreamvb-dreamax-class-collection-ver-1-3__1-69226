VERSION 5.00
Begin VB.Form FrmEx7 
   Caption         =   "dTextValidate - Example"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   4620
      TabIndex        =   13
      Top             =   450
      Width           =   975
   End
   Begin VB.CheckBox ChkSpcKey 
      Caption         =   "Allow Space Key"
      Height          =   225
      Left            =   2280
      TabIndex        =   12
      Top             =   1350
      Width           =   2940
   End
   Begin VB.CheckBox ChkDelKey 
      Caption         =   "Allow Delete Key"
      Height          =   225
      Left            =   2280
      TabIndex        =   11
      Top             =   1020
      Value           =   1  'Checked
      Width           =   2745
   End
   Begin VB.TextBox txtCustom 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   2490
      Width           =   1815
   End
   Begin VB.TextBox txtCust 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2265
      TabIndex        =   7
      Text            =   "[A-Za-z:\]"
      Top             =   465
      Width           =   2295
   End
   Begin VB.TextBox txtAlpha 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   1815
   End
   Begin VB.TextBox txtAlphaNum 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1845
      Width           =   1815
   End
   Begin VB.TextBox txtNum 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   465
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   5040
      TabIndex        =   0
      Top             =   2460
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2265
      Width           =   525
   End
   Begin VB.Label lblCust 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Format"
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   225
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha and Numeric"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1605
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplha"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   870
      Width           =   405
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numeric"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   585
   End
End
Attribute VB_Name = "FrmEx7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private o_Text1 As New dTextValidate
Private o_Text2 As New dTextValidate
Private o_Text3 As New dTextValidate
Private o_Text4 As New dTextValidate

Private Sub DestroyCtrls()
    Set o_Text1 = Nothing
    Set o_Text2 = Nothing
    Set o_Text3 = Nothing
    Set o_Text4 = Nothing
End Sub

Private Sub SetupCtrls()
    'fNumeric
    o_Text1.TextBox = txtNum
    o_Text1.AllowDelete = ChkDelKey
    o_Text1.AllowSpace = ChkSpcKey
    o_Text1.Format = fNumeric
    'fNumeric
    o_Text2.TextBox = txtAlpha
    o_Text2.AllowDelete = ChkDelKey
    o_Text2.AllowSpace = ChkSpcKey
    o_Text2.Format = fAplha
    'fAlphaNumeric
    o_Text3.TextBox = txtAlphaNum
    o_Text3.AllowDelete = ChkDelKey
    o_Text3.AllowSpace = ChkSpcKey
    o_Text3.Format = fAlphaNumeric
    'fCustom
    o_Text4.TextBox = txtCustom
    o_Text4.AllowDelete = ChkDelKey
    o_Text4.AllowSpace = ChkSpcKey
    o_Text4.Format = fCustom
    o_Text4.CustomFormat = txtCust.Text
End Sub

Private Sub ChkDelKey_Click()
    Call SetupCtrls
End Sub

Private Sub ChkSpcKey_Click()
    Call SetupCtrls
End Sub

Private Sub cmdExit_Click()
    Call DestroyCtrls
    Unload FrmEx7
End Sub

Private Sub cmdUpdate_Click()
    Call SetupCtrls
End Sub

Private Sub Form_Load()
    Call SetupCtrls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmEx7 = Nothing
End Sub

