VERSION 5.00
Begin VB.Form FrmExp43 
   Caption         =   "XorCrypt - ActiveX Example"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKey 
      Height          =   360
      Left            =   465
      TabIndex        =   6
      Text            =   "ABC"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "Decrypt"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1515
      TabIndex        =   4
      Top             =   1665
      Width           =   900
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   465
      TabIndex        =   3
      Top             =   1665
      Width           =   900
   End
   Begin VB.TextBox TxtSrc 
      Height          =   360
      Left            =   465
      TabIndex        =   1
      Text            =   "Secret Message"
      Top             =   1110
      Width           =   3120
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2580
      TabIndex        =   0
      Top             =   1665
      Width           =   900
   End
   Begin Exp43.dXorCrypt dXorCrypt1 
      Left            =   2085
      Top             =   345
      _ExtentX        =   794
      _ExtentY        =   714
      Key             =   "Secret"
   End
   Begin VB.Label lblTitle1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      Height          =   195
      Left            =   465
      TabIndex        =   5
      Top             =   150
      Width           =   270
   End
   Begin VB.Label lblTitle2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Text"
      Height          =   195
      Left            =   465
      TabIndex        =   2
      Top             =   855
      Width           =   885
   End
End
Attribute VB_Name = "FrmExp43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sBuffer As String

Private Sub Cmd1_Click()
    Cmd1.Enabled = False
    Cmd2.Enabled = True
    'Encrypt
    dXorCrypt1.Key = TxtKey.Text
    sBuffer = TxtSrc.Text
    sBuffer = dXorCrypt1.XorCrypt(sBuffer)
    TxtSrc.Text = sBuffer
End Sub

Private Sub Cmd2_Click()
    Cmd1.Enabled = True
    Cmd2.Enabled = False
    'Decrypt
    dXorCrypt1.Key = TxtKey.Text
    sBuffer = dXorCrypt1.XorCrypt(sBuffer)
    TxtSrc.Text = sBuffer
    sBuffer = vbNullString
End Sub

Private Sub cmdexit_Click()
    sBuffer = vbNullString
    Unload FrmExp43
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp43 = Nothing
End Sub
