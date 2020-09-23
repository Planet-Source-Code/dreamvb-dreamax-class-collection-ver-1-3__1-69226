VERSION 5.00
Begin VB.Form FrmExp5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ChangeIcon Dialog ActiveX - Example"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Exp5.dChangeIconDialog dChangeIconDialog1 
      Left            =   2865
      Top             =   1080
      _ExtentX        =   794
      _ExtentY        =   714
   End
   Begin VB.TextBox txtIndex 
      Height          =   330
      Left            =   915
      TabIndex        =   5
      Text            =   "4"
      Top             =   480
      Width           =   570
   End
   Begin VB.TextBox TxtFile 
      Height          =   330
      Left            =   915
      TabIndex        =   3
      Top             =   90
      Width           =   3255
   End
   Begin VB.CommandButton cmdShowDLG 
      Caption         =   "Show Dialog"
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1500
      TabIndex        =   0
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   390
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   675
   End
End
Attribute VB_Name = "FrmExp5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload FrmExp5
End Sub

Private Sub cmdShowDLG_Click()
    With dChangeIconDialog1
        'Set the starting Index
        .Filename = TxtFile.Text
        .IconIndex = Val(txtIndex.Text)
        .ShowDialog
        'Show message.
        MsgBox "Filename: " & .Filename _
        & vbCrLf & "Icon Index: " & .IconIndex
    End With
End Sub

Private Sub Form_Load()
    TxtFile.Text = Environ("windir") & "\explorer.exe"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp5 = Nothing
End Sub
