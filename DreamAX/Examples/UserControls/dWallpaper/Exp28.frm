VERSION 5.00
Begin VB.Form FrmExp28 
   BackColor       =   &H000080FF&
   Caption         =   "Wall Paper ActiveX - Example"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin Exp28.dWallPaper dWallPaper3 
      Align           =   2  'Align Bottom
      Height          =   1050
      Left            =   0
      TabIndex        =   2
      Top             =   1410
      Width           =   5175
      _extentx        =   9128
      _extenty        =   1852
      image           =   "Exp28.frx":0000
      maskcolor       =   -2147483633
      backstyle       =   0
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   390
         Width           =   1155
      End
   End
   Begin Exp28.dWallPaper dWallPaper2 
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4170
      _extentx        =   7355
      _extenty        =   979
      image           =   "Exp28.frx":804A
      backstyle       =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wallpapers can also be Transparent"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   2565
   End
End
Attribute VB_Name = "FrmExp28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload FrmExp28
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp28 = Nothing
End Sub
