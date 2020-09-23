VERSION 5.00
Begin VB.Form FrmExp7 
   Caption         =   "Color Picker ActiveX -  Example"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicDraw 
      AutoRedraw      =   -1  'True
      Height          =   2700
      Left            =   1050
      ScaleHeight     =   176
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   2
      Top             =   135
      Width           =   3855
   End
   Begin Exp7.dColorPanel dColorPanel1 
      Height          =   3645
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   885
      _extentx        =   1561
      _extenty        =   6429
      openfullcolordlg=   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3765
      TabIndex        =   0
      Top             =   3150
      Width           =   1215
   End
End
Attribute VB_Name = "FrmExp7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OldX As Single
Dim OldY As Single
Dim CanMove As Boolean

Private Sub cmdExit_Click()
    Unload FrmExp7
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp7 = Nothing
End Sub

Private Sub PicDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    OldY = Y
    CanMove = True
End Sub

Private Sub PicDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim UseColor As OLE_COLOR
    
    If (CanMove) Then
        If (Button = vbLeftButton) Then UseColor = dColorPanel1.ForegroundColor
        If (Button = vbRightButton) Then UseColor = dColorPanel1.BackgroundColor
        PicDraw.Line (X, Y)-(OldX, OldY), UseColor
    End If
End Sub

Private Sub PicDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CanMove = False
End Sub
