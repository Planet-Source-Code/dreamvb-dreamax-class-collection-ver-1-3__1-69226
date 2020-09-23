VERSION 5.00
Begin VB.Form FrnEx5 
   Caption         =   "dMouseHover - Example"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1170
      ScaleHeight     =   435
      ScaleWidth      =   3345
      TabIndex        =   1
      Top             =   615
      Width           =   3405
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1695
      Width           =   1215
   End
End
Attribute VB_Name = "FrnEx5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oButton1 As dMouseHover
Attribute oButton1.VB_VarHelpID = -1
Private WithEvents oPicBox As dMouseHover
Attribute oPicBox.VB_VarHelpID = -1

Private Sub oButton1_MouseEnter()
    cmdExit.FontBold = True
    cmdExit.BackColor = &HE0E0E0
End Sub

Private Sub oButton1_MouseLeave()
    cmdExit.FontBold = False
    cmdExit.BackColor = vbButtonFace
End Sub

Private Sub oPicBox_MouseEnter()
    Picture1.Cls
    Picture1.ForeColor = vbYellow
    Picture1.BackColor = vbRed
    Picture1.CurrentX = 100
    Picture1.CurrentY = (Picture1.ScaleHeight - Picture1.TextHeight("Aa")) \ 2
    Picture1.Print "MouseEnter"
End Sub

Private Sub oPicBox_MouseLeave()
    Picture1.Cls
    Picture1.BackColor = vbButtonFace
    Picture1.ForeColor = vbBlack
    Picture1.CurrentX = 100
    Picture1.CurrentY = (Picture1.ScaleHeight - Picture1.TextHeight("Aa")) \ 2
    Picture1.Print "MouseLeave"
End Sub

Private Sub Form_Load()
    'Command button.
    Set oButton1 = New dMouseHover
    oButton1.AttachControl cmdExit
    'Picturebox.
    Set oPicBox = New dMouseHover
    oPicBox.AttachControl Picture1
End Sub

Private Sub cmdExit_Click()
    Set oButton1 = Nothing
    Set oPicBox = Nothing
    Unload FrnEx5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrnEx5 = Nothing
End Sub

