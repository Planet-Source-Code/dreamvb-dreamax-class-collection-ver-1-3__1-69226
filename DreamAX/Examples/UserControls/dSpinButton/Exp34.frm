VERSION 5.00
Begin VB.Form FrmExp34 
   Caption         =   "dSpinButton ActiveX - Example"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   525
      TabIndex        =   2
      Text            =   "0"
      Top             =   180
      Width           =   2850
   End
   Begin Exp34.dSpinButton dSpinButton1 
      Height          =   540
      Left            =   135
      TabIndex        =   1
      Top             =   180
      Width           =   360
      _extentx        =   635
      _extenty        =   953
      value           =   10
   End
   Begin VB.CommandButton mnuexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3735
      TabIndex        =   0
      Top             =   945
      Width           =   900
   End
End
Attribute VB_Name = "FrmExp34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dSpinButton1_Change()
    txtValue = dSpinButton1
End Sub

Private Sub Form_Load()
    dSpinButton1_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp34 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp34
End Sub
