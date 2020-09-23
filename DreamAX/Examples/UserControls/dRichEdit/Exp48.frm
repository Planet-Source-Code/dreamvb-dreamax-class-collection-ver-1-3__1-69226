VERSION 5.00
Begin VB.Form FrmExp48 
   Caption         =   "dRichEdit ActiveX - Example"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   340
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1080
   End
   Begin Exp48.dRichEdit dRichEdit1 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   4577
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Scrollbars      =   3
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   340
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1080
   End
End
Attribute VB_Name = "FrmExp48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrmExp48
End Sub

Private Sub Form_Load()
    DoEvents
    dRichEdit1.LoadFromFile App.Path & "\demo.rtf"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp48 = Nothing
End Sub

