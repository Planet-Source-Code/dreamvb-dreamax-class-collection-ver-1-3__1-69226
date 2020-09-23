VERSION 5.00
Begin VB.Form FrmExp17 
   Caption         =   "Install Label"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Begin Exp17.dInstallDesc dInstallDesc1 
      Height          =   180
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      BlockColor      =   128
      AutoFit         =   -1  'True
      HintColor       =   8388608
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   1035
      TabIndex        =   1
      Top             =   2700
      Width           =   885
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next >"
      Height          =   360
      Left            =   30
      TabIndex        =   0
      Top             =   2700
      Width           =   885
   End
End
Attribute VB_Name = "FrmExp17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DescList As Collection

Private Sub cmdexit_Click()
    Set DescList = Nothing
    Unload FrmExp17
End Sub

Private Sub cmdnext_Click()
    dInstallDesc1.NextItem
End Sub

Private Sub Form_Load()
    Set DescList = New Collection
    With DescList
        'Add some items
        .Add "Install Windows"
        .Add "Checking Hardware"
        .Add "Copying Startup Files"
        .Add "Installing Windows"
        .Add "Setting up Start Menu"
        .Add "Finished"
        'Set the items
        dInstallDesc1.SetItems DescList
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp17 = Nothing
End Sub
