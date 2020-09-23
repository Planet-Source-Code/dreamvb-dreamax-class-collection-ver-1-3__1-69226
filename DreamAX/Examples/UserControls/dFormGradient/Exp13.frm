VERSION 5.00
Begin VB.Form FrmExp13 
   AutoRedraw      =   -1  'True
   Caption         =   "Gradient ActiveX - Example"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin Exp13.dFormGradient dFormGradient1 
      Left            =   90
      Top             =   75
      _ExtentX        =   794
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   3450
      TabIndex        =   0
      Top             =   1605
      Width           =   1080
   End
End
Attribute VB_Name = "FrmExp13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload FrmExp13
End Sub

Private Sub Form_Load()
    'We need to tell the control what form to use.
    dFormGradient1.FormObject = FrmExp13
    'Sets the Gradient Direction
    dFormGradient1.Direction = Horizontal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp13 = Nothing
End Sub
