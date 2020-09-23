VERSION 5.00
Begin VB.Form FrmExp32 
   Caption         =   "dCDMenuList ActiveX - Example"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin Exp32.dCDMenuList dCDMenuList1 
      Height          =   2610
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   4604
   End
   Begin VB.CommandButton mnuexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3765
      TabIndex        =   0
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label LblInfo 
      Caption         =   "#0"
      Height          =   780
      Left            =   3720
      TabIndex        =   2
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "FrmExp32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FixPath(lPtah As String) As String
    If Right(lPtah, 1) = "\" Then
        FixPath = lPtah
    Else
        FixPath = lPtah & "\"
    End If
End Function

Private Sub dCDMenuList1_Change()
    LblInfo.Caption = "Index: " & dCDMenuList1.ListIndex _
    & vbCrLf & "Key: " & dCDMenuList1.ItemKey(dCDMenuList1.ListIndex)
End Sub

Private Sub Form_Load()
Dim sPath As String
Dim X As Integer

    sPath = FixPath(App.Path) & "icons\"
    
    For X = 0 To 5
        dCDMenuList1.AddItem "Test Menu Caption " & X, "Menu description " & X _
        , "KEY" & X + 1, LoadPicture(sPath & "home.bmp")
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp32 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp32
End Sub
