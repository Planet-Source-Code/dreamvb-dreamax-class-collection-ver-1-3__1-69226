VERSION 5.00
Begin VB.Form FrmExp26 
   Caption         =   "dColorPallete ActiveX - Example"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pview 
      Height          =   300
      Left            =   2430
      ScaleHeight     =   240
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   630
      Width           =   1395
   End
   Begin Exp26.dColorPallete dColorPallete1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3840
      _extentx        =   6773
      _extenty        =   847
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   2625
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1110
      Width           =   1215
   End
End
Attribute VB_Name = "FrmExp26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Unload FrmExp26
End Sub

Private Sub dColorPallete1_ItemClick(Button As MouseButtonConstants, ColorValue As stdole.OLE_COLOR)
    pview.BackColor = ColorValue
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp26 = Nothing
End Sub
