VERSION 5.00
Begin VB.Form FrmExp30 
   Caption         =   "dCheckbox ActiveX - Example"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin Exp30.dCheckBox dCheckBox4 
      Height          =   285
      Left            =   105
      TabIndex        =   9
      Top             =   1755
      Width           =   2865
      _extentx        =   5054
      _extenty        =   503
      caption         =   "Event Hover Out"
      forecolor       =   33023
      font            =   "Exp30.frx":0000
      style           =   0
      highlight       =   16576
      showhighlight   =   -1  'True
   End
   Begin Exp30.dCheckBox dCheckBox1 
      Height          =   300
      Left            =   105
      TabIndex        =   6
      Top             =   1335
      Width           =   1215
      _extentx        =   2143
      _extenty        =   529
      caption         =   "Old Style"
      font            =   "Exp30.frx":002C
      enabled         =   0   'False
      style           =   0
   End
   Begin Exp30.dCheckBox dXPCheckBox5 
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   930
      Width           =   3675
      _extentx        =   6482
      _extenty        =   476
      value           =   1
      caption         =   "Value of check box is:"
      backcolor       =   14737632
      forecolor       =   255
      font            =   "Exp30.frx":0058
   End
   Begin Exp30.dCheckBox dXPCheckBox1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   2070
      _extentx        =   3651
      _extenty        =   450
      caption         =   "Left Aligned"
      font            =   "Exp30.frx":0080
      showhighlight   =   -1  'True
   End
   Begin VB.CommandButton mnuexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2085
      Width           =   900
   End
   Begin Exp30.dCheckBox dXPCheckBox2 
      Height          =   255
      Left            =   2445
      TabIndex        =   2
      Top             =   195
      Width           =   2070
      _extentx        =   3651
      _extenty        =   450
      caption         =   "Right Aligned"
      alignment       =   2
      font            =   "Exp30.frx":00AC
   End
   Begin Exp30.dCheckBox dXPCheckBox3 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   555
      Width           =   2070
      _extentx        =   3651
      _extenty        =   450
      value           =   1
      caption         =   "Check box is Checked"
      font            =   "Exp30.frx":00D8
   End
   Begin Exp30.dCheckBox dXPCheckBox4 
      Height          =   255
      Left            =   2460
      TabIndex        =   4
      Top             =   540
      Width           =   2070
      _extentx        =   3651
      _extenty        =   450
      caption         =   "Checkbox Disabled"
      alignment       =   2
      font            =   "Exp30.frx":0104
      enabled         =   0   'False
   End
   Begin Exp30.dCheckBox dCheckBox2 
      Height          =   300
      Left            =   1530
      TabIndex        =   7
      Top             =   1335
      Width           =   855
      _extentx        =   1508
      _extenty        =   529
      value           =   1
      caption         =   "Moden"
      font            =   "Exp30.frx":0130
      style           =   1
   End
   Begin Exp30.dCheckBox dCheckBox3 
      Height          =   375
      Left            =   2565
      TabIndex        =   8
      Top             =   1335
      Width           =   1395
      _extentx        =   2461
      _extenty        =   661
      caption         =   "Button"
      font            =   "Exp30.frx":015C
      style           =   3
   End
End
Attribute VB_Name = "FrmExp30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dCheckBox4_HoverIn()
    dCheckBox4.Caption = "Event Hover In"
End Sub

Private Sub dCheckBox4_HoverOut()
    dCheckBox4.Caption = "Event Hover Out"
End Sub

Private Sub dXPCheckBox5_Click()
    dXPCheckBox5.Caption = "Value of check box is: " & CBool(dXPCheckBox5.Value)
End Sub

Private Sub Form_Load()
    dXPCheckBox5_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp30 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp30
End Sub
