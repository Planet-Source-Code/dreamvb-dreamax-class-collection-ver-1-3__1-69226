VERSION 5.00
Begin VB.Form FrmExp39 
   BackColor       =   &H00FFE3B5&
   Caption         =   "dCtrColorSet ActiveX - Example"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   ForeColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Left            =   3135
      TabIndex        =   11
      Top             =   2445
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3135
      TabIndex        =   10
      Top             =   2115
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   3135
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1575
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3135
      TabIndex        =   8
      Top             =   495
      Width           =   1590
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   150
      TabIndex        =   7
      Top             =   2070
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   150
      TabIndex        =   5
      Top             =   1455
      Width           =   1215
   End
   Begin VB.ListBox LstItems 
      Height          =   1230
      Left            =   1560
      TabIndex        =   4
      Top             =   945
      Width           =   1215
   End
   Begin VB.ComboBox cboItems 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   555
      Width           =   1215
   End
   Begin Exp39.dCtrlColorSet dCtrlColorSet1 
      Left            =   225
      Top             =   2745
      _extentx        =   794
      _extenty        =   714
      ctrlbackcolor   =   12632319
      ctrlforecolor   =   -2147483630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   4845
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2745
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Set BackColor and ForeColor for all controls at once"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   690
      TabIndex        =   2
      Top             =   105
      Width           =   4455
   End
End
Attribute VB_Name = "FrmExp39"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload FrmExp39
End Sub

Private Sub Form_Load()
Dim x As Integer

    'Note for command buttons to work set the style to 1-Graphical
    With dCtrlColorSet1
        'Set Backcolor of form for all controls
        .CtrlBackColor = FrmExp39.BackColor
        'Set Forecolor of form for all controls
        .CtrlForeColor = FrmExp39.ForeColor
        'Apply color settings above
        .Activate
    End With
    
    'Fill cobo and listbox with some examples.
    cboItems.AddItem "[Select]"
    cboItems.AddItem "Red"
    cboItems.AddItem "Blue"
    cboItems.AddItem "Green"
    cboItems.ListIndex = 0
    
    For x = 0 To 10
        LstItems.AddItem "Item : " & x
    Next x
    
    x = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp39 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp39
End Sub
