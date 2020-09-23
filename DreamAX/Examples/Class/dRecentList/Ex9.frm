VERSION 5.00
Begin VB.Form FrmEx9 
   Caption         =   "dRecentList - Example"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Items"
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Top             =   2550
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Items"
      Height          =   375
      Left            =   1380
      TabIndex        =   8
      Top             =   2055
      Width           =   1185
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Items"
      Height          =   375
      Left            =   90
      TabIndex        =   7
      Top             =   2055
      Width           =   1185
   End
   Begin VB.CommandButton cmdSetMax 
      Caption         =   "Set Max"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1500
      Width           =   840
   End
   Begin VB.TextBox txtMax 
      Height          =   375
      Left            =   3900
      TabIndex        =   5
      Text            =   "3"
      Top             =   1500
      Width           =   555
   End
   Begin VB.TextBox txtItem 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "C:\Example.gif"
      Top             =   1500
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2010
      TabIndex        =   3
      Top             =   1500
      Width           =   825
   End
   Begin VB.ListBox LstItems 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3855
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   870
   End
   Begin VB.Label lblRecent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recent Items:"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   990
   End
End
Attribute VB_Name = "FrmEx9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents RecentLst As dRecentList
Attribute RecentLst.VB_VarHelpID = -1

Private Sub RefreshLst()
Dim x As Integer
    With RecentLst
        'load the items info listbox.
        LstItems.Clear
        For x = 1 To .MaxItems
            LstItems.AddItem .Item(x)
        Next x
    End With
End Sub

Private Sub cmdAdd_Click()
    'Adds a New Item
    RecentLst.AddItem txtItem.Text
    'Refresh.
    Call RefreshLst
End Sub

Private Sub cmdClear_Click()
    'Clears items and items that were saved to the Registry.
    RecentLst.Clear
    'Refresh.
    Call RefreshLst
End Sub

Private Sub cmdLoad_Click()
    'Load the Items
    RecentLst.LoadItems
    'Refresh.
    Call RefreshLst
End Sub

Private Sub cmdSave_Click()
    'Save the Items
    RecentLst.SaveItems
    'Refresh.
    Call RefreshLst
End Sub

Private Sub cmdSetMax_Click()
    RecentLst.MaxItems = Val(txtMax.Text)
    Call RefreshLst
End Sub

Private Sub Form_Load()
    'Create thw Class Object.
    Set RecentLst = New dRecentList
    'Set max items
    RecentLst.MaxItems = 3
    'Set Key for saveing Items
    RecentLst.AppName = "Test"
    'Add some Test Items
    RecentLst.AddItem "C:\MyLogo.bmp"
    RecentLst.AddItem "C:\Cats.bmp"
    RecentLst.AddItem "C:\House.bmp"
    'Refresh
    Call RefreshLst
End Sub

Private Sub RecentLst_Error()
    MsgBox Err.Description, vbInformation, "Error_" & Err.Number
End Sub

Private Sub cmdExit_Click()
    Unload FrmEx9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmEx9 = Nothing
End Sub

