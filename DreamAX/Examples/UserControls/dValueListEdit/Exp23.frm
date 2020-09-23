VERSION 5.00
Begin VB.Form FrmExp23 
   Caption         =   "dValueListEdit - ActiveX Example"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOp 
      Caption         =   "Exit"
      Height          =   375
      Index           =   10
      Left            =   3465
      TabIndex        =   10
      Top             =   4005
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "SetFont Arial"
      Height          =   375
      Index           =   9
      Left            =   3465
      TabIndex        =   12
      Top             =   3645
      Width           =   1470
   End
   Begin Exp23.dValueListEdit dValueListEdit1 
      Align           =   3  'Align Left
      Height          =   4905
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   8652
      GridColor       =   14737632
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "SelectedColor"
      Height          =   375
      Index           =   8
      Left            =   3465
      TabIndex        =   8
      Top             =   3285
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "GridColor"
      Height          =   375
      Index           =   7
      Left            =   3465
      TabIndex        =   7
      Top             =   2925
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "ForeColor"
      Height          =   375
      Index           =   6
      Left            =   3465
      TabIndex        =   6
      Top             =   2565
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "BackColor"
      Height          =   375
      Index           =   5
      Left            =   3465
      TabIndex        =   5
      Top             =   2205
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Item Count"
      Height          =   375
      Index           =   4
      Left            =   3465
      TabIndex        =   4
      Top             =   1845
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Add New Value"
      Height          =   375
      Index           =   3
      Left            =   3465
      TabIndex        =   3
      Top             =   1485
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "BoldValue Caption"
      Height          =   375
      Index           =   2
      Left            =   3465
      TabIndex        =   2
      Top             =   1125
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Set Age 25"
      Height          =   375
      Index           =   1
      Left            =   3465
      TabIndex        =   1
      Top             =   765
      Width           =   1470
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Get FirstName"
      Height          =   375
      Index           =   0
      Left            =   3465
      TabIndex        =   0
      Top             =   405
      Width           =   1470
   End
   Begin VB.Label lblIdx 
      Caption         =   "Index: 1"
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   90
      Width           =   1455
   End
End
Attribute VB_Name = "FrmExp23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOp_Click(Index As Integer)
Static cCol1 As Integer
Static cCol2 As Integer
Static cCol3 As Integer
Static cCol4 As Integer
Dim NewFont As IFontDisp

    Select Case Index
        Case 0
            MsgBox dValueListEdit1.ItemValue("fname")
        Case 1
            dValueListEdit1.ItemValue("A1") = 25
        Case 2
            dValueListEdit1.BoldValueNames = (Not dValueListEdit1.BoldValueNames)
        Case 3
            dValueListEdit1.AddItem "Time", "", Time
        Case 4
            MsgBox "Item Count: " & dValueListEdit1.ItemCount
        Case 5
            If (cCol1 > 15) Then cCol1 = 0
            dValueListEdit1.BackColor = QBColor(cCol1)
            cCol1 = (cCol1 + 1)
        Case 6
            If (cCol2 > 15) Then cCol2 = 0
            dValueListEdit1.ForeColor = QBColor(cCol2)
            cCol2 = (cCol2 + 1)
        Case 7
            If (cCol3 > 15) Then cCol3 = 0
            dValueListEdit1.GridColor = QBColor(cCol3)
            cCol3 = (cCol3 + 1)
        Case 8
            If (cCol4 > 15) Then cCol4 = 0
            dValueListEdit1.SelectedColor = QBColor(cCol4)
            cCol4 = (cCol4 + 1)
        Case 9
            Set NewFont = New StdFont
            NewFont.Name = "Arial"
            Set dValueListEdit1.Font = NewFont
            Set NewFont = Nothing
        Case 10
            Unload FrmExp23
    End Select
    
End Sub

Private Sub dValueListEdit1_ItemClick(Index As Integer)
    lblIdx.Caption = "Index: " & Index
End Sub

Private Sub Form_Load()
    'Add some examples.
    dValueListEdit1.AddItem "ID", , "0001"
    dValueListEdit1.AddItem "Firstname", "fname", "David"
    dValueListEdit1.AddItem "Lastname", , "Jones"
    dValueListEdit1.AddItem "Age", "A1", 28
    dValueListEdit1.AddItem "Gender", , "Male"
    dValueListEdit1.AddItem "Address1", , "85 Rocky Hills"
    dValueListEdit1.AddItem "Income", , "$500"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp23 = Nothing
End Sub
