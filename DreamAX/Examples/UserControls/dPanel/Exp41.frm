VERSION 5.00
Begin VB.Form FrmExp41 
   Caption         =   "dPlanel ActiveX - Example"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin Exp41.dPanel dPanel1 
      Height          =   1620
      Left            =   2130
      TabIndex        =   7
      Top             =   330
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2858
      BevelOuter      =   2
      Caption         =   "Preview"
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
   Begin VB.ComboBox cboAlign 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1770
      Width           =   1215
   End
   Begin VB.ComboBox cboOut 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cboIn 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   390
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption Alignment:"
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   5
      Top             =   1530
      Width           =   1320
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outer Bevel:"
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   2
      Top             =   840
      Width           =   885
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inter Bevel:"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   1
      Top             =   150
      Width           =   810
   End
End
Attribute VB_Name = "FrmExp41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UpdateBevel()
    dPanel1.BevelInner = cboIn.ListIndex
    dPanel1.BevelOuter = cboOut.ListIndex
    dPanel1.CaptionAlign = cboAlign.ListIndex
End Sub

Private Sub cboAlign_Click()
    Call UpdateBevel
End Sub

Private Sub cboIn_Click()
    Call UpdateBevel
End Sub

Private Sub cboOut_Click()
    Call UpdateBevel
End Sub

Private Sub cmdExit_Click()
    Unload FrmExp41
End Sub

Private Sub Form_Load()
    cboIn.AddItem "bvNone"
    cboIn.AddItem "bvLowered"
    cboIn.AddItem "bvRaised"
    
    cboOut.AddItem "bvNone"
    cboOut.AddItem "bvLowered"
    cboOut.AddItem "bvRaised"
    
    cboAlign.AddItem "aLeft"
    cboAlign.AddItem "aCenter"
    cboAlign.AddItem "aRight"
    
    cboIn.ListIndex = 0
    cboOut.ListIndex = 2
    cboAlign.ListIndex = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp41 = Nothing
End Sub

