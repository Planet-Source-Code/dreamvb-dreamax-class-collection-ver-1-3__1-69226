VERSION 5.00
Begin VB.Form frmTips 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShow 
      Height          =   210
      Left            =   75
      TabIndex        =   4
      Top             =   3375
      Width           =   5625
   End
   Begin VB.PictureBox pTipHolder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   75
      ScaleHeight     =   3135
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   120
      Width           =   4395
      Begin VB.PictureBox pSidePanel 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   3105
         Left            =   15
         ScaleHeight     =   3105
         ScaleWidth      =   675
         TabIndex        =   5
         Top             =   15
         Width           =   675
         Begin VB.Image ImgIco 
            Height          =   330
            Left            =   165
            Top             =   210
            Width           =   300
         End
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   750
         TabIndex        =   7
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label lblHeader 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   765
         TabIndex        =   6
         Top             =   165
         Width           =   3540
      End
      Begin VB.Line lnTop 
         BorderColor     =   &H00808080&
         X1              =   690
         X2              =   1905
         Y1              =   645
         Y2              =   645
      End
   End
   Begin VB.CommandButton cmdButton 
      Height          =   375
      Index           =   2
      Left            =   4575
      TabIndex        =   2
      Top             =   2880
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Height          =   375
      Index           =   1
      Left            =   4575
      TabIndex        =   1
      Top             =   585
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Height          =   375
      Index           =   0
      Left            =   4575
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_IconShow As Boolean
Public m_BorderColor As OLE_COLOR
Public TipCollection As Collection
Public ChkVal As Integer
Private TipCnt As Integer

Private Sub chkShow_Click()
    ChkVal = chkShow.Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
Static TipCnt As Integer
On Error Resume Next

    Select Case Index
        Case 0
            'Show Next Tip
            If (TipCnt >= TipCollection.Count) Then TipCnt = 0
            TipCnt = TipCnt + 1
            lblDesc.Caption = TipCollection(TipCnt)
        Case 1
            'Shows Prev
            TipCnt = TipCnt - 1
            If (TipCnt <= 1) Then TipCnt = 1
            lblDesc.Caption = TipCollection(TipCnt)
        Case 2
            'Unload the Tips
            Set TipCollection = Nothing
            Unload frmTips
    End Select
End Sub

Private Sub Form_Activate()
    Call cmdButton_Click(0)
    DoEvents
End Sub

Private Sub Form_Load()
    Set frmTips.Icon = Nothing
    TipCnt = 1
End Sub

Private Sub Form_Resize()
    'Draws a small thin border around the tips area
    pTipHolder.Line (0, 0)-(pTipHolder.ScaleWidth - 8, pTipHolder.ScaleHeight - 8), m_BorderColor, B
    pTipHolder.Refresh
    lnTop.X2 = (pTipHolder.ScaleWidth - lnTop)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTips = Nothing
End Sub

