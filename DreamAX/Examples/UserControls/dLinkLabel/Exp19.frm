VERSION 5.00
Begin VB.Form FrmExp19 
   Caption         =   "Link Label ActiveX - Example"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin Exp19.dLinkLabel dLinkLabel4 
      Height          =   345
      Left            =   3690
      TabIndex        =   4
      Top             =   1575
      Width           =   780
      _extentx        =   1376
      _extenty        =   609
      caption         =   "Exit"
      textcolor       =   4210752
      backcolor       =   14737632
      font            =   "Exp19.frx":0000
      showunderline   =   0   'False
      mousepointer    =   32649
   End
   Begin Exp19.dLinkLabel dLinkLabel3 
      Height          =   195
      Left            =   2175
      TabIndex        =   3
      Top             =   1035
      Width           =   1350
      _extentx        =   2381
      _extenty        =   344
      caption         =   "Event HoverOut"
      hovercolor      =   0
      activecolor     =   0
      font            =   "Exp19.frx":002C
      visitedcolor    =   0
   End
   Begin Exp19.dLinkLabel dLinkLabel2 
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   765
      Width           =   2145
      _extentx        =   3784
      _extenty        =   344
      caption         =   "Open My Windows Folder"
      font            =   "Exp19.frx":0054
   End
   Begin VB.CheckBox chkVis 
      Caption         =   "URL Visited"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   315
      Width           =   2160
   End
   Begin Exp19.dLinkLabel dLinkLabel1 
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   300
      Width           =   1020
      _extentx        =   1799
      _extenty        =   344
      caption         =   "Visit Google"
      url             =   "http://www.google.com"
      font            =   "Exp19.frx":007C
      mousepointer    =   32649
   End
End
Attribute VB_Name = "FrmExp19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkVis_Click()
    dLinkLabel1.IsVisited = chkVis.Value
End Sub

Private Sub dLinkLabel1_HoverOut()
    chkVis.Value = Abs(dLinkLabel1.IsVisited)
End Sub

Private Sub dLinkLabel2_HoverIn()
    dLinkLabel2.Font.Bold = True
End Sub

Private Sub dLinkLabel2_HoverOut()
    dLinkLabel2.Font.Bold = False
End Sub

Private Sub dLinkLabel2_MouseUp(shift As Integer, X As Single, Y As Single)
    dLinkLabel2.IsVisited = Not dLinkLabel2.IsVisited
End Sub

Private Sub dLinkLabel3_HoverIn()
    dLinkLabel3.Caption = "Event Hover In"
End Sub

Private Sub dLinkLabel3_HoverOut()
        dLinkLabel3.Caption = "Event Hover Out"
End Sub

Private Sub dLinkLabel4_MouseUp(shift As Integer, X As Single, Y As Single)
    Unload FrmExp19
End Sub

Private Sub Form_Load()
    dLinkLabel2.Url = Environ("WinDir")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp19 = Nothing
End Sub
