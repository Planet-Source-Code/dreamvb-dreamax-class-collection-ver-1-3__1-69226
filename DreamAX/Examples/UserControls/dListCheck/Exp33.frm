VERSION 5.00
Begin VB.Form FrmExp33 
   Caption         =   "dListCheck - ActiveX Example"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin Exp33.dListCheck dListCheck1 
      Height          =   1800
      Left            =   90
      TabIndex        =   2
      Top             =   105
      Width           =   4005
      _extentx        =   7064
      _extenty        =   3175
      headerbackcolor =   8388608
      headerforecolor =   16777215
      selectbackcolor =   14737632
      selectforecolor =   0
      forecolor       =   4210752
   End
   Begin VB.CommandButton mnuexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label lblIdx 
      Height          =   1215
      Left            =   4275
      TabIndex        =   1
      Top             =   270
      Width           =   1785
   End
End
Attribute VB_Name = "FrmExp33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dListCheck1_Change()
    dListCheck1_Click
End Sub

Private Sub dListCheck1_Click()
Dim s As String
Dim Idx As Long
    
    Idx = dListCheck1.ListIndex
        
    If dListCheck1.IsHeader(Idx) Then
        s = ""
    Else
        s = "IsChecked = " & dListCheck1.IsChecked(Idx)
    End If
        
    lblIdx.Caption = "Index = " & Idx _
    & vbCrLf & "IsHeader = " & dListCheck1.IsHeader(Idx) _
    & vbCrLf & s
    
    s = ""
    
End Sub

Private Sub Form_Load()
    With dListCheck1
        .AddItem "Network", , LHeader
        .AddItem "Hide My Computer Icon"
        .AddItem "Disable Network Crawing"
        .AddItem "Increase Internet Connections"
        .AddItem "Disable Anoymous logins"
        .AddItem "Internet Explorer", , LHeader
        .AddItem "Disable Automatic Updates"
        .AddItem "Disable Image Toolbar"
        .AddItem "Disable Go Button"
        .AddItem "Clear Cache on Shutdown"
        .AddItem "Disable JavaScript"
        .AddItem "Disable ActiveX Controls"
        .AddItem "Disable Automatic Install of IE7"
        .AddItem "Services", , LHeader
        .AddItem "Disable Error-Report Servive"
        .AddItem "Disable Auto-Updates Service"
        .AddItem "Disable Firewall Service"
        .AddItem "Disable Indexing Services"
        .AddItem "Windows", , LHeader
        .AddItem "Disable Startmenu"
        .AddItem "Clear Recent Documents"
        .AddItem "Clear Page Cache at Shutdown"
        
        .ListIndex = 4
        .IsChecked(5) = False
        .IsChecked(3) = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp33 = Nothing
End Sub

Private Sub mnuexit_Click()
    Unload FrmExp33
End Sub
