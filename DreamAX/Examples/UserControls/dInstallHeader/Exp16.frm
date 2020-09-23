VERSION 5.00
Begin VB.Form FrmExp16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Install Header - ActiveX Example"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdback 
      Caption         =   "< Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2580
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1275
      Width           =   1000
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   3675
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1275
      Width           =   1000
   End
   Begin Exp16.dInstallheader dInstallheader1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _extentx        =   10610
      _extenty        =   1588
      captionfont     =   "Exp16.frx":0000
      msgfont         =   "Exp16.frx":002E
      picture         =   "Exp16.frx":005C
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4830
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1275
      Width           =   1000
   End
End
Attribute VB_Name = "FrmExp16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaptions As Collection
Dim PageIdx As Integer

Private Sub Page(idx As Integer)
    With dInstallheader1
        Select Case idx
            Case 0
                .Caption = mCaptions("LicA")
                .Message = mCaptions("LicB")
            Case 1
                .Caption = mCaptions("InstallA")
                .Message = mCaptions("InstallB")
            Case 2
                .Caption = mCaptions("StartA")
                .Message = mCaptions("StartB")
        End Select
    End With
End Sub

Private Sub cmdback_Click()
    PageIdx = PageIdx - 1
    If (PageIdx <= 0) Then
        PageIdx = 0
        cmdback.Enabled = False
        cmdNext.Enabled = True
    End If
    
    Call Page(PageIdx)
End Sub

Private Sub cmdExit_Click()
    Unload FrmExp16
End Sub

Private Sub cmdNext_Click()
    PageIdx = PageIdx + 1
    If (PageIdx > 0) Then cmdback.Enabled = True
    
    If (PageIdx >= 2) Then
        PageIdx = 2
        cmdNext.Enabled = False
    End If
    
    Call Page(PageIdx)
End Sub

Private Sub Form_Load()
    Set mCaptions = New Collection
    
    mCaptions.Add "License Agreement", "LicA"
    mCaptions.Add "Please read the license before installing this software", "LicB"
    
    mCaptions.Add "Select Install Location", "InstallA"
    mCaptions.Add "Where should " & App.EXEName & " be installed?", "InstallB"

    mCaptions.Add "Select Start Menu Folder", "StartA"
    mCaptions.Add "Where should setup place the program's shortcuts?", "StartB"
    
    Call Page(0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp16 = Nothing
End Sub
