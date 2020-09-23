VERSION 5.00
Begin VB.Form frmNews 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Readme GUI Client"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9345
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   623
      TabIndex        =   15
      Top             =   4350
      Width           =   9345
      Begin VB.CommandButton cmdabout 
         Caption         =   "&About"
         Height          =   345
         Left            =   7125
         TabIndex        =   17
         Top             =   105
         Width           =   990
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   345
         Left            =   8205
         TabIndex        =   16
         Top             =   105
         Width           =   990
      End
      Begin VB.Line ln3D 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   20
         Y1              =   1
         Y2              =   1
      End
      Begin VB.Line ln3D 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   20
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox PicTab 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3105
      Index           =   0
      Left            =   105
      ScaleHeight     =   3105
      ScaleWidth      =   9015
      TabIndex        =   13
      Top             =   1185
      Width           =   9015
      Begin VB.Label lblIntroTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   14
         Top             =   60
         Width           =   210
      End
   End
   Begin VB.PictureBox PicTab 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3105
      Index           =   1
      Left            =   -105
      ScaleHeight     =   3105
      ScaleWidth      =   9015
      TabIndex        =   9
      Top             =   7125
      Width           =   9015
      Begin VB.ListBox LstAX 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         IntegralHeight  =   0   'False
         Left            =   135
         TabIndex        =   10
         Top             =   390
         Width           =   8775
      End
      Begin VB.Label lblInfo1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select an item form the list for Updates and information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   90
         Width           =   4590
      End
      Begin VB.Label lblInfo2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   135
         TabIndex        =   11
         Top             =   1965
         Width           =   6345
      End
   End
   Begin VB.PictureBox PicTab 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3105
      Index           =   2
      Left            =   990
      ScaleHeight     =   3105
      ScaleWidth      =   9015
      TabIndex        =   6
      Top             =   4485
      Width           =   9015
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmNews.frx":0000
         Top             =   525
         Width           =   8790
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please see below if you like to be part of this project."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   8
         Top             =   60
         Width           =   4395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   90
         X2              =   8280
         Y1              =   330
         Y2              =   330
      End
   End
   Begin VB.PictureBox pictop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   0
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   623
      TabIndex        =   0
      Top             =   0
      Width           =   9345
      Begin VB.ComboBox cboView 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7110
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   345
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   4545
         TabIndex        =   5
         Top             =   510
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6555
         TabIndex        =   2
         Top             =   420
         Width           =   480
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DreamAX+Class Collection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   885
         TabIndex        =   1
         Top             =   105
         Width           =   4605
      End
      Begin VB.Image ImgLogo 
         Height          =   645
         Left            =   120
         Picture         =   "frmNews.frx":0331
         Top             =   75
         Width           =   645
      End
   End
   Begin VB.Label lblTitle1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   645
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dIni As New dINIFile
Private TmpListCol As Collection

Function FixPath(lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Private Sub ArrangeTabs(Index)
Dim TabIdx As Integer

    For TabIdx = 0 To PicTab.Count - 1
        PicTab(TabIdx).Visible = False
    Next TabIdx
    
    PicTab(Index).Visible = True
    PicTab(Index).Top = 1185
    PicTab(Index).Left = 150

End Sub

Private Sub cboView_Click()

    lblTitle1.Caption = cboView.Text
    
    Select Case cboView.ListIndex
        Case 0
            Call ArrangeTabs(cboView.ListIndex)
        Case 1, 2
            Call ArrangeTabs(1)
            Call FillListBox(cboView.Text & ".", LstAX)
            lblTitle1.Caption = cboView.Text & " " & LstAX.ListCount & " Items found"
        Case 3
            Call ArrangeTabs(2)
    End Select

End Sub

Private Sub FillListBox(SerachName As String, TLstBox As ListBox)
Dim Selections As New Collection
Dim Item
Dim sPos As Integer
Dim SerLen As Integer

    'Get Serach name Length
    SerLen = Len(SerachName)
    
    'Destroy the collection.
    Set TmpListCol = New Collection
    'Clear the listbox.
    TLstBox.Clear
    'Get all the ini selections.
    Set Selections = dIni.GetSelectionNames
    'Loop though the collections.
    For Each Item In Selections
        sPos = InStr(1, Item, SerachName, vbTextCompare)
        '
        If (sPos > 0) Then
            'Add Item Name to the listbox and collection.
            TmpListCol.Add Left$(Item, SerLen)
            TLstBox.AddItem Mid$(Item, SerLen + 1)
        End If
    Next Item
    
    TLstBox.ListIndex = 0
    
    'Destroy the selection collection
    Set Selections = Nothing
    Item = vbNullString
    
End Sub

Private Sub cmdabout_Click()
    MsgBox "DreamAX+Class Collection" & vbCrLf & "GUI Readme Viewer V1.0", vbInformation, "About"
End Sub

Private Sub cmdExit_Click()
    Unload frmNews
End Sub

Private Sub Form_Load()
    'Update File to read
    dIni.FileName = FixPath(App.Path) & "data.txt"
    'Check the INI file was found.
    If Not (dIni.IniFound) Then
        MsgBox "File Was Not Found:" _
        & vbCrLf & dIni.FileName, vbCritical, "File Not Found"
        Unload frmNews
        Exit Sub
    Else
        'Show introduction Text.
        lblIntroTxt.Caption = "DreamAX is a collection of ActiveX and Class For Visual Basic 6, The aim of the project" _
        & vbCrLf & "was to try and keep good old Visual Basic 6 alive, with the new avail of Visual Basic .NET" _
        & vbCrLf & vbCrLf & "This project will aim at providing developers, with new controls and classes that can be easy" _
        & vbCrLf & "inserted into there applications. This collection of codes aims to offer developers to use" _
        & vbCrLf & "many controls from stranded controls to more modem day ones."
        'Fill the comboboxes
        cboView.AddItem "Introduction"
        cboView.AddItem "ActiveX"
        cboView.AddItem "Class"
        cboView.AddItem "Helping Out"
        cboView.ListIndex = 0
    End If
End Sub

Private Sub ShowInfo(Index As Integer, TSerachItem As String, TLabel As Label)
Dim sItemSelection As String
Dim Info1 As String
Dim Info2 As String
Dim Info3 As String

    'Get the Selection we need to read from the INI
    sItemSelection = TmpListCol(Index) & TSerachItem
    
    If Not (dIni.SelectionExists(sItemSelection)) Then
        MsgBox "No information for '" & sItemSelection & _
        "' was found.", vbExclamation, frmNews.Caption
        Exit Sub
    Else
        Info1 = Trim$(dIni.INIReadKey(sItemSelection, "Info1"))
        Info2 = Trim$(dIni.INIReadKey(sItemSelection, "Info2"))
        Info3 = Trim$(dIni.INIReadKey(sItemSelection, "Info3"))
        '
        
        If Len(Info2) = 0 Then
            Info2 = "No Changes"
        End If
        
        
        TLabel.Caption = "Control: " & Info1 _
        & vbCrLf & "Name: " & TSerachItem _
        & vbCrLf & "Update: " & Info2 _
        & vbCrLf & "Description: " & Info3

    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set TmpListCol = Nothing
    Set frmNews = Nothing
End Sub

Private Sub LstAX_Click()
    Call ShowInfo(LstAX.ListIndex + 1, LstAX.Text, lblInfo2)
End Sub

Private Sub pBottom_Resize()
    ln3D(0).X2 = pBottom.ScaleWidth
    ln3D(1).X2 = pBottom.ScaleWidth
End Sub

Private Sub pictop_Resize()
    pictop.Line (0, pictop.ScaleHeight - 2)-(pictop.ScaleWidth, pictop.ScaleHeight - 2), &HC0C0C0
End Sub
