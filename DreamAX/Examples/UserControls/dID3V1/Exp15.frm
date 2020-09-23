VERSION 5.00
Begin VB.Form FrmExp15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ID3v1 ActiveX Reader - Example"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Exp15.dID3v1 dID3v11 
      Left            =   1935
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   350
      Left            =   165
      TabIndex        =   7
      Top             =   3390
      Width           =   980
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID3v1"
      Height          =   3105
      Left            =   90
      TabIndex        =   11
      Top             =   195
      Width           =   5160
      Begin VB.TextBox txtInfo 
         Height          =   345
         Index           =   0
         Left            =   930
         MaxLength       =   30
         TabIndex        =   1
         Top             =   690
         Width           =   4080
      End
      Begin VB.TextBox txtInfo 
         Height          =   345
         Index           =   1
         Left            =   930
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1110
         Width           =   4080
      End
      Begin VB.TextBox txtInfo 
         Height          =   345
         Index           =   2
         Left            =   930
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1530
         Width           =   4080
      End
      Begin VB.TextBox txtInfo 
         Height          =   345
         Index           =   3
         Left            =   930
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1935
         Width           =   780
      End
      Begin VB.TextBox txtInfo 
         Height          =   345
         Index           =   4
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2505
         Width           =   3870
      End
      Begin VB.ComboBox cboGen 
         Height          =   315
         Left            =   2475
         TabIndex        =   5
         Top             =   1980
         Width           =   2565
      End
      Begin VB.TextBox txtInfo 
         Height          =   345
         Index           =   5
         Left            =   4590
         MaxLength       =   3
         TabIndex        =   0
         Top             =   225
         Width           =   435
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Has Tag"
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   18
         Top             =   765
         Width           =   345
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         Caption         =   "Artist:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   17
         Top             =   1185
         Width           =   390
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         Caption         =   "Album:"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   16
         Top             =   1605
         Width           =   480
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         Caption         =   "Year:"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   15
         Top             =   2010
         Width           =   375
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   14
         Top             =   2565
         Width           =   780
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         Caption         =   "Genre"
         Height          =   195
         Index           =   5
         Left            =   1860
         TabIndex        =   13
         Top             =   2025
         Width           =   435
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         Caption         =   "# Track:"
         Height          =   195
         Index           =   6
         Left            =   3915
         TabIndex        =   12
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   3495
      TabIndex        =   10
      Top             =   3390
      Width           =   980
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   350
      Left            =   1260
      TabIndex        =   8
      Top             =   3390
      Width           =   980
   End
   Begin VB.CommandButton cmdundo 
      Caption         =   "&Undo"
      Height          =   350
      Left            =   2355
      TabIndex        =   9
      Top             =   3390
      Width           =   980
   End
End
Attribute VB_Name = "FrmExp15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Filename As String
Private cboTmp As String
Private c_idx As Integer

Private Function FixTrack(Num As Integer) As Integer
Dim f As Integer
    f = Num
    Do Until (f <= 256)
        f = f - 256
    Loop
    
    If (f = 256) Then f = 0
    
    FixTrack = f
    f = 0
End Function

Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Private Sub cboGen_Change()
    cboGen.Text = cboTmp
End Sub

Private Sub cboGen_Click()
    cboTmp = cboGen.Text
    c_idx = cboGen.ListIndex
End Sub

Private Sub cmdOpen_Click()
    'Open the MP3 for reading
    dID3v11.OpenMP3 (m_Filename)
    'Make sure the MP3 File is open
    If Not dID3v11.IsOpen Then
        MsgBox m_Filename & vbCrLf & "Was Not Found.", vbExclamation, "File Not Found"
        Exit Sub
    End If
    'Display to the user if the MP3 has a Info TAG
    lblTag.Caption = "Has Tag " & dID3v11.HasTag
    'Display MP3 Tag Information.
    Call DisplayInfo
End Sub

Private Sub cmdundo_Click()
    Call DisplayInfo
End Sub

Private Sub cmdUpdate_Click()
    
    If (dID3v11.IsOpen <> True) Then
        MsgBox "You need to first open the file before updateing.", vbExclamation, "File Not Open"
        Exit Sub
    Else
        'Update the MP3 Info
        txtInfo(5).Text = FixTrack(txtInfo(5).Text)
        With dID3v11
            .Title = txtInfo(0).Text
            .Artist = txtInfo(1).Text
            .Album = txtInfo(2).Text
            .mYear = txtInfo(3).Text
            .Comment = txtInfo(4).Text
            .Track = txtInfo(5).Text
            .Genre = c_idx
            Call .UpdateMP3
        End With
    End If
    
    'Show MP3 Tag Information
    Call DisplayInfo
End Sub

Private Sub Command2_Click()
    'Close the MP3 file if it's open
    If (dID3v11.IsOpen) Then
        dID3v11.CloseMP3
    End If
    
    Unload FrmExp15
End Sub

Private Sub DisplayInfo()
On Error Resume Next
    'Show the MP3 Info
    With dID3v11
        txtInfo(0).Text = .Title
        txtInfo(1).Text = .Artist
        txtInfo(2).Text = .Album
        txtInfo(3).Text = .mYear
        txtInfo(4).Text = .Comment
        txtInfo(5).Text = .Track
        cboGen.ListIndex = .Genre
    End With
End Sub

Private Sub Form_Load()
Dim x As Integer
    'Get MP3 Filename
    m_Filename = Environ("programfiles") & "\Winamp\demo.mp3"
    'init Genues
    For x = 0 To dID3v11.GenresCount
        cboGen.AddItem dID3v11.Genres(x)
    Next x
    cboGen.ListIndex = 0
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Index = 3) Or (Index = 5) Then
        If Not ((KeyAscii >= 48) And (KeyAscii <= 57) Or (KeyAscii = 8)) Then KeyAscii = 0
    End If
End Sub

