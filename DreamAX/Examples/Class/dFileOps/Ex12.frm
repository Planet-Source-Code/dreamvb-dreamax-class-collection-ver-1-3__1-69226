VERSION 5.00
Begin VB.Form FrmEx12 
   Caption         =   "dFileOps - Example"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   660
      Width           =   6435
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   5490
      TabIndex        =   0
      Top             =   3675
      Width           =   990
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "#0"
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
      Left            =   120
      TabIndex        =   2
      Top             =   375
      Width           =   240
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   105
      Width           =   675
   End
End
Attribute VB_Name = "FrmEx12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fOp As New dFileOps

Private Sub cmdExit_Click()
    Unload FrmEx12
End Sub

Private Sub Form_Load()
Dim Tmp As String

    lblFile.Caption = "C:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB"
    Tmp = "C:\work\"
    
    'Example
    txtOut.Text = "ExtractFileDrive: " & fOp.ExtractFileDrive(lblFile.Caption) _
    & vbCrLf & "ExtractFileDir: " & fOp.ExtractFileDir(lblFile.Caption) _
    & vbCrLf & "ExtractFilePath: " & fOp.ExtractFilePath(lblFile.Caption) _
    & vbCrLf & "ExtractFileName: " & fOp.ExtractFileName(lblFile.Caption) _
    & vbCrLf & "ExtractFileTitle: " & fOp.ExtractFileTitle(lblFile.Caption) _
    & vbCrLf & "ExtractFileExt: " & fOp.ExtractFileExt(lblFile.Caption) _
    & vbCrLf & "FileExtExisits = " & fOp.FileExists(lblFile.Caption) _
    & vbCrLf & "FileAge: " & fOp.FileAge(lblFile.Caption) _
    & vbCrLf & "DirectoryExists: C:\work\ " & fOp.DirectoryExists(Tmp) _
    & vbCrLf & "AddBackSlash C:\Test " & fOp.AddBackSlash("C:\Test") _
    & vbCrLf & "GetFileType: this.exe " & fOp.GetFileType("this.exe") _
    & vbCrLf & vbCrLf & "See dFileOps class for more functions"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmEx12 = Nothing
End Sub

