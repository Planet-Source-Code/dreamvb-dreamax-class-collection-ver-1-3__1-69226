VERSION 5.00
Begin VB.Form FrmExp40 
   Caption         =   "dPlayWave ActiveX - Example"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   4050
      TabIndex        =   10
      Top             =   1890
      Width           =   840
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   3090
      TabIndex        =   9
      Top             =   1890
      Width           =   840
   End
   Begin VB.ComboBox cboA 
      Height          =   315
      Left            =   2820
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1260
      Width           =   1215
   End
   Begin VB.ComboBox cboLoop 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1260
      Width           =   1215
   End
   Begin VB.ComboBox cboFrom 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   5490
   End
   Begin Exp40.dPlayWave dPlayWave1 
      Left            =   4485
      Top             =   1215
      _ExtentX        =   794
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1890
      Width           =   840
   End
   Begin VB.Label lblwAsynchronous 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asynchronous"
      Height          =   195
      Left            =   2820
      TabIndex        =   7
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label lblLoop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loop"
      Height          =   195
      Left            =   1515
      TabIndex        =   5
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Play From:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   990
      Width           =   735
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename or:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   255
      Width           =   855
   End
End
Attribute VB_Name = "FrmExp40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Function LoadWaveData() As Byte()
Dim fp As Long
Dim m_Bytes() As Byte
    fp = FreeFile
    Open txtFile.Text For Binary As #fp
        ReDim Preserve m_Bytes(0 To LOF(fp) - 1)
        Get #fp, , m_Bytes
    Close #fp

   LoadWaveData = m_Bytes
   Erase m_Bytes
   
End Function

Private Sub cboFrom_Click()
    Select Case cboFrom.ListIndex
        Case 0
            txtFile.Text = FixPath(App.Path) & "resources\start.wav"
        Case 1
            txtFile.Text = "ringin"
            MsgBox "App must be compiled for resource sounds to be played."
        Case 2
            txtFile.Text = FixPath(App.Path) & "resources\chaingun.wav"
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload FrmExp40
End Sub

Private Sub cmdPlay_Click()
Dim Bytes() As Byte
    With dPlayWave1
        'Filename
        .wFilename = txtFile.Text
        'Src Type
        .wPlaySource = cboFrom.ListIndex
        'Enable this to loop the waves.
        .wLoop = cboLoop.ListIndex
        .wAsynchronous = cboA.ListIndex
        'Check if playing from memory
        If (.wPlaySource = tMemory) Then
            'Set false or the app will crash
            .wAsynchronous = False
            'Load in the wave bytes
            Bytes = LoadWaveData
            'Get data pointer
            .wDataPtr = VarPtr(Bytes(0))
            'Play the wave
            .wPlay
            Exit Sub
            Erase Bytes
        End If
        'Play wave
        .wPlay
    End With
    
End Sub

Private Sub cmdStop_Click()
    dPlayWave1.wStop
End Sub

Private Sub Form_Load()
    
    'Add Play Types
    cboFrom.AddItem "Filename"
    cboFrom.AddItem "Resource"
    cboFrom.AddItem "Memory"
    cboFrom.ListIndex = 0
    'Loop
    cboLoop.AddItem "False"
    cboLoop.AddItem "True"
    cboLoop.ListIndex = 0
    cboA.AddItem "False"
    cboA.AddItem "True"
    cboA.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dPlayWave1.wStop
    Set FrmExp40 = Nothing
End Sub

