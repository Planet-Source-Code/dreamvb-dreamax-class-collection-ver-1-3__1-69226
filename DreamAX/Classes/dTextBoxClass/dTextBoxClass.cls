VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dTextBoxClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private WithEvents mTxtObj As TextBox
Attribute mTxtObj.VB_VarHelpID = -1
Private m_Modified As Boolean

Private Const WM_CUT As Long = &H300
Private Const WM_COPY As Long = &H301
Private Const WM_PASTE As Long = &H302
Private Const WM_CLEAR As Long = &H303

Private Const EM_UNDO = &HC7
Private Const EM_CANUNDO = &HC6
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINELENGTH = &HC1
Private Const EM_FMTLINES = &HC8

Private Sub mtxtObj_Change()
    Modified = True
End Sub

Public Property Get TextBox() As TextBox
    Set TextBox = mTxtObj
End Property

Public Property Get TextLength() As Long
    TextLength = Len(mTxtObj)
End Property

Public Property Let TextBox(ByVal NewObj As TextBox)
    Set mTxtObj = NewObj
End Property

Public Property Get Modified() As Boolean
    Modified = m_Modified
End Property

Public Property Let Modified(ByVal NewModified As Boolean)
    m_Modified = NewModified
End Property

Public Sub AppendText(ByVal Text As String)
    mTxtObj.SelStart = Len(mTxtObj.Text)
    mTxtObj.SelText = Text
End Sub

Public Property Get SelectedText() As String
    SelectedText = mTxtObj.SelText
End Property

Public Property Get SelectionStart() As Long
    SelectionStart = mTxtObj.SelStart
End Property

Public Property Let SelectionStart(ByVal NewStart As Long)
    mTxtObj.SelStart = NewStart
End Property
'
Public Property Get SelectionLength() As Long
    SelectionLength = mTxtObj.SelLength
End Property

Public Property Let SelectionLength(ByVal vNewLen As Long)
    mTxtObj.SelLength = vNewLen
End Property

Public Property Get Lines(Index As Long) As String
Dim mLines() As String
Dim Idx As Long

    'Line Index
    Idx = (Index - 1)
    'Check vaild index
    If (Idx <= 0) Then Idx = 1
    'Split the lines
    mLines = Split(mTxtObj.Text, vbCrLf)
    Lines = mLines(Idx)
    '
    Erase mLines
    
End Property

Public Property Get LinesCount() As Long
    LinesCount = SendMessage(mTxtObj.hwnd, EM_GETLINECOUNT, -1, 0)
End Property

Public Property Get LineColumn() As Long
    'Returns the cursor colum pos
    Dim Idx As Long
    
    Idx = mTxtObj.SelStart - SendMessage(mTxtObj.hwnd, EM_LINEINDEX, (LineIndex - 1), ByVal 0&)
    
    If (Idx <= 0) Then
        Idx = 1
    End If
    
    LineColumn = Idx
    
End Property

Public Property Get LineIndex() As Long
Dim Idx As Long
    
    'Get the line index from the cursot position
    Idx = SendMessage(mTxtObj.hwnd, EM_LINEFROMCHAR, -1, ByVal 0&)
    
    If (Idx >= 0) Then
        LineIndex = (Idx + 1)
    Else
        LineIndex = Idx
    End If
    
End Property

Public Property Get LineLength(ByVal Number As Long) As Long
Dim iRet As Long
Dim Idx As Long
Dim OldCelPos As Long

    'Store the line Number
    Idx = Number
    OldCelPos = mTxtObj.SelStart
    
    If (Idx > LinesCount) Then Idx = 0
    'Goto the current Line
    Call GotoLine(Idx)
    'Get the length of the line
    iRet = SendMessage(mTxtObj.hwnd, EM_LINELENGTH, mTxtObj.SelStart, ByVal 0&)
    'Check vaild Index
    If (iRet <= 0) Then iRet = 0
    
    mTxtObj.SelStart = OldCelPos
    LineLength = iRet
End Property

Public Property Get Text() As String
    Text = mTxtObj.Text
End Property

Public Property Let Text(ByVal NewText As String)
    mTxtObj.Text = NewText
End Property

Property Get CanUndo() As Boolean
    CanUndo = SendMessage(mTxtObj.hwnd, EM_CANUNDO, 0, 0) <> 0
End Property

Public Property Get CanPaste() As Boolean
    CanPaste = Len(Clipboard.GetText(vbCFText)) > 0
End Property

Sub WordWrap(Wrap As Boolean)
    SendMessage mTxtObj.hwnd, EM_FMTLINES, -1, 0
End Sub

Public Sub Cut()
    SendMessage mTxtObj.hwnd, WM_CUT, 0, 0
End Sub

Public Sub Copy()
    SendMessage mTxtObj.hwnd, WM_COPY, 0, 0
End Sub

Public Sub Paste()
    SendMessage mTxtObj.hwnd, WM_PASTE, 0, 0
End Sub

Public Sub Delete()
    SendMessage mTxtObj.hwnd, WM_CLEAR, 0, 0
End Sub

Public Sub Undo()
    SendMessage mTxtObj.hwnd, EM_UNDO, 0, 0
End Sub

Public Sub Clear()
    mTxtObj.Text = vbNullString
End Sub

Public Sub LineDelete(ByVal Number As Long)
    Call SelectLine(Number)
    mTxtObj.SelText = ""
End Sub

Public Sub GotoLine(ByVal Number As Long)
Dim Idx As Long
Dim iRet As Long

    Idx = Number
    
    If (Idx <= 0) Then Idx = 1
    If (Idx > LinesCount) Then Idx = LinesCount
    
    iRet = SendMessage(mTxtObj.hwnd, EM_LINEINDEX, (Idx - 1), 0)
    If (iRet <= 0) Then iRet = 0
    
    mTxtObj.SelStart = iRet
    mTxtObj.SelLength = 0
    mTxtObj.SetFocus
    
End Sub

Public Sub LoadFromFile(ByVal Filename As String)
Dim fp As Long
Dim Bytes() As Byte
    fp = FreeFile
    
    Open Filename For Binary As #fp
        If LOF(fp) > 0 Then
            ReDim Bytes(0 To LOF(fp))
            Get #fp, , Bytes
        End If
    Close #fp
    
    mTxtObj.Text = StrConv(Bytes, vbUnicode)
    Erase Bytes
End Sub

Public Sub SelectAll()
    mTxtObj.SelStart = 0
    mTxtObj.SelLength = Len(mTxtObj.Text)
    mTxtObj.SetFocus
End Sub

Public Sub SelectRange(Start As Integer, Length As Integer)
    SelectionStart = Start
    SelectionLength = Length
End Sub

Public Sub SelectLine(ByVal Number As Long)
Dim Idx As Long
    'Store Line Number
    Idx = Number
    If (Idx > LinesCount) Then
        Exit Sub
    Else
        'Goto the current Line
        Call GotoLine(Idx)
        mTxtObj.SelLength = LineLength(Idx)
        mTxtObj.SetFocus
    End If
End Sub

Public Sub SetFocus()
    mTxtObj.SetFocus
End Sub

Public Sub SaveToFile(ByVal Filename As String)
Dim fp As Long
    fp = FreeFile
    Open Filename For Output As #fp
        Print #fp, mTxtObj.Text
    Close #fp
End Sub

