VERSION 5.00
Begin VB.UserControl dFilePatten 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   HasDC           =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   1215
   ToolboxBitmap   =   "dFilePatten.ctx":0000
   Begin VB.ComboBox cboPatten 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "dFilePatten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Patten As String
Private m_FileListBox As FileListBox
Private cboTmp As String
'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Function ExtractPatten(lStr As String) As String
Dim s_pos As Integer
Dim e_pos As Integer
On Error Resume Next

    s_pos = InStr(1, lStr, "(", vbBinaryCompare)
    e_pos = InStr(s_pos + 1, lStr, ")", vbBinaryCompare)
    
    If (s_pos > 0) And (e_pos > 0) Then
        ExtractPatten = Trim(Mid(lStr, s_pos + 1, e_pos - s_pos - 1))
    End If
    
    s_pos = 0
    e_pos = 0
    
End Function

Private Sub AddPatten()
Dim vLst() As String
Dim Cnt As Integer

    cboPatten.Clear
    If Len(Trim(m_Patten)) = 0 Then Exit Sub
    '
    vLst = Split(m_Patten, "|")
    For Cnt = 0 To UBound(vLst)
        cboPatten.AddItem vLst(Cnt)
    Next Cnt
    
End Sub

Private Sub cboPatten_Change()
    cboPatten.Text = cboTmp
End Sub

Private Sub cboPatten_Click()
    cboTmp = cboPatten.Text
    cboTmp = Replace(cboTmp, ",", ";")
    If m_FileListBox Is Nothing Then Exit Sub
    m_FileListBox.Pattern = ExtractPatten(cboTmp)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    m_Patten = PropBag.ReadProperty("Patten", "")
    Set m_FileListBox = PropBag.ReadProperty("vFileListBox", Nothing)
    
    If (cboPatten.ListCount <> 0) Then
        cboPatten.ListIndex = PropBag.ReadProperty("PattenIndex", 0)
    End If

    cboPatten.List(Index) = PropBag.ReadProperty("List" & Index, "")
    cboPatten.Enabled = PropBag.ReadProperty("Enabled", True)
    Set cboPatten.Font = PropBag.ReadProperty("Font", Ambient.Font)
    cboPatten.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    cboPatten.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    cboPatten.Text = PropBag.ReadProperty("Text", "")
    cboPatten.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    cboPatten.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    cboPatten.Enabled = PropBag.ReadProperty("Enabled", True)
    cboPatten.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set cboPatten.Font = PropBag.ReadProperty("Font", Ambient.Font)
    cboPatten.Text = PropBag.ReadProperty("Text", "")
    cboPatten.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    cboPatten.Width = UserControl.Width
    UserControl.Height = cboPatten.Height
End Sub

Public Property Get Patten() As String
    Patten = m_Patten
End Property

Public Property Let Patten(vNewPatten As String)
    m_Patten = vNewPatten
    Call AddPatten
    PropertyChanged "Patten"
End Property

Private Sub UserControl_Terminate()
    cboPatten.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
    Call PropBag.WriteProperty("Patten", m_Patten, "")
    Call PropBag.WriteProperty("vFileListBox", m_FileListBox, Nothing)
    Call PropBag.WriteProperty("PattenIndex", cboPatten.ListIndex, 0)
    Call PropBag.WriteProperty("List" & Index, cboPatten.List(Index), "")
    Call PropBag.WriteProperty("Enabled", cboPatten.Enabled, True)
    Call PropBag.WriteProperty("Font", cboPatten.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", cboPatten.ForeColor, &H80000008)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", cboPatten.MousePointer, 0)
    Call PropBag.WriteProperty("Text", cboPatten.Text, "")
    Call PropBag.WriteProperty("ToolTipText", cboPatten.ToolTipText, "")
    Call PropBag.WriteProperty("ForeColor", cboPatten.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", cboPatten.Enabled, True)
    Call PropBag.WriteProperty("MousePointer", cboPatten.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Font", cboPatten.Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", cboPatten.Text, "")
    Call PropBag.WriteProperty("ToolTipText", cboPatten.ToolTipText, "")
End Sub

Public Property Get vFileListBox() As FileListBox
    vFileListBox = m_FileListBox
End Property

Public Property Let vFileListBox(ByVal vNewValue As FileListBox)
    Set m_FileListBox = vNewValue
    PropertyChanged "vFileListBox"
End Property

Public Property Get PattenIndex() As Integer
Attribute PattenIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    PattenIndex = cboPatten.ListIndex
End Property

Public Property Let PattenIndex(ByVal New_PattenIndex As Integer)
    cboPatten.ListIndex() = New_PattenIndex
    PropertyChanged "PattenIndex"
End Property
'
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = cboPatten.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    cboPatten.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = cboPatten.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    cboPatten.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = cboPatten.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    cboPatten.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = cboPatten.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set cboPatten.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cboPatten.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cboPatten.Font = New_Font
    PropertyChanged "Font"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    cboPatten.Refresh
End Sub

Private Sub cboPatten_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboPatten_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cboPatten_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = cboPatten.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    cboPatten.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = cboPatten.ListCount
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = cboPatten.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    cboPatten.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    cboPatten.RemoveItem Index
End Sub

