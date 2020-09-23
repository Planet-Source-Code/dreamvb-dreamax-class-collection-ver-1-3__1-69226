VERSION 5.00
Begin VB.UserControl dTxtBoxBorder 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ScaleHeight     =   495
   ScaleWidth      =   1215
   ToolboxBitmap   =   "dTxtBoxBorder.ctx":0000
   Begin VB.TextBox txtBox 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1170
   End
End
Attribute VB_Name = "dTxtBoxBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum mTextAlignemnt
    aLeft = 0
    aRight = 1
    aCenter = 2
End Enum

Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()

Private m_BorderColor As OLE_COLOR
Private m_ShowBorder As Boolean

Private Sub DrawBorder()
Dim m_OffsetXY As Long
Dim m_OffSet As Long
    
    m_OffsetXY = 0
    m_OffSet = 0
    
    With UserControl
       ' .Cls
        If (m_ShowBorder) Then
            m_OffsetXY = 15
            m_OffSet = 30
            UserControl.Line (0, 0)-(.ScaleWidth - 8, .ScaleHeight - 8), m_BorderColor, B
            .Refresh
        End If
        
        'Resize Textbox
        txtBox.Move m_OffsetXY, m_OffsetXY, .ScaleWidth - m_OffSet, .ScaleHeight - m_OffSet
    End With
    
End Sub

Private Sub UserControl_InitProperties()
    txtBox.Text = Ambient.DisplayName
    ShowBorder = True
    BorderColor = vbBlue
End Sub

Private Sub UserControl_Resize()
    DrawBorder
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtBox.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtBox.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Alignment() As mTextAlignemnt
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtBox.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As mTextAlignemnt)
    txtBox.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Private Sub txtBox_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtBox.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtBox.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtBox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtBox.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtBox.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtBox.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtBox.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtBox.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Private Sub txtBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = txtBox.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtBox.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub txtBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = txtBox.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    txtBox.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub txtBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtBox.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtBox.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtBox.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtBox.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtBox.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtBox.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtBox.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtBox.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtBox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtBox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub txtBox_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtBox.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtBox.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtBox.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtBox.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtBox.Locked = PropBag.ReadProperty("Locked", False)
    txtBox.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtBox.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtBox.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtBox.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtBox.SelText = PropBag.ReadProperty("SelText", "")
    txtBox.Text = PropBag.ReadProperty("Text", Ambient.DisplayName)
    txtBox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    BorderColor = PropBag.ReadProperty("BorderColor", vbBlue)
    ShowBorder = PropBag.ReadProperty("ShowBorder", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", txtBox.BackColor, &H80000005)
    Call PropBag.WriteProperty("Alignment", txtBox.Alignment, 0)
    Call PropBag.WriteProperty("Enabled", txtBox.Enabled, True)
    Call PropBag.WriteProperty("Font", txtBox.Font, Ambient.Font)
    Call PropBag.WriteProperty("Locked", txtBox.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txtBox.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtBox.MousePointer, 0)
    Call PropBag.WriteProperty("SelLength", txtBox.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtBox.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtBox.SelText, "")
    Call PropBag.WriteProperty("Text", txtBox.Text, Ambient.DisplayName)
    Call PropBag.WriteProperty("ForeColor", txtBox.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BorderColor", BorderColor, vbBlue)
    Call PropBag.WriteProperty("ShowBorder", ShowBorder, True)
End Sub

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal vNewColor As OLE_COLOR)
    m_BorderColor = vNewColor
    Call DrawBorder
    PropertyChanged "BorderColor"
End Property

Public Property Get ShowBorder() As Boolean
    ShowBorder = m_ShowBorder
End Property

Public Property Let ShowBorder(ByVal vShow As Boolean)
    m_ShowBorder = vShow
    Call DrawBorder
    PropertyChanged "BorderColor"
End Property
