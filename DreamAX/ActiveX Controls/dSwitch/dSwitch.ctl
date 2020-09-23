VERSION 5.00
Begin VB.UserControl dSwitch 
   AutoRedraw      =   -1  'True
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   72
   ToolboxBitmap   =   "dSwitch.ctx":0000
   Begin VB.PictureBox pSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   60
      Picture         =   "dSwitch.ctx":0312
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "dSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Private m_SwitchState As Boolean
Private m_OnColor As OLE_COLOR
Private m_OffColor As OLE_COLOR

Private TKey As KeyCodeConstants
Private m_AllowToggle As Boolean

Enum TState
    bOff = 0
    bOn = 1
End Enum

Event StateChange(ButtonState As TState)

Event DblClick()
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub SetSwitchState()
Dim W As Long
Dim H As Long
Dim BoxColor As OLE_COLOR

    UserControl.Cls

    W = (pSrc.ScaleWidth \ 2)
    H = pSrc.ScaleHeight
    
    If (m_SwitchState) Then
        'Set the switch in the on state
        TransparentBlt UserControl.hdc, 0, 0, W, H, pSrc.hdc, 0, 0, W, H, RGB(255, 0, 255)
        'This sets the little square color when the switch is in the on state
        If (UserControl.Enabled) Then
            'Set the enabled color
            BoxColor = m_OnColor
        Else
            'Set the disabled color
            BoxColor = vb3DShadow
        End If
        UserControl.Line (11, 32)-(17, 29), BoxColor, BF
    Else
        'Set the switch in the off state
        TransparentBlt UserControl.hdc, 0, 0, W, H, pSrc.hdc, 32, 0, W, H, RGB(255, 0, 255)
        If (UserControl.Enabled) Then
            'Set the enabled color
            BoxColor = m_OffColor
        Else
            'Set the disabled color
            BoxColor = vb3DShadow
        End If
        
        UserControl.Line (16, 26)-(8, 26), BoxColor
        UserControl.Line (15, 27)-(8, 27), BoxColor
        UserControl.Line (15, 28)-(7, 28), BoxColor
    End If
    
    UserControl.Refresh
    
End Sub

Private Sub UserControl_InitProperties()
    OnColorState = vbRed
    OffColorState = vbWhite
    m_SwitchState = True
    m_AllowToggle = True
    TKey = vbKeySpace
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If (Button = vbLeftButton) Then
        m_SwitchState = (Not m_SwitchState)
        'Update the switch state
        Call SetSwitchState
        RaiseEvent StateChange(Abs(m_SwitchState))
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Size (pSrc.ScaleWidth * Screen.TwipsPerPixelX) \ 2, (pSrc.ScaleHeight * Screen.TwipsPerPixelY)
    SetSwitchState
End Sub

Private Sub UserControl_Show()
    SetSwitchState
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call SetSwitchState
    PropertyChanged "BackColor"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    OnColorState = PropBag.ReadProperty("OnColorState", vbRed)
    OffColorState = PropBag.ReadProperty("OffColorState", vbWhite)
    m_SwitchState = PropBag.ReadProperty("ButtonState", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    TKey = PropBag.ReadProperty("ToggleKey", vbKeySpace)
    m_AllowToggle = PropBag.ReadProperty("AllowToggleSupport", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("OnColorState", OnColorState, vbRed)
    Call PropBag.WriteProperty("OffColorState", OffColorState, vbWhite)
    Call PropBag.WriteProperty("ButtonState", m_SwitchState, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ToggleKey", TKey, vbKeySpace)
    Call PropBag.WriteProperty("AllowToggleSupport", m_AllowToggle, True)
End Sub

Public Property Get ButtonState() As TState
    ButtonState = Abs(m_SwitchState)
End Property

Public Property Let ButtonState(ByVal vNewValue As TState)
    m_SwitchState = CBool(vNewValue)
    Call SetSwitchState
    PropertyChanged "ButtonState"
End Property

Public Property Get OnColorState() As OLE_COLOR
    OnColorState = m_OnColor
End Property

Public Property Let OnColorState(ByVal vNewValue As OLE_COLOR)
    m_OnColor = vNewValue
    Call SetSwitchState
    PropertyChanged "OnColorState"
End Property

Public Property Get OffColorState() As OLE_COLOR
    OffColorState = m_OffColor
End Property

Public Property Let OffColorState(ByVal vNewValue As OLE_COLOR)
    m_OffColor = vNewValue
    Call SetSwitchState
    PropertyChanged "OffColorState"
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call SetSwitchState
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If (m_AllowToggle) And (KeyCode = TKey) Then
        'If KeyCode Maths the Toggle key then we can set the button state.
        Call UserControl_MouseDown(vbLeftButton, 0, 0, 0)
    End If
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

Public Property Get ToggleKey() As KeyCodeConstants
    ToggleKey = TKey
End Property

Public Property Let ToggleKey(ByVal vNewTKey As KeyCodeConstants)
    TKey = vNewTKey
    PropertyChanged "ToggleKey"
End Property

Public Property Get AllowToggleSupport() As Boolean
    AllowToggleSupport = m_AllowToggle
End Property

Public Property Let AllowToggleSupport(ByVal vNewValue As Boolean)
    m_AllowToggle = vNewValue
    PropertyChanged "AllowToggleSupport"
End Property
