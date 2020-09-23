VERSION 5.00
Begin VB.UserControl dLed 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   HasDC           =   0   'False
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
   ToolboxBitmap   =   "dLed.ctx":0000
End
Attribute VB_Name = "dLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Private m_OnColor As OLE_COLOR
Private m_OffColor As OLE_COLOR
Private m_LedValue As LedValue

Enum LedValue
    LedOff = 0
    LedOn = 1
End Enum
'Event Declarations:
Event DblClick()
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub RenderLedDisplay()
Dim mColor As OLE_COLOR

    With UserControl
        .Cls
        If (Enabled) Then
            If (m_LedValue) Then
                'Led on color.
                .BackColor = TranslateColor(m_OnColor)
            Else
                'Led off color
                .BackColor = TranslateColor(m_OffColor)
            End If
        Else
            'Disabled color
            .BackColor = TranslateColor(&H808080)
        End If
        
        'Top
        UserControl.Line (0, 0)-(.ScaleWidth, 0), TranslateColor(&HFFFFFF)
        'Left
        UserControl.Line (0, 1)-(0, UserControl.ScaleHeight), TranslateColor(&HFFFFFF)
        'Right
        UserControl.Line (.ScaleWidth - 1, 1)-(UserControl.ScaleWidth - 1, .ScaleHeight), TranslateColor(&H808080)
        'Bottom
        UserControl.Line (0, .ScaleHeight - 1)-(.ScaleWidth - 1, .ScaleHeight - 1), TranslateColor(&H808080)
        'Inner square
        UserControl.Line (1, 1)-(.ScaleWidth - 2, .ScaleHeight - 2), TranslateColor(&HC0C0C0), B
        .Refresh
    End With
End Sub

Private Sub UserControl_InitProperties()
    OnColor = vbBlue
    OffColor = &H7F0000
End Sub

Private Sub UserControl_Resize()
    Call RenderLedDisplay
End Sub

Private Sub UserControl_Show()
    Call RenderLedDisplay
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    OnColor = PropBag.ReadProperty("OnColor", vbBlue)
    OffColor = PropBag.ReadProperty("OffColor", &H7F0000)
    Value = PropBag.ReadProperty("Value", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("OnColor", OnColor, vbBlue)
    Call PropBag.WriteProperty("OffColor", OffColor, &H7F0000)
    Call PropBag.WriteProperty("Value", Value, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Public Property Get OnColor() As OLE_COLOR
    OnColor = m_OnColor
End Property

Public Property Let OnColor(ByVal NewColor As OLE_COLOR)
    m_OnColor = NewColor
    Call RenderLedDisplay
    PropertyChanged "OnColor"
End Property

Public Property Get OffColor() As OLE_COLOR
    OffColor = m_OffColor
End Property

Public Property Let OffColor(ByVal NewColor As OLE_COLOR)
    m_OffColor = NewColor
    Call RenderLedDisplay
    PropertyChanged "OffColor"
End Property

Public Property Get Value() As LedValue
    Value = m_LedValue
End Property

Public Property Let Value(ByVal vNewValue As LedValue)
    m_LedValue = vNewValue
    Call RenderLedDisplay
    PropertyChanged "Value"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call RenderLedDisplay
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
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

