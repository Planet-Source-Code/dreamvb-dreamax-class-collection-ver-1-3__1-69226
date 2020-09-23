VERSION 5.00
Begin VB.UserControl BevelExt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   495
   ScaleWidth      =   1215
   ToolboxBitmap   =   "BevelExt.ctx":0000
End
Attribute VB_Name = "BevelExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum mLineStyle
    [PS_SOLID ] = 0
    [PS_DASH] = 1
    [PS_DOT] = 2
    [PS_DASHDOT] = 3
    [PS_DASHDOTDOT] = 4
End Enum

Enum mStyle
    [bsLowered] = 0
    [bsRaised] = 1
End Enum

Enum mShape
    [bsBox] = 0
    [bsBottomLine] = 1
    [bsTopLine] = 2
    [bsLeftLine] = 3
    [bsRightLine] = 4
    [bsSpacer] = 5
    [bsFrame] = 6
    [bsOutLine] = 7
    [bsFrameEllipse] = 8
    [bsEllipse] = 9
    [bsFocusRect] = 10
    [bsRoundFrame] = 11
    [bsRoundOutLine] = 12
End Enum

Private m_style As mStyle
Private m_shape As mShape
Private m_linestyle As mLineStyle

Private m_Transparent As Boolean
Private m_OutLineColor As OLE_COLOR
Private mTmpColor As OLE_COLOR
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."


Public Sub GuiLineTo(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional LnWidth As Long = 1, Optional LineColor As Long, Optional bStyle As mLineStyle = [PS_SOLID ])
Dim hPen As Long

    hPen = CreatePen(bStyle, LnWidth, LineColor) 'Create a soild pen
    DeleteObject SelectObject(hdc, hPen)
    
    If X1 >= 0 Then MoveToEx hdc, X1, Y1, 0
    LineTo hdc, X2, Y2 'Draw the line
    
End Sub

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0)
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub ReDrawPanel(pStyle As mStyle, pShape As mShape, pLineStyle As mLineStyle)
Dim TheDc As Long, mWidth As Long, mHeight As Long, hPen As Long, mRect As RECT
Dim LineColors(3) As Long

    'This is used to show the user control., if it's not visable
    ' we do this becase of the spacer shape we added as we do not need to see this
    ' while in usermode, only in design mode
    UserControl.Cls 'Clear
    
    If Not m_Transparent Then
        UserControl.BackColor = mTmpColor
    Else
        UserControl.BackStyle = 0
        UserControl.MaskColor = RGB(255, 0, 255)
        UserControl.BackColor = UserControl.MaskColor
    End If
    
    If Not IsWindowVisible(UserControl.hwnd) Then ShowWindow UserControl.hwnd, 1
    
    mWidth = UserControl.ScaleWidth \ Screen.TwipsPerPixelX
    mHeight = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
    TheDc = UserControl.hdc
    
    'This part works out the colors needed depending on the style
    If pStyle = bsRaised Then 'Bevel Raised Style
        LineColors(0) = TranslateColor(vbWhite)
        LineColors(1) = LineColors(0)
        LineColors(2) = TranslateColor(vb3DShadow)
        LineColors(3) = LineColors(2)
    End If
    
    If pStyle = bsLowered Then 'Bevel Lowered style
        LineColors(0) = TranslateColor(vb3DShadow)
        LineColors(1) = LineColors(0)
        LineColors(2) = TranslateColor(vbWhite)
        LineColors(3) = LineColors(2)
    End If
    
    Select Case pShape
        Case bsBox 'Bevel Box
            GuiLineTo TheDc, 0, 0, mWidth, 0, 0, LineColors(0), pLineStyle      'Top-Line
            GuiLineTo TheDc, 0, 0, 0, mHeight, 0, LineColors(1), pLineStyle  'Left-Line
            GuiLineTo TheDc, mWidth - 1, 0, mWidth - 1, mHeight, 0, LineColors(2), pLineStyle  'Right-Line
            GuiLineTo TheDc, 0, mHeight - 1, mWidth, mHeight - 1, 0, LineColors(3), pLineStyle  'Bottom-Line
        Case bsBottomLine 'Bottom 3DLine
            GuiLineTo TheDc, 0, mHeight - 2, mWidth, mHeight - 2, 0, LineColors(0), pLineStyle  'Bottom-Line 1
            GuiLineTo TheDc, 0, mHeight - 1, mWidth, mHeight - 1, 0, LineColors(3), pLineStyle  'Bottom-Line 2
        Case bsTopLine 'Top 3DLine
            GuiLineTo TheDc, 0, 0, mWidth, 0, 0, LineColors(0), pLineStyle  'Top-Line 1
            GuiLineTo TheDc, 0, 1, mWidth, 1, 0, LineColors(3), pLineStyle  'Top-Line 2
        Case bsLeftLine 'Left 3DLine
            GuiLineTo TheDc, 0, 0, 0, mHeight, 0, LineColors(0), pLineStyle  'Left-Line 1
            GuiLineTo TheDc, 1, 0, 1, mHeight, 0, LineColors(3), pLineStyle  'Left-Line 2
        Case bsRightLine 'Right 3DLine
            GuiLineTo TheDc, mWidth - 1, 0, mWidth - 1, mHeight, 0, LineColors(3), pLineStyle  'Right-Line 1
            GuiLineTo TheDc, mWidth - 2, 0, mWidth - 2, mHeight, 0, LineColors(0), pLineStyle  'Right-Line 2
        Case bsSpacer 'Spacer note. only visable in design-mode
            If UserControl.Ambient.UserMode Then ShowWindow UserControl.hwnd, 0
            'Hide the usercontrol when not in design mode
            GuiLineTo TheDc, 0, 0, mWidth, 0, , 0, PS_DASHDOTDOT 'Top-Line
            GuiLineTo TheDc, 0, 0, 0, mHeight, , 0, PS_DASHDOTDOT 'Left-Line
            GuiLineTo TheDc, mWidth - 1, 0, mWidth - 1, mHeight, , 0, PS_DASHDOTDOT 'Right-Line
            GuiLineTo TheDc, 0, mHeight - 1, mWidth, mHeight - 1, , 0, PS_DASHDOTDOT 'Bottom-Line
        Case bsFrame 'Draw a Frame just like the one in vb without a caption
            GuiLineTo TheDc, 0, 0, mWidth - 1, 0, 0, LineColors(0), pLineStyle  'Top-Line 1
            GuiLineTo TheDc, 0, 1, mWidth, 1, 0, LineColors(3), pLineStyle  'Top-Line 2
            '
            GuiLineTo TheDc, 0, 1, 0, mHeight, 0, LineColors(0), pLineStyle  'Left-Line 1
            GuiLineTo TheDc, 1, 1, 1, mHeight, 0, LineColors(3), pLineStyle  'Left-Line 2
            '
            GuiLineTo TheDc, mWidth - 1, 1, mWidth - 1, mHeight, 0, LineColors(3), pLineStyle  'Right-Line 1
            GuiLineTo TheDc, mWidth - 2, 1, mWidth - 2, mHeight, 0, LineColors(0), pLineStyle  'Right-Line 2
            '
            GuiLineTo TheDc, 0, mHeight - 2, mWidth - 1, mHeight - 2, 0, LineColors(0), pLineStyle  'Bottom-Line
            GuiLineTo TheDc, 0, mHeight - 1, mWidth - 1, mHeight - 1, 0, LineColors(3), pLineStyle  'Bottom-Line
        Case bsFrameEllipse ' draw a frame Ellipse like a vb frame control
            GuiEllipse TheDc, 1, 1, mWidth + 1, mHeight - 1, 0, LineColors(3), pLineStyle
            GuiEllipse TheDc, 0, 0, mWidth, mHeight - 2, 0, LineColors(0), pLineStyle
        Case bsEllipse 'draw an Ellipse using the outline color
            GuiEllipse TheDc, 0, 0, mWidth, mHeight, 0, m_OutLineColor, pLineStyle
        Case bsOutLine ' Draw a rectangle with a colored outline
            hPen = CreatePen(pLineStyle, 0, TranslateColor(m_OutLineColor))
            DeleteObject SelectObject(TheDc, hPen)
            Rectangle TheDc, 0, 0, mWidth, mHeight
            'RoundRect TheDc, 0, 0, mWidth, mHeight, 5, 5
        Case bsFocusRect 'Draws a focus rect
            SetRect mRect, 0, 0, mWidth, mHeight
            DrawFocusRect TheDc, mRect
        Case bsRoundFrame 'Rounded Frame
            RoundRectangle TheDc, 0, 0, mWidth - 2, mHeight - 1, 12, 12, 0, LineColors(0), pLineStyle
            RoundRectangle TheDc, 1, 1, mWidth - 1, mHeight, 11, 11, 0, LineColors(3), pLineStyle
        Case bsRoundOutLine
            RoundRectangle TheDc, 0, 0, mWidth, mHeight, 12, 12, 0, TranslateColor(m_OutLineColor), pLineStyle
    End Select
    
    UserControl.MaskPicture = UserControl.Image
    UserControl.Refresh
    
End Sub

Private Sub RoundRectangle(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, Optional LnWidth As Long = 1, Optional LineColor As Long, Optional bStyle As mLineStyle = [PS_SOLID ])
Dim hPen As Long
    hPen = CreatePen(bStyle, LnWidth, LineColor)
    DeleteObject SelectObject(hdc, hPen)
    RoundRect hdc, X1, Y1, X2, Y2, X3, Y3
End Sub

Private Sub GuiEllipse(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional LnWidth As Long = 1, Optional LineColor As Long, Optional bStyle As mLineStyle = [PS_SOLID ])
Dim hPen As Long
    hPen = CreatePen(bStyle, LnWidth, LineColor)
    DeleteObject SelectObject(hdc, hPen)
    Ellipse hdc, X1, Y1, X2, Y2
End Sub

Public Property Get Style() As mStyle
    Style = m_style
End Property

Public Property Let Style(ByVal vNewStyle As mStyle)
    m_style = vNewStyle
    PropertyChanged "Style"
    ReDrawPanel m_style, m_shape, m_linestyle
End Property

Private Sub UserControl_Initialize()
    m_style = bsLowered
    m_Transparent = True
    m_OutLineColor = vbBlack
    mTmpColor = vbButtonFace
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_style = PropBag.ReadProperty("Style", 0)
    m_shape = PropBag.ReadProperty("Shape", 0)
    m_OutLineColor = PropBag.ReadProperty("OutLineColor", vbBlack)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    mTmpColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    m_Transparent = PropBag.ReadProperty("Transparent", True)
    m_linestyle = PropBag.ReadProperty("LineStyle", 0)
End Sub

Private Sub UserControl_Resize()
    ReDrawPanel m_style, m_shape, m_linestyle
End Sub

Private Sub UserControl_Show()
    ReDrawPanel m_style, m_shape, m_linestyle
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", m_style, 0)
    Call PropBag.WriteProperty("Shape", m_shape, 0)
    Call PropBag.WriteProperty("OutLineColor", m_OutLineColor, vbBlack)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Transparent", m_Transparent, True)
    Call PropBag.WriteProperty("BackColor", mTmpColor, vbButtonFace)
    Call PropBag.WriteProperty("LineStyle", m_linestyle, 0)
End Sub

Public Property Get Shape() As mShape
    Shape = m_shape
End Property

Public Property Let Shape(ByVal vNewShape As mShape)
    m_shape = vNewShape
    PropertyChanged "Shape"
    ReDrawPanel m_style, m_shape, m_linestyle
End Property

Public Property Get OutLineColor() As OLE_COLOR
    OutLineColor = m_OutLineColor
End Property

Public Property Let OutLineColor(ByVal vNewOutLineColor As OLE_COLOR)
    m_OutLineColor = vNewOutLineColor
    PropertyChanged "OutLineColor"
    ReDrawPanel m_style, m_shape, m_linestyle
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = mTmpColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    mTmpColor = New_BackColor
    PropertyChanged "BackColor"
    ReDrawPanel m_style, m_shape, m_linestyle
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal vNewValue As Boolean)
    m_Transparent = vNewValue
    PropertyChanged "Transparent"
    ReDrawPanel m_style, m_shape, m_linestyle
End Property

Public Property Get LineStyle() As mLineStyle
    LineStyle = m_linestyle
End Property

Public Property Let LineStyle(ByVal vNewLStyle As mLineStyle)
    m_linestyle = vNewLStyle
    PropertyChanged "LineStyle"
    ReDrawPanel m_style, m_shape, m_linestyle
End Property
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
    UserControl.Size Width, Height
End Sub

