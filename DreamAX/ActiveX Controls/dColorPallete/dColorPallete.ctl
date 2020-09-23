VERSION 5.00
Begin VB.UserControl dColorPallete 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   HasDC           =   0   'False
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ToolboxBitmap   =   "dColorPallete.ctx":0000
   Begin VB.PictureBox pSrc 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "dColorPallete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (ByRef pChoosecolor As ChooseColor) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Private Const BlockSize As Integer = 16
'Color dialog consts.
Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF

Private mColors(27) As OLE_COLOR

Private m_Color1 As OLE_COLOR
Private m_Color2 As OLE_COLOR
Private m_ColorDLGFullOpen As Boolean

Private xMouse As Integer
Private yMouse As Integer
Private MouseButton As MouseButtonConstants

Private ColorIdx As Integer

Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type RGBTRIPLE
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Event ItemClick(Button As MouseButtonConstants, ColorValue As OLE_COLOR)

Private Sub LongToRGB(LngColor As Long, RgbType As RGBTRIPLE)
On Error Resume Next
    'Convert Long Color To RGB
    RgbType.Red = (LngColor Mod 256)
    RgbType.Green = ((LngColor And &HFF00) / 256) Mod 256
    RgbType.Blue = ((LngColor And &HFF0000) / 65536)
End Sub

Private Function GetColorFromDLG(Optional InitColor As OLE_COLOR, Optional cFullOpen As Boolean = True) As Long
Dim cc As ChooseColor
Dim m_InitColor As Long
Dim aColorRef(15) As Long
Dim Counter As Integer
Dim j As RGBTRIPLE
    
    Call LongToRGB(InitColor, j)
    
    'Fill in the custom colors with shaded color
    For Counter = 240 To 15 Step -15
        aColorRef((Counter \ 15) - 1) = RGB(j.Red + Counter, j.Green + Counter, j.Blue + Counter)
    Next Counter
    
    ' Translate the initial OLE color to a long value
    If (InitColor <> 0) And OleTranslateColor(InitColor, 0, m_InitColor) Then
        m_InitColor = CLR_INVALID
    End If
    
    'Fill ChooseColor Type
    With cc
        .lStructSize = Len(cc)
        .hwndOwner = UserControl.hWnd
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = m_InitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(cFullOpen, CC_FULLOPEN, 0)
        'Show the color Dialogbox
        If ChooseColor(cc) Then
            'Return choosen color
            GetColorFromDLG = .rgbResult
        Else
            'Cancel button was pressed by user.
            GetColorFromDLG = -1
        End If
    End With
    
    'Free up used types
    ZeroMemory cc, Len(cc)
    ZeroMemory j, Len(j)
End Function

Private Sub CreateDisplayColor(oColor As OLE_COLOR)
    With pSrc
        .Cls
        .Width = 15
        .Height = 15
        'Top Line
        pSrc.Line (0, 0)-(.ScaleWidth, 0), vbWhite
        'Left Line
        pSrc.Line (0, 0)-(0, .ScaleHeight), vbWhite
        'Inner square
        pSrc.Line (1, 1)-(.ScaleWidth - 2, .ScaleHeight - 2), .BackColor, B
        'Right Gray Line
        pSrc.Line (.ScaleWidth - 1, 1)-(.ScaleWidth - 1, .ScaleHeight), vb3DShadow
        'Bottom GrayLine
        pSrc.Line (1, .ScaleHeight - 1)-(.ScaleWidth - 1, .ScaleHeight - 1), vb3DShadow
        'Center color
        pSrc.Line (2, 2)-(.ScaleWidth - 3, .ScaleHeight - 3), oColor, BF
        .Refresh
    End With
End Sub

Private Sub PutDisplayColors()
    'Create the first color
    Call CreateDisplayColor(Color1)
    BitBlt UserControl.hDC, 3, 4, 15, 15, pSrc.hDC, 0, 0, vbSrcCopy
    'Create the second color
    Call CreateDisplayColor(Color2)
    BitBlt UserControl.hDC, 10, 11, 15, 15, pSrc.hDC, 0, 0, vbSrcCopy
    'Update the control
    UserControl.Refresh
End Sub

Private Sub RenderColorItem()
On Error Resume Next
Dim X As Integer
Dim Y As Integer
Dim ColorIdx As Integer

    With UserControl
        .Cls
        'Set width and height for each color item
        With pSrc
            .Width = BlockSize
            .Height = BlockSize
            .BackColor = vbButtonFace
            
            'This is used to draw each of the color item boxes.
            pSrc.Line (0, 0)-(.ScaleWidth, 0), vb3DShadow
            pSrc.Line (0, 0)-(0, .ScaleHeight), vb3DShadow
            pSrc.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbWhite
            pSrc.Line (1, 1)-(.ScaleWidth - 2, 1), vbWhite
            pSrc.Line (1, 1)-(1, .ScaleHeight - 2), vbWhite
            pSrc.Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), vbWhite
        End With
        
        'Color preview boxes code
        UserControl.Line (0, 0)-(31, 31), vbWhite, B
        UserControl.Line (0, 0)-(31, 0), vb3DShadow
        UserControl.Line (0, 0)-(0, 31), vb3DShadow
        UserControl.Line (1, 1)-(30, 1), vbWhite
        UserControl.Line (1, 30)-(1, 1), vbWhite
        
        'Draw the Grid effect for the preview box
        For X = 2 To 30 Step 2
            For Y = 2 To 30 Step 2
                UserControl.PSet (X, Y), vbButtonFace
            Next Y
        Next X
        
        'Left preview box Gray line
        UserControl.Line (30, 1)-(30, 30), vbButtonFace
        'Bottom preview box Gray line
        UserControl.Line (1, 30)-(30, 30), vbButtonFace

        'Color items code
        'Bitblt the color items
        For X = 0 To 13
            For Y = 0 To 1
                'Draws the color
                pSrc.Line (2, 2)-(13, 13), mColors(ColorIdx), BF
                'Color Index
                ColorIdx = (ColorIdx + 1)
                BitBlt .hDC, 32 + (X * BlockSize), (Y * BlockSize), BlockSize, BlockSize, pSrc.hDC, 0, 0, vbSrcCopy
            Next Y
        Next X
        
        'Render display colors
        Call PutDisplayColors
    End With
    
End Sub

Private Sub UserControl_DblClick()
Dim TmpRgb As RGBTRIPLE
Dim LnColor As OLE_COLOR

    If (ColorIdx <> -1) Then
        LnColor = GetColorFromDLG(mColors(ColorIdx), OpenFullColorDLG)
        'Convert long color to rgb
        Call LongToRGB(LnColor, TmpRgb)
    Else
        Exit Sub
    End If
    
    'See if cancel was pressed.
    If (LnColor = -1) Then
        Exit Sub
    Else
        'Update main colors mColors
        mColors(ColorIdx) = RGB(TmpRgb.Red, TmpRgb.Green, TmpRgb.Blue)
        'Redraw pattete.
        Call RenderColorItem
        'Check the mouse button pressed.
        If (MouseButton = vbLeftButton) Then
            'Update first color
            Color1 = LnColor
        Else
            'Update second color
            Color2 = LnColor
        End If
        
        'Update display images
        Call PutDisplayColors
    End If

End Sub

Private Sub UserControl_Initialize()
    'init default Colors
    mColors(0) = 0
    mColors(1) = vbWhite
    mColors(2) = RGB(70, 70, 70)
    mColors(3) = RGB(220, 220, 220)
    mColors(4) = RGB(120, 120, 120)
    mColors(5) = RGB(180, 180, 180)
    mColors(6) = RGB(153, 0, 48)
    mColors(7) = RGB(156, 90, 60)
    mColors(8) = RGB(237, 28, 36)
    mColors(9) = RGB(255, 163, 177)
    mColors(10) = RGB(255, 126, 0)
    mColors(11) = RGB(229, 170, 122)
    mColors(12) = RGB(255, 194, 14)
    mColors(13) = RGB(245, 228, 156)
    mColors(14) = RGB(255, 242, 0)
    mColors(15) = RGB(255, 249, 189)
    mColors(16) = RGB(168, 230, 29)
    mColors(17) = RGB(221, 249, 188)
    mColors(18) = RGB(34, 177, 76)
    mColors(19) = RGB(157, 187, 97)
    mColors(20) = RGB(0, 183, 239)
    mColors(21) = RGB(153, 217, 234)
    mColors(22) = RGB(77, 109, 243)
    mColors(23) = RGB(112, 154, 209)
    mColors(24) = RGB(47, 54, 153)
    mColors(25) = RGB(84, 109, 142)
    mColors(26) = RGB(111, 49, 152)
    mColors(27) = RGB(181, 165, 213)
End Sub

Private Sub UserControl_InitProperties()
    Color1 = mColors(0)
    Color2 = mColors(1)
    OpenFullColorDLG = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    xMouse = (X \ BlockSize) - 2
    yMouse = (Y \ BlockSize)
    
    ColorIdx = (-1)
    
    If (xMouse < 0) Or (xMouse > 13) Then
    ElseIf (yMouse < 0) Or (yMouse > 1) Then
    Else
        ColorIdx = (xMouse + yMouse) + xMouse
    End If
    
    If (ColorIdx = -1) Then
        Exit Sub
    Else
        'Store mouse button presssed.
        MouseButton = Button
        RaiseEvent ItemClick(MouseButton, mColors(ColorIdx))
        
        If (MouseButton = vbLeftButton) Then
            Color1 = mColors(ColorIdx)
        Else
            Color2 = mColors(ColorIdx)
        End If
        
        'Render the preview colors
        Call PutDisplayColors
        
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Color1 = PropBag.ReadProperty("Color1", mColors(0))
    Color2 = PropBag.ReadProperty("Color2", mColors(1))
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_ColorDLGFullOpen = PropBag.ReadProperty("OpenFullColorDLG", True)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Show()
    Call RenderColorItem
End Sub

Private Sub UserControl_Terminate()
    Erase mColors
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color1", Color1, mColors(0))
    Call PropBag.WriteProperty("Color2", Color2, mColors(1))
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OpenFullColorDLG", m_ColorDLGFullOpen, True)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    'Resize the control
    UserControl.Height = (BlockSize * 2) * Screen.TwipsPerPixelY
    UserControl.Width = (13 * BlockSize + 32) * Screen.TwipsPerPixelX + 240
End Sub

Public Property Get Color1() As OLE_COLOR
    Color1 = m_Color1
End Property

Public Property Let Color1(ByVal vNewColor As OLE_COLOR)
    m_Color1 = vNewColor
    Call PutDisplayColors
    PropertyChanged "Color1"
End Property

Public Property Get Color2() As OLE_COLOR
    Color2 = m_Color2
End Property

Public Property Let Color2(ByVal vNewColor As OLE_COLOR)
    m_Color2 = vNewColor
    Call PutDisplayColors
    PropertyChanged "Color2"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get OpenFullColorDLG() As Boolean
    OpenFullColorDLG = m_ColorDLGFullOpen
End Property

Public Property Let OpenFullColorDLG(ByVal vNewValue As Boolean)
    m_ColorDLGFullOpen = vNewValue
    PropertyChanged "OpenFullColorDLG"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

