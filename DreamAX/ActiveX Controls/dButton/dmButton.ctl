VERSION 5.00
Begin VB.UserControl dmButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   77
   ToolboxBitmap   =   "dmButton.ctx":0000
   Begin VB.PictureBox pIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "dmButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Enum TAlignment
    aLeft = 0
    aRight = 1
    aCenter = 2
End Enum

Enum ButtonState
    bFlat = 0
    bDown = 1
    bUp = 2
End Enum

Private bState As ButtonState
Private m_ForeColor As OLE_COLOR
Private bLineColor(5) As OLE_COLOR
Private m_caption As String
Private m_StyleSheet As String
Private StyleLoaded As Boolean

'Image Alignment
Private Img_Align As TAlignment
'Caption Alignment
Private m_Align As TAlignment

Event Click(Button As Integer)
Event AccessKeyChange()

Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_BTNHIGHLIGHT As Long = 20

Private Function ReadIniValue(Selection As String, ValueName As String, FileName As String, Optional Default As String = "")
Dim iRet As Long
Dim Buf As String
    'Create buffer
    Buf = Space(128)
    'Get the ini value
    iRet = GetPrivateProfileString(Selection, ValueName, "", Buf, 128, FileName)
    'Return the value
    If (iRet <> 0) Then ReadIniValue = Left(Buf, iRet)
    'Clear up
    Buf = vbNullString
    iRet = 0
End Function

Private Sub LoadStyle(Selection As String)
    If (StyleLoaded) Then
        bLineColor(0) = Val(ReadIniValue(Selection, "top", m_StyleSheet, "0"))
        bLineColor(1) = Val(ReadIniValue(Selection, "left", m_StyleSheet, "0"))
        bLineColor(2) = Val(ReadIniValue(Selection, "right", m_StyleSheet, "0"))
        bLineColor(3) = Val(ReadIniValue(Selection, "bottom", m_StyleSheet, "0"))
        bLineColor(4) = Val(ReadIniValue(Selection, "backcolor", m_StyleSheet, "0"))
        bLineColor(5) = Val(ReadIniValue(Selection, "Forecolor", m_StyleSheet, "0"))
    End If
End Sub

Private Sub RenderButton()
Dim CenterY As Integer
Dim xImgPos As Integer
Dim xCapPos As Integer
Dim ch As String
Dim chCount As Integer
        
Const ImgWidth As Integer = 16
    'Office97 Default button
    If (bState = bFlat) Then
        bLineColor(0) = GetSysColor(COLOR_BTNFACE)
        bLineColor(1) = bLineColor(0)
        bLineColor(2) = bLineColor(0)
        bLineColor(3) = bLineColor(0)
        bLineColor(4) = bLineColor(0)
        bLineColor(5) = m_ForeColor
    ElseIf (bState = bUp) Then
        bLineColor(0) = GetSysColor(COLOR_BTNHIGHLIGHT)
        bLineColor(1) = bLineColor(0)
        bLineColor(2) = GetSysColor(COLOR_BTNSHADOW)
        bLineColor(3) = bLineColor(2)
        bLineColor(4) = GetSysColor(COLOR_BTNFACE)
        bLineColor(5) = m_ForeColor
    Else
        bLineColor(0) = GetSysColor(COLOR_BTNSHADOW)
        bLineColor(1) = bLineColor(0)
        bLineColor(2) = GetSysColor(COLOR_BTNHIGHLIGHT)
        bLineColor(3) = bLineColor(2)
        bLineColor(4) = GetSysColor(COLOR_BTNFACE)
        bLineColor(5) = m_ForeColor
    End If
    
    If (bState = bFlat) Then
        Call LoadStyle("Button.Flat")
    ElseIf (bState = bUp) Then
        Call LoadStyle("Button.hover")
    Else
        Call LoadStyle("Button.down")
    End If
    
    With UserControl
        'Used for centering the image
        CenterY = (.ScaleHeight - ImgWidth) \ 2
        .Cls
        UserControl.BackColor = bLineColor(4)
        UserControl.ForeColor = bLineColor(5)
        UserControl.Line (0, 0)-(.ScaleWidth - 1, 0), bLineColor(0) 'TopLine
        UserControl.Line (0, 0)-(0, .ScaleHeight - 1), bLineColor(1) 'Left Line
        UserControl.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight), bLineColor(2)   ' RightLine
        UserControl.Line (0, .ScaleHeight - 1)-(.ScaleWidth - 1, .ScaleHeight - 1), bLineColor(3) 'Buttom Line
        
        'Image Alignment
        Select Case Img_Align
            Case aLeft
                xImgPos = 2
            Case aCenter
                xImgPos = (.ScaleWidth - ImgWidth) \ 2
            Case aRight
                xImgPos = (.ScaleWidth - ImgWidth) - 2
        End Select
        
        'Caption Alignment
        Select Case m_Align
            Case aLeft
                xCapPos = 20
            Case aCenter
                xCapPos = (.ScaleWidth - .TextWidth(m_caption)) \ 2
            Case aRight
                xCapPos = (.ScaleWidth - .TextWidth(m_caption) - 2)
        End Select
        
        
        'add on the icon
        If (pIcon.Picture <> 0) Then
            TransparentBlt .hdc, xImgPos, CenterY, ImgWidth, ImgWidth, pIcon.hdc, 0, 0, ImgWidth, ImgWidth, UserControl.MaskColor
        Else
            If (m_Align = aLeft) Then
                xCapPos = 2
            End If
        End If
        
        'Print the buttons caption
        .CurrentX = xCapPos
        .CurrentY = CenterY + 2
        
        'Added for Access key support
        For chCount = 1 To Len(m_caption)
            ch = Mid$(m_caption, chCount, 1)
            If Mid$(m_caption, chCount, 1) = "&" Then
                .FontUnderline = True
                .AccessKeys = Mid(m_caption, chCount + 1, 1)
            Else
                 UserControl.Print ch;
                .FontUnderline = False
            End If
        Next
        
        .Refresh
    End With
    
    ch = vbNullString
    chCount = 0
    
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent AccessKeyChange
End Sub

Private Sub UserControl_InitProperties()
    m_caption = "CoolButton"
    m_ForeColor = vbBlack
    Img_Align = aLeft
    m_Align = aLeft
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    bState = bDown
    Call RenderButton
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If (x < 0) Or (x > UserControl.ScaleWidth) _
    Or (Y < 0) Or (Y > UserControl.ScaleHeight) Then
        ReleaseCapture
        bState = bFlat
        Call RenderButton
    ElseIf GetCapture() <> UserControl.hwnd Then
        bState = bUp
        Call RenderButton
        SetCapture UserControl.hwnd
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent Click(Button)
    bState = bFlat
    Call RenderButton
End Sub

Private Sub UserControl_Resize()
    RenderButton
End Sub

Private Sub UserControl_Show()
    RenderButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_caption = PropBag.ReadProperty("Caption", "CoolButton")
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    m_StyleSheet = PropBag.ReadProperty("StyleSheet", "")
    Img_Align = PropBag.ReadProperty("ImageAlign", 0)
    m_Align = PropBag.ReadProperty("CaptionAlign", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", vbMagenta)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", m_caption, "CoolButton")
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("StyleSheet", m_StyleSheet, "")
    Call PropBag.WriteProperty("ImageAlign", Img_Align, 0)
    Call PropBag.WriteProperty("CaptionAlign", m_Align, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, vbMagenta)
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = pIcon.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set pIcon.Picture = New_Picture
    Call RenderButton
    PropertyChanged "Picture"
End Property

Public Property Get Caption() As String
    Caption = m_caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    m_caption = vNewCaption
    Call RenderButton
    PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    m_ForeColor = vNewValue
    Call RenderButton
    PropertyChanged "ForeColor"
End Property

Public Property Get StyleSheet() As String
    StyleSheet = m_StyleSheet
End Property

Public Property Let StyleSheet(ByVal vNewValue As String)
    m_StyleSheet = vNewValue
    PropertyChanged "StyleSheet"
    StyleLoaded = True
End Property

Public Property Get ImageAlign() As TAlignment
    ImageAlign = Img_Align
End Property

Public Property Let ImageAlign(ByVal vNewAlign As TAlignment)
    Img_Align = vNewAlign
    Call RenderButton
    PropertyChanged "ImageAlign"
End Property

Public Property Get CaptionAlign() As TAlignment
    CaptionAlign = m_Align
End Property

Public Property Let CaptionAlign(ByVal vNewAlign As TAlignment)
    m_Align = vNewAlign
    Call RenderButton
    PropertyChanged "CaptionAlign"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call RenderButton
    PropertyChanged "Font"
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

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    Call RenderButton
    PropertyChanged "MaskColor"
End Property

