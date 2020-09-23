VERSION 5.00
Begin VB.UserControl dXPButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ToolboxBitmap   =   "dXPButton.ctx":0000
   Begin VB.Timer ButTmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   420
      Top             =   600
   End
   Begin VB.PictureBox pSrc1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1440
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      Top             =   3675
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   75
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   3675
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   75
      Width           =   45
   End
End
Attribute VB_Name = "dXPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Private Const m_DefCaption As String = "XPButton"

Private m_Default As Boolean
Private m_Align As TCapAlign
Private m_bStyle As TButStyle
Private c_Rect As RECT

Enum TCapAlign
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum TButStyle
    XP_Homestead = 1
    XP_Metallic = 2
    XP_Blue = 3
End Enum

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()
Event DblClick()

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub Render(ButtonIndex As Integer)
On Error Resume Next
Dim X As Long
Dim Y As Long
Dim bWidth As Long
Dim bHeight As Long

    With UserControl
        .Cls
        
        X = .ScaleWidth - 3: Y = .ScaleHeight - 3
        bWidth = (X - 3): bHeight = (Y - 3)
        
        'Check if the default property is enabled
        If (m_Default) And (ButtonIndex = 0) Then ButtonIndex = 4

        If (Not .Enabled) Then ButtonIndex = 2
        
        'Draw the source button, this be used to render the main button
        BitBlt pSrc1.hdc, 0, 0, 18, 21, pSrc.hdc, (ButtonIndex * 18), 0, vbSrcCopy
        'Draws the small Cornners on the button
        StretchBlt .hdc, 0, 0, 3, 3, pSrc1.hdc, 0, 0, 3, 3, vbSrcCopy
        StretchBlt .hdc, 0, Y, 3, 3, pSrc1.hdc, 0, 18, 3, 3, vbSrcCopy
        StretchBlt .hdc, X, Y, 3, 3, pSrc1.hdc, 15, 18, 3, 3, vbSrcCopy
        StretchBlt .hdc, X, 0, 3, 3, pSrc1.hdc, 15, 0, 3, 3, vbSrcCopy
        'Draw the buttons outlines
        StretchBlt .hdc, 3, 0, bWidth, 3, pSrc1.hdc, 3, 0, 12, 3, vbSrcCopy 'Top
        StretchBlt .hdc, 0, 3, 3, bHeight, pSrc1.hdc, 0, 3, 3, 15, vbSrcCopy 'Left
        StretchBlt .hdc, X, 3, 3, bHeight, pSrc1.hdc, 15, 3, 3, 15, vbSrcCopy 'Right
        StretchBlt .hdc, 3, Y, bWidth, 3, pSrc1.hdc, 3, 18, 12, 3, vbSrcCopy 'Bottom
        'Draw the buttons body
        StretchBlt .hdc, 3, 3, bWidth, bHeight, pSrc1.hdc, 3, 3, 12, 15, vbSrcCopy
        
        'Center the caption
        c_Rect.Top = (.ScaleHeight - .TextHeight(lblCap.Caption)) \ 2
        'Caption Alignments
        If (m_Align = aLeft) Then c_Rect.Left = 3
        If (m_Align = aRight) Then c_Rect.Left = (.ScaleWidth - .TextWidth(lblCap.Caption)) - 3
        If (m_Align = aCenter) Then c_Rect.Left = (.ScaleWidth - .TextWidth(lblCap.Caption)) \ 2
        'Set the label position
        lblCap.Left = c_Rect.Left
        lblCap.Top = c_Rect.Top
        'Refresh the control
        .Refresh
    End With
    
End Sub

Private Sub LoadPicImage()
On Error Resume Next
    If (m_bStyle = 0) Then m_bStyle = XP_Homestead
    pSrc.Picture = LoadResPicture(m_bStyle, vbResBitmap)
End Sub

Private Sub ButTmr_Timer()
On Error Resume Next
Dim Pnt As POINTAPI
Dim Ret As Long

    'Get Mouse Cords
    Ret = GetCursorPos(Pnt)
    
    If (Ret) Then
        ScreenToClient UserControl.hwnd, Pnt
    End If
    'Check if we ar in th buttons bounds
    If Pnt.X < UserControl.ScaleLeft Or Pnt.Y < UserControl.ScaleTop Or _
        Pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
        Pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        'Disable the timer
        ButTmr.Enabled = False
        'Render the first button
        Call Render(0)
    End If
End Sub

Private Sub lblCap_Click()
    UserControl_Click
End Sub

Private Sub lblCap_DblClick()
    UserControl_DblClick
End Sub

Private Sub lblCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblCap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_Initialize()
    m_Default = False
    m_Align = aCenter
    m_bStyle = XP_Homestead
    Call Render(0)
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button <> vbLeftButton) Then Exit Sub
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Call Render(1)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButTmr.Enabled = True
    If (X >= 0) And (Y >= 0) And _
       (X <= UserControl.ScaleWidth) And (Y <= UserControl.ScaleHeight) Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
        If (Button = vbLeftButton) Then
            Call Render(0)
        Else
            Call Render(3)
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button <> vbLeftButton) Then Exit Sub
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Call Render(0)
End Sub

Private Sub UserControl_Resize()
    Call Render(0)
End Sub

Private Sub UserControl_Show()
    LoadPicImage
    Call Render(0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Default = PropBag.ReadProperty("Default", False)
    m_Align = PropBag.ReadProperty("Alignment", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    m_bStyle = PropBag.ReadProperty("Style", 1)
    Set lblCap.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCap.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    lblCap.Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Default", m_Default, False)
    Call PropBag.WriteProperty("Alignment", m_Align, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("Style", m_bStyle, 1)
    Call PropBag.WriteProperty("Font", lblCap.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCap.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Caption", lblCap.Caption, Ambient.DisplayName)
End Sub

Public Property Get Default() As Boolean
    Default = m_Default
End Property

Public Property Let Default(ByVal NewDefault As Boolean)
    m_Default = NewDefault
    Call Render(0)
    PropertyChanged "Default"
End Property

Public Property Get Alignment() As TCapAlign
    Alignment = m_Align
End Property

Public Property Let Alignment(ByVal NewAlign As TCapAlign)
    m_Align = NewAlign
    Call Render(0)
    PropertyChanged "Alignment"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call Render(0)
    PropertyChanged "Enabled"
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

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = TranslateColor(UserControl.BackColor)
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = TranslateColor(New_BackColor)
    Call Render(0)
    PropertyChanged "BackColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Style() As TButStyle
    Style = m_bStyle
End Property

Public Property Let Style(ByVal vStyle As TButStyle)
    m_bStyle = vStyle
    Call LoadPicImage
    Call Render(0)
    PropertyChanged "Style"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCap.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCap.Font = New_Font
    Call Render(0)
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = TranslateColor(lblCap.ForeColor)
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCap.ForeColor() = TranslateColor(New_ForeColor)
    Call Render(0)
    PropertyChanged "ForeColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCap.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCap.Caption() = New_Caption
    Call Render(0)
    PropertyChanged "Caption"
End Property


