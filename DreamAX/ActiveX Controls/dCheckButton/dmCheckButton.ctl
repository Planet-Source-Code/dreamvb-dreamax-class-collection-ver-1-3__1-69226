VERSION 5.00
Begin VB.UserControl dCheckButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   77
   ToolboxBitmap   =   "dmCheckButton.ctx":0000
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
Attribute VB_Name = "dCheckButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum TAlignmentA
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

Enum CButtonState
    bUp = 1
    bDown = 2
End Enum

Private bState As CButtonState
Private m_caption As String

'Control Rect
Dim c_Rect As RECT
'Image Alignment
Private Img_Align As TAlignmentA
'Caption Alignment
Private m_Align As TAlignmentA
'Check state
Private m_IsChecked As Boolean
Event AccessKeyChange()

Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_BTNHIGHLIGHT As Long = 20
'
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8

Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."

Private Sub RenderButton()
Dim CenterY As Integer
Dim xImgPos As Integer
Const ImgWidth As Integer = 16

    With UserControl
        'Used for centering the image
        SetRect c_Rect, 0, 0, .ScaleWidth, .ScaleHeight
        CenterY = (c_Rect.Bottom - ImgWidth) \ 2
        .Cls
        
        If (m_IsChecked) Then
            'Button Down state
            DrawEdge .hdc, c_Rect, BDR_SUNKENOUTER, BF_RECT
        Else
            DrawEdge .hdc, c_Rect, BDR_RAISEDINNER, BF_RECT
        End If

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
                c_Rect.Left = ImgWidth + 2
            Case aCenter
                If (Img_Align = aCenter) Then
                    c_Rect.Right = xImgPos
                End If
            Case aRight
                If (Img_Align = aRight) Then
                    c_Rect.Right = (.ScaleWidth - ImgWidth - 3)
                Else
                    c_Rect.Right = (.ScaleWidth - 3)
               End If
        End Select
        
        'add on the icon
        If (pIcon.Picture <> 0) Then
            TransparentBlt .hdc, xImgPos, CenterY, ImgWidth, ImgWidth, pIcon.hdc, 0, 0, ImgWidth, ImgWidth, UserControl.MaskColor
        Else
            If (m_Align = aLeft) Then
                c_Rect.Left = 2
            End If
        End If

        'Draw the Caption
        DrawText .hdc, m_caption, Len(m_caption), c_Rect, DT_SINGLELINE Or m_Align Or DT_VCENTER
    End With
    
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent AccessKeyChange
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    m_IsChecked = False
    m_caption = "CheckButton"
    Img_Align = aLeft
    m_Align = aLeft
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        m_IsChecked = (Not m_IsChecked)
        Call RenderButton
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
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
    Img_Align = PropBag.ReadProperty("ImageAlign", 0)
    m_Align = PropBag.ReadProperty("CaptionAlign", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", vbMagenta)
    m_IsChecked = PropBag.ReadProperty("Checked", False)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", m_caption, "CoolButton")
    Call PropBag.WriteProperty("ImageAlign", Img_Align, 0)
    Call PropBag.WriteProperty("CaptionAlign", m_Align, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, vbMagenta)
    Call PropBag.WriteProperty("Checked", m_IsChecked, False)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
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

Public Property Get ImageAlign() As TAlignmentA
    ImageAlign = Img_Align
End Property

Public Property Let ImageAlign(ByVal vNewAlign As TAlignmentA)
    Img_Align = vNewAlign
    Call RenderButton
    PropertyChanged "ImageAlign"
End Property

Public Property Get CaptionAlign() As TAlignmentA
    CaptionAlign = m_Align
End Property

Public Property Let CaptionAlign(ByVal vNewAlign As TAlignmentA)
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

Public Property Get Checked() As Boolean
    Checked = m_IsChecked
End Property

Public Property Let Checked(vNewCheck As Boolean)
    m_IsChecked = vNewCheck
    Call RenderButton
    PropertyChanged "Checked"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    Call RenderButton
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call RenderButton
    PropertyChanged "BackColor"
End Property

