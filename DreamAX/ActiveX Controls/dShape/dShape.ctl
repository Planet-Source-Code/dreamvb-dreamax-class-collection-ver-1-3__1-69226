VERSION 5.00
Begin VB.UserControl dShape 
   BackStyle       =   0  'Transparent
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "dShape.ctx":0000
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   480
   End
   Begin VB.Shape ShpA 
      BackStyle       =   1  'Opaque
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1050
   End
End
Attribute VB_Name = "dShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Click()
Event DblClick()
Event HoverIn()
Event HoverOut()

'Enums
Enum ShpBkStyle
    dTransparent = 0
    dOpaque = 1
End Enum

Enum TShapeCapAlign
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

'Variables
Private m_pBackColor As Boolean
Private m_ShapeBackColor As OLE_COLOR
Private m_Align As TShapeCapAlign

Private Sub AlignCaption()
Dim vCenter As Long
    'Align the caption
    vCenter = (UserControl.ScaleHeight - UserControl.TextHeight(lblA.Caption)) \ 2
    lblA.Top = vCenter

    Select Case m_Align
        Case aLeft
            lblA.Left = 3
        Case aCenter
            lblA.Left = (UserControl.ScaleWidth - UserControl.TextWidth(lblA.Caption)) \ 2
        Case aRight
            lblA.Left = (UserControl.ScaleWidth - UserControl.TextWidth(lblA.Caption) - 3)
    End Select
End Sub

Private Sub UserControl_InitProperties()
    Alignment = aLeft
    ParentBackColor = True
    BackColor = vbWhite
    Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    ShpA.Width = UserControl.ScaleWidth
    ShpA.Height = UserControl.ScaleHeight
    'Align caption
    Call AlignCaption
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_ShapeBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_ShapeBackColor = New_BackColor
    ParentBackColor = m_pBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BackStyle() As ShpBkStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = ShpA.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As ShpBkStyle)
    ShpA.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = ShpA.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    ShpA.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = ShpA.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    ShpA.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = ShpA.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    ShpA.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

Public Property Get DrawMode() As DrawModeConstants
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
    DrawMode = ShpA.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As DrawModeConstants)
    ShpA.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = ShpA.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    ShpA.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

Public Property Get Shape() As ShapeConstants
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
    Shape = ShpA.Shape
End Property

Public Property Let Shape(ByVal New_Shape As ShapeConstants)
    ShpA.Shape() = New_Shape
    Call UserControl_Resize
    PropertyChanged "Shape"
End Property

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

    If (X < 0) Or (X > UserControl.ScaleWidth) _
    Or (Y < 0) Or (Y > UserControl.ScaleHeight) Then
        ReleaseCapture
        'Mouse out
        RaiseEvent HoverOut
    ElseIf GetCapture() <> UserControl.hwnd Then
        SetCapture UserControl.hwnd
        'mouse in
        RaiseEvent HoverIn
    End If
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    ShpA.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    ShpA.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    ShpA.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    ShpA.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    ShpA.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    ShpA.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    ShpA.Shape = PropBag.ReadProperty("Shape", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    ShpA.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    ParentBackColor = PropBag.ReadProperty("ParentBackColor", True)
    lblA.Caption = PropBag.ReadProperty("Caption", "Label1")
    Alignment = PropBag.ReadProperty("Alignment", 0)
    Set lblA.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = lblA.Font
    lblA.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
End Sub

Private Sub UserControl_Show()
    Call AlignCaption
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", BackColor, &H80000005)
    Call PropBag.WriteProperty("BackStyle", ShpA.BackStyle, 0)
    Call PropBag.WriteProperty("BorderColor", ShpA.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BorderStyle", ShpA.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderWidth", ShpA.BorderWidth, 1)
    Call PropBag.WriteProperty("DrawMode", ShpA.DrawMode, 13)
    Call PropBag.WriteProperty("FillColor", ShpA.FillColor, &H0&)
    Call PropBag.WriteProperty("Shape", ShpA.Shape, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("FillStyle", ShpA.FillStyle, 1)
    Call PropBag.WriteProperty("ParentBackColor", ParentBackColor, True)
    Call PropBag.WriteProperty("Caption", lblA.Caption, "Label1")
    Call PropBag.WriteProperty("Alignment", Alignment, 0)
    Call PropBag.WriteProperty("Font", lblA.Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", UserControl.Font, lblA.Font)
    Call PropBag.WriteProperty("ForeColor", lblA.ForeColor, &H80000012)
End Sub

Public Property Get FillStyle() As FillStyleConstants
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = ShpA.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleConstants)
    ShpA.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

Public Property Get ParentBackColor() As Boolean
    ParentBackColor = m_pBackColor
End Property

Public Property Let ParentBackColor(ByVal NewValue As Boolean)
    m_pBackColor = NewValue
    
    If (ParentBackColor) Then
        ShpA.BackColor = UserControl.Parent.BackColor
    Else
        ShpA.BackColor = BackColor
    End If
    
    PropertyChanged "ParentBackColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblA.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblA.Caption() = New_Caption
    Call AlignCaption
    PropertyChanged "Caption"
End Property

Public Property Get Alignment() As TShapeCapAlign
    Alignment = m_Align
End Property

Public Property Let Alignment(ByVal NewAlign As TShapeCapAlign)
    m_Align = NewAlign
    Call AlignCaption
    PropertyChanged "Alignment"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblA.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblA.Font = New_Font
    Set UserControl.Font() = New_Font
    Call AlignCaption
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblA.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblA.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

