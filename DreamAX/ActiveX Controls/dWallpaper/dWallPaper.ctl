VERSION 5.00
Begin VB.UserControl dWallPaper 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   ControlContainer=   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   2910
   ToolboxBitmap   =   "dWallPaper.ctx":0000
End
Attribute VB_Name = "dWallPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum bkStyle
    Transparent = 0
    None = 1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private m_hBursh As Long
Private m_Image As Picture
Private m_CRect As RECT

Private Sub UserControl_Resize()
    Call RenderWallPaper
End Sub

Private Sub UserControl_Show()
    CreateBrush
End Sub

Private Sub RenderWallPaper()
    'This Renders the Wall Paper
    With UserControl
        .Cls
        'Get the user controls Rect
        GetClientRect UserControl.hwnd, m_CRect
        'Fill usercontrols rect with the Brush
        FillRect .hdc, m_CRect, m_hBursh
        'Used for Transparent color
        .MaskPicture = .Image
    End With
End Sub

Private Sub CreateBrush()
    'Creates the Brush that will become the controls wall Paper
    If (m_Image Is Nothing) Then
        Exit Sub
    Else
        m_hBursh = CreatePatternBrush(m_Image)
        Call RenderWallPaper
    End If
End Sub

Private Sub UserControl_Terminate()
    'Clear up
    UserControl.Cls
    Set m_Image = Nothing
    DeleteObject m_hBursh
End Sub

'Control Properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_Image = PropBag.ReadProperty("Image", Nothing)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, &HFF00FF)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Public Property Get Image() As Picture
    Set Image = m_Image
    Call CreateBrush
End Property

Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    Call RenderWallPaper
    PropertyChanged "Image"
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

Public Property Get BackStyle() As bkStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As bkStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
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

Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
    UserControl.Size Width, Height
End Sub

