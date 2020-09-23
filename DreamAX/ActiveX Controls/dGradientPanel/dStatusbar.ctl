VERSION 5.00
Begin VB.UserControl dGradientPanel 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   ControlContainer=   -1  'True
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   115
   ToolboxBitmap   =   "dStatusbar.ctx":0000
End
Attribute VB_Name = "dGradientPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private m_StartColor As OLE_COLOR
Private m_EndColor As OLE_COLOR
Private m_Direction As GRADIENT_DIR1

Private m_Caption As String
Private m_CapLeft As Long
Private m_CapTop As Long

'Types
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Enum GRADIENT_DIR1
    Horizontal = &H0
    Vertical = &H1
End Enum

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type
'Event Declarations:
Event Resize()


Private Sub RenderGPanel()
Dim rc As RECT
    
    rc.Left = 0
    rc.top = 0
    rc.Right = UserControl.ScaleWidth
    rc.Bottom = UserControl.ScaleHeight
    
    GDI_GradientFill UserControl.hdc, rc, m_StartColor, m_EndColor, m_Direction
    TextOut UserControl.hdc, m_CapLeft, m_CapTop, m_Caption, Len(m_Caption)
    UserControl.Refresh
    
End Sub

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub setTriVertexColor(tTV As TRIVERTEX, oColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    lRed = (oColor And &HFF&) * &H100&
    lGreen = (oColor And &HFF00&)
    lBlue = (oColor And &HFF0000) \ &H100&
    
    setTriVertexColorComponent tTV.Red, lRed
    setTriVertexColorComponent tTV.Green, lGreen
    setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef oColor As Integer, ByVal lComponent As Long)
    If (lComponent And &H8000&) = &H8000& Then
        oColor = (lComponent And &H7F00&)
        oColor = oColor Or &H8000
    Else
        oColor = lComponent
    End If
End Sub

Private Sub GDI_GradientFill(hdc As Long, mRect As RECT, mStartColor As OLE_COLOR, mEndColor As OLE_COLOR, gDir As GRADIENT_DIR1)
Dim gRect As GRADIENT_RECT
Dim tTV(0 To 1) As TRIVERTEX
    'Function used to paint a Gradient effect on the listbox
    
    setTriVertexColor tTV(1), TranslateColor(mEndColor)
    tTV(0).x = mRect.Left
    tTV(0).y = mRect.top
    
    setTriVertexColor tTV(0), TranslateColor(mStartColor)
    tTV(1).x = mRect.Right
    tTV(1).y = mRect.Bottom

    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    GradientFill hdc, tTV(0), 2, gRect, 1, gDir
    
End Sub

Private Sub UserControl_InitProperties()
    StartColor = vbBlack
    EndColor = vbWhite
    Direction = Horizontal
    LabelText = Ambient.DisplayName
    LabelTop = 10
    LabelLeft = 10
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    StartColor = PropBag.ReadProperty("StartColor", vbBlack)
    EndColor = PropBag.ReadProperty("EndColor", vbWhite)
    Direction = PropBag.ReadProperty("Direction", 0)
    LabelText = PropBag.ReadProperty("LabelText", Ambient.DisplayName)
    LabelTop = PropBag.ReadProperty("LabelTop", 10)
    LabelLeft = PropBag.ReadProperty("LabelLeft", 10)
    Set UserControl.Font = PropBag.ReadProperty("LabelFont", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("LabelForeColor", &H80000012)
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    Call RenderGPanel
End Sub

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("StartColor", StartColor, vbBlack)
    Call PropBag.WriteProperty("EndColor", EndColor, vbWhite)
    Call PropBag.WriteProperty("Direction", Direction, 0)
    Call PropBag.WriteProperty("LabelText", LabelText, Ambient.DisplayName)
    Call PropBag.WriteProperty("LabelTop", LabelTop, 10)
    Call PropBag.WriteProperty("LabelLeft", LabelLeft, 10)
    Call PropBag.WriteProperty("LabelFont", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("LabelForeColor", UserControl.ForeColor, &H80000012)
End Sub

Public Property Get StartColor() As OLE_COLOR
    StartColor = m_StartColor
End Property

Public Property Let StartColor(ByVal NewColor As OLE_COLOR)
    m_StartColor = NewColor
    Call RenderGPanel
    PropertyChanged "StartColor"
End Property

Public Property Get EndColor() As OLE_COLOR
    EndColor = m_EndColor
End Property

Public Property Let EndColor(ByVal NewColor As OLE_COLOR)
    m_EndColor = NewColor
    Call RenderGPanel
    PropertyChanged "EndColor"
End Property

Public Property Get Direction() As GRADIENT_DIR1
    Direction = m_Direction
End Property

Public Property Let Direction(ByVal NewDir As GRADIENT_DIR1)
    m_Direction = NewDir
    Call RenderGPanel
    PropertyChanged "Direction"
End Property

Public Property Get LabelText() As String
    LabelText = m_Caption
End Property

Public Property Let LabelText(ByVal NewText As String)
    m_Caption = NewText
    Call RenderGPanel
    PropertyChanged "LabelText"
End Property

Public Property Get LabelTop() As Long
    LabelTop = m_CapTop
End Property

Public Property Let LabelTop(ByVal vNewValue As Long)
    m_CapTop = vNewValue
    Call RenderGPanel
    PropertyChanged "LabelTop"
End Property

Public Property Get LabelLeft() As Long
    LabelLeft = m_CapLeft
End Property

Public Property Let LabelLeft(ByVal vNewValue As Long)
    m_CapLeft = vNewValue
    Call RenderGPanel
    PropertyChanged "LabelLeft"
End Property

Public Property Get LabelFont() As Font
Attribute LabelFont.VB_Description = "Returns a Font object."
    Set LabelFont = UserControl.Font
End Property

Public Property Set LabelFont(ByVal New_LabelFont As Font)
    Set UserControl.Font = New_LabelFont
    Call RenderGPanel
    PropertyChanged "LabelFont"
End Property

Public Property Get LabelForeColor() As OLE_COLOR
Attribute LabelForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    LabelForeColor = UserControl.ForeColor
End Property

Public Property Let LabelForeColor(ByVal New_LabelForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_LabelForeColor
    Call RenderGPanel
    PropertyChanged "LabelForeColor"
End Property

