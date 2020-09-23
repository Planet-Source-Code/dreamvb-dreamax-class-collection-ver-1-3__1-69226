VERSION 5.00
Begin VB.UserControl dFormGradient 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dFormGradient.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dFormGradient.ctx":0173
End
Attribute VB_Name = "dFormGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Private m_StartColor As OLE_COLOR
Private m_EndColor As OLE_COLOR
Private m_Direction As GRADIENT_DIR
Private WithEvents TFrm As Form
Attribute TFrm.VB_VarHelpID = -1

'Types
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Enum GRADIENT_DIR
    Horizontal = &H0
    Vertical = &H1
End Enum

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

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

Private Sub GDI_GradientFill(hdc As Long, mRect As RECT, mStartColor As OLE_COLOR, mEndColor As OLE_COLOR, gDir As GRADIENT_DIR)
Dim gRect As GRADIENT_RECT
Dim tTV(0 To 1) As TRIVERTEX
    'Function used to paint a Gradient effect on the listbox
    
    setTriVertexColor tTV(1), TranslateColor(mEndColor)
    tTV(0).X = mRect.Left
    tTV(0).Y = mRect.top
    
    setTriVertexColor tTV(0), TranslateColor(mStartColor)
    tTV(1).X = mRect.Right
    tTV(1).Y = mRect.Bottom

    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    GradientFill hdc, tTV(0), 2, gRect, 1, gDir
    
End Sub

Private Sub TFrm_Resize()
Dim c_Rect As RECT
Dim sMode As Integer
    
    If (Not UserControl.Ambient.UserMode) Then Exit Sub
    'Get the forms scale mode
    sMode = TFrm.ScaleMode
    'Set scale mode to pixels
    TFrm.ScaleMode = vbPixels
    'Clear the dc
    TFrm.Cls
    
    With c_Rect
        .Left = 0
        .top = 0
        .Right = TFrm.ScaleWidth
        .Bottom = TFrm.ScaleHeight
    End With
    
    Call GDI_GradientFill(TFrm.hdc, c_Rect, m_StartColor, m_EndColor, m_Direction)
    'Restore the forms scalemode
    TFrm.ScaleMode = sMode
End Sub

Private Sub UserControl_InitProperties()
    m_StartColor = vbBlack
    m_EndColor = vbWhite
    m_Direction = Horizontal
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

Private Sub UserControl_Show()
    Call TFrm_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_StartColor = PropBag.ReadProperty("StartColor", vbBlack)
    m_EndColor = PropBag.ReadProperty("EndColor", vbWhite)
    m_Direction = PropBag.ReadProperty("Direction", 0)
End Sub

Private Sub UserControl_Terminate()
    Set TFrm = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("StartColor", m_StartColor, vbBlack)
    Call PropBag.WriteProperty("EndColor", m_EndColor, vbWhite)
    Call PropBag.WriteProperty("Direction", m_Direction, 0)
End Sub

Public Property Get StartColor() As OLE_COLOR
    StartColor = m_StartColor
End Property

Public Property Let StartColor(ByVal NewColor As OLE_COLOR)
    m_StartColor = NewColor
    Call TFrm_Resize
    PropertyChanged "StartColor"
End Property

Public Property Get EndColor() As OLE_COLOR
    EndColor = m_EndColor
End Property

Public Property Let EndColor(ByVal NewColor As OLE_COLOR)
    m_EndColor = NewColor
    Call TFrm_Resize
    PropertyChanged "EndColor"
End Property

Public Property Get Direction() As GRADIENT_DIR
    Direction = m_Direction
End Property

Public Property Let Direction(ByVal NewDir As GRADIENT_DIR)
    m_Direction = NewDir
    Call TFrm_Resize
    PropertyChanged "Direction"
End Property

Public Property Let FormObject(ByVal NewTForm As Form)
    Set TFrm = NewTForm
End Property
