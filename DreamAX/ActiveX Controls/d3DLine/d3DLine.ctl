VERSION 5.00
Begin VB.UserControl d3DLine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   6
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "d3DLine.ctx":0000
End
Attribute VB_Name = "d3DLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Enum TLineDir
    dHorizontal = 0
    dVertical = 1
End Enum

Private m_ColorA As OLE_COLOR
Private m_ColorB As OLE_COLOR
Private m_LineDir As TLineDir

Public Sub GuiLineTo(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, LineColor As Long)
Dim hPen As Long
    'Create a soild pen
    hPen = CreatePen(0, 1, LineColor)
    DeleteObject SelectObject(hdc, hPen)
    
    If (X1 >= 0) Then MoveToEx hdc, X1, Y1, 0
    'Draw the line
    LineTo hdc, X2, Y2
    
End Sub

Private Sub DrawLine()
    'Draw the Lines.
    With UserControl
        .Cls
        'Draw Horizontal
        If (m_LineDir = dHorizontal) Then
            GuiLineTo .hdc, 0, 0, .ScaleWidth, 0, GDI_TranslateColor(m_ColorB)
            GuiLineTo .hdc, 0, 1, .ScaleWidth, 1, GDI_TranslateColor(m_ColorA)
        End If
        'Draw Vertical
        If (m_LineDir = dVertical) Then
            GuiLineTo .hdc, 0, 0, 0, .ScaleHeight, GDI_TranslateColor(m_ColorB)
            GuiLineTo .hdc, 1, 0, 1, .ScaleHeight, GDI_TranslateColor(m_ColorA)
        End If
        
        .MaskPicture = UserControl.Image
        .Refresh
   End With
End Sub

Private Sub UserControl_Initialize()
    Call DrawLine
End Sub

Private Sub UserControl_InitProperties()
    Color1 = vbWhite
    Color2 = &H80000010
    Direction = dHorizontal
End Sub

Private Sub UserControl_Resize()
 On Error Resume Next
    DrawLine
    If Err Then Err.Clear
End Sub

Private Function GDI_TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, GDI_TranslateColor) Then
        GDI_TranslateColor = &HFFFF&
    End If
End Function

Public Property Get Color1() As OLE_COLOR
    Color1 = m_ColorA
End Property

Public Property Let Color1(ByVal vNewValue As OLE_COLOR)
    m_ColorA = vNewValue
    Call DrawLine
    PropertyChanged "Color1"
End Property

Public Property Get Color2() As OLE_COLOR
    Color2 = m_ColorB
End Property

Public Property Let Color2(ByVal vNewValue As OLE_COLOR)
    m_ColorB = vNewValue
    Call DrawLine
    PropertyChanged "Color2"
End Property

Private Sub UserControl_Show()
    DrawLine
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color1", Color1, vbWhite)
    Call PropBag.WriteProperty("Color2", Color2, &H80000010)
    Call PropBag.WriteProperty("Direction", Direction, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ColorA = PropBag.ReadProperty("Color1", vbWhite)
    m_ColorB = PropBag.ReadProperty("Color2", &H80000010)
    Direction = PropBag.ReadProperty("Direction", 0)
End Sub

Public Property Get Direction() As TLineDir
    Direction = m_LineDir
End Property

Public Property Let Direction(ByVal NewDir As TLineDir)
    m_LineDir = NewDir
    Call DrawLine
    PropertyChanged "Direction"
End Property
