VERSION 5.00
Begin VB.UserControl dDotNetButton 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   79
   ToolboxBitmap   =   "dDotNetButton.ctx":0000
End
Attribute VB_Name = "dDotNetButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20

Enum TAlign0
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private m_GotFocus As Boolean
Private m_CRect As RECT
Private m_CapAlign As TAlign0
Private m_Caption As String
Private m_ShowRect As Boolean

'Event Declarations:
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub DrawBox(TRect As RECT, LineColor As Long)
    UserControl.Line (TRect.Left, TRect.Top)-(TRect.Right, TRect.Bottom), LineColor, B
End Sub

Private Sub BSetFocus()
    'Check if we have focus, and insure that m_ShowRect is enabled
    If (m_GotFocus) And (m_ShowRect) Then
        SetRect m_CRect, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
        DrawBox m_CRect, vbBlack
        SetRect m_CRect, 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3
        DrawBox m_CRect, vbWhite
    End If
End Sub

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub DrawButton(Optional bState As Boolean = False)

    With UserControl
        .Cls

        SetRect m_CRect, 0, 0, .ScaleWidth - 1, .ScaleHeight - 1
        
        If (bState) Then
            'Button Down state
            DrawBox m_CRect, vbBlack
            m_CRect.Top = m_CRect.Top + 2
            m_CRect.Left = m_CRect.Left + 2
        Else
            'Button up State
            DrawBox m_CRect, vbBlack
            SetRect m_CRect, 1, 1, .ScaleWidth - 2, .ScaleHeight - 2
            DrawBox m_CRect, vbWhite
        End If
        
        'Text Alignment
        Select Case m_CapAlign
            Case aLeft
                m_CRect.Left = 5
            Case aRight
                m_CRect.Right = (.ScaleWidth - 5)
            Case aCenter
        End Select
        'Draw the Caption
        DrawText .hdc, m_Caption, Len(m_Caption), m_CRect, DT_SINGLELINE Or m_CapAlign Or DT_VCENTER
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Caption = PropBag.ReadProperty("Caption", "FlatButton")
    ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CaptionAlignment", CaptionAlignment, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowRect, True)
End Sub

Private Sub UserControl_GotFocus()
    m_GotFocus = True
End Sub

Private Sub UserControl_InitProperties()
    CaptionAlignment = aCenter
    Caption = Ambient.DisplayName
    ShowFocusRect = True
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_LostFocus()
    m_GotFocus = False
    Call UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        Call DrawButton(True)
        'Focus
        m_GotFocus = True
        Call BSetFocus
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        Call DrawButton
        Call BSetFocus
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Call DrawButton
End Sub

Public Property Get CaptionAlignment() As TAlign0
    CaptionAlignment = m_CapAlign
End Property

Public Property Let CaptionAlignment(ByVal NewCapAlign As TAlign0)
    m_CapAlign = NewCapAlign
    Call DrawButton
    PropertyChanged "CaptionAlignment"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call DrawButton
    PropertyChanged "BackColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call DrawButton
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    Call DrawButton
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

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

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
    UserControl.Size Width, Height
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewCap As String)
    m_Caption = vNewCap
    Call DrawButton
    PropertyChanged "Caption"
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowRect
End Property

Public Property Let ShowFocusRect(ByVal vNewRect As Boolean)
    m_ShowRect = vNewRect
    Call DrawButton
    PropertyChanged "ShowFocusRect"
End Property
