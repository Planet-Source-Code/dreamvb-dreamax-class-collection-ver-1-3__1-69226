VERSION 5.00
Begin VB.UserControl dCheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "dCheckBox.ctx":0000
   Begin VB.PictureBox pSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   165
      Picture         =   "dCheckBox.ctx":00FA
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "dCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" _
        (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc _
        As Long, ByVal lParam As String, ByVal wParam As Long, _
        ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, _
        ByVal n4 As Long, ByVal un As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum TCheckState
    cUchecked = 0
    cChecked = 1
End Enum

Enum CAlign
    aLeft = 0
    aRight = 2
End Enum

Enum TCheckStyleA
    Win95 = 0
    Flat = 1
    XP = 2
    Button = 3
End Enum

Private Const m_DefCaption As String = "XPCheckBox"
Private Const m_CheckSize As Integer = 13
Private Const DST_PREFIXTEXT = &H2
Private Const DSS_NORMAL = &H0
Private Const DSS_DISABLED = &H20
'
Private Const BF_RECT As Long = &HF
'
Private Const BDR_SUNKEN95 As Long = &HA
Private Const BDR_RAISED95 As Long = &H5

Private m_Checked As TCheckState
Private m_Align As CAlign
Private m_CheckBoxS As TCheckStyleA
Private m_Caption As String
Private c_Rect As RECT
Private m_HoverIn As Boolean
Private m_HighLight As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_ShowHighLight As Boolean

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Click()
Event DblClick()
Event HoverIn()
Event HoverOut()

Private Function DrawTextA(DrawOnDC As Long, x As Long, y As Long, _
       hStr As String, tEnabled As Boolean, Clr As Long) As Long
Dim OT As Long
    ' This sub will draw text in an enabled or
    ' disabled state. The text may contain a
    ' accelerator mnemonic (&)
    ' Parameters:
    '  DrawOnDC: The DC to draw on
    '         X: Top X coordinate
    '         Y: Top Y coordinate
    '      hStr: String to print
    '  tEnabled: State to draw text in (True=Enabled, False=Disabled)
    '       Clr: Color to draw text with. Only useful if tEnabled
    '            parameter is True
    
    If DrawOnDC = 0 Then Exit Function
    
    ' Set new text color and save the old one
    OT = GetTextColor(DrawOnDC)
    SetTextColor DrawOnDC, Clr
    ' Draw the text
    DrawTextA = DrawStateText(DrawOnDC, 0&, 0&, hStr, Len(hStr), _
               x, y, 0&, 0&, DST_PREFIXTEXT Or IIf(tEnabled = True, _
               DSS_NORMAL, DSS_DISABLED))
    'Restore old text color
    SetTextColor DrawOnDC, OT
End Function

Private Sub RenderButton()
Dim x
Dim y, col

Static t As Boolean

    With UserControl
        SetRect c_Rect, 0, 0, .ScaleWidth, .ScaleHeight

        If (m_Checked) Then
            'Button Down state
            DrawEdge .hdc, c_Rect, BDR_SUNKEN95, BF_RECT
            UserControl.Line (4, 4)-(.ScaleWidth - 5, .ScaleHeight - 5), 0, B
        Else
            DrawEdge .hdc, c_Rect, BDR_RAISED95, BF_RECT
        End If
        
        
    End With

End Sub

Private Sub Render()
Dim y As Long
Dim x As Long
Dim OffsetY As Integer
Dim OffSetX As Integer
On Error Resume Next
    
    With UserControl
        If (Style = XP) Then OffSetX = 0
        If (Style = Win95) Then OffSetX = 13
        If (Style = Flat) Then OffSetX = 26
        
        If (.Height < 195) Then .Height = 195
        If (.Width < 195) Then .Width = 195
        
        .Cls
            y = (.ScaleHeight - m_CheckSize) \ 2
            'Set the rect for the caption positioning
            SetRect c_Rect, m_CheckSize, y, .ScaleWidth, .ScaleHeight
            
            If (m_Align = aLeft) Then
                x = 0
                c_Rect.Left = (m_CheckSize + 3)
            End If
            
            If (m_Align = aRight) Then
                x = (.ScaleWidth - m_CheckSize)
                c_Rect.Left = 3
            End If
            
            If (m_Checked) Then
                If (.Enabled) Then
                    OffsetY = 13
                Else
                    OffsetY = 39
                End If
            Else
                If (.Enabled) Then
                    OffsetY = 0
                Else
                    OffsetY = 26
                End If
            End If
            
            'Render Button
            If (Style = Button) Then
                RenderButton
                c_Rect.Left = (.ScaleWidth - .TextWidth(m_Caption)) \ 2
                c_Rect.Top = (.ScaleHeight - .TextHeight(m_Caption)) \ 2
            Else
                BitBlt .hdc, x, y, m_CheckSize, m_CheckSize, pSrc.hdc, OffSetX, OffsetY, vbSrcCopy
            End If
            'Check if Highting is enabled.
            If (ShowHighLight) Then
                If (m_HoverIn) Then
                    'Show highlight color.
                    .ForeColor = HighLight
                Else
                    'Normal Color
                    .ForeColor = m_ForeColor
                End If
            Else
                'Just do normal forecolor.
                .ForeColor = m_ForeColor
            End If
            
            'Draw on the caption
            DrawTextA .hdc, c_Rect.Left, c_Rect.Top, m_Caption, .Enabled, .ForeColor
        .Refresh
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_Checked = cUchecked
    m_Align = aLeft
    m_Caption = m_DefCaption
    Style = XP
    ForeColor = &H80000012
    HighLight = vbBlue
    ShowHighLight = False
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    If (Button <> vbLeftButton) Then Exit Sub
    m_Checked = (Not CBool(m_Checked))
    Call Render
End Sub

Private Sub UserControl_Resize()
    Render
End Sub

Private Sub UserControl_Show()
    Render
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Checked = PropBag.ReadProperty("Value", 0)
    m_Caption = PropBag.ReadProperty("Caption", m_DefCaption)
    m_Align = PropBag.ReadProperty("Alignment", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Style = PropBag.ReadProperty("Style", 2)
    HighLight = PropBag.ReadProperty("HighLight", vbBlue)
    ShowHighLight = PropBag.ReadProperty("ShowHighLight", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Checked, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, m_DefCaption)
    Call PropBag.WriteProperty("Alignment", m_Align, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Style", Style, 2)
    Call PropBag.WriteProperty("HighLight", HighLight, vbBlue)
    Call PropBag.WriteProperty("ShowHighLight", ShowHighLight, False)
End Sub

Public Property Get Value() As TCheckState
    Value = m_Checked
End Property

Public Property Let Value(ByVal vNewValue As TCheckState)
    m_Checked = vNewValue
    Call Render
    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    m_Caption = NewCaption
    Call Render
    PropertyChanged "Caption"
End Property

Public Property Get Alignment() As CAlign
    Alignment = m_Align
End Property

Public Property Let Alignment(ByVal NewAlign As CAlign)
    m_Align = NewAlign
    Call Render
    PropertyChanged "Caption"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call Render
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    'ForeColor = UserControl.ForeColor
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    'UserControl.ForeColor() = New_ForeColor
    Call Render
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call Render
    PropertyChanged "Font"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call Render
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

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Check if ShowHighLight is true
    If Not ShowHighLight Then
        RaiseEvent MouseMove(Button, Shift, x, y)
    Else
        'Code used for HighLight color.
        If (x < 0) Or (x > UserControl.ScaleWidth) _
        Or (y < 0) Or (y > UserControl.ScaleHeight) Then
            m_HoverIn = False
            ReleaseCapture
            Render
            RaiseEvent HoverOut
        ElseIf GetCapture() <> UserControl.hwnd Then
            m_HoverIn = True
            Render
            SetCapture UserControl.hwnd
            RaiseEvent HoverIn
        End If
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

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Style() As TCheckStyleA
    Style = m_CheckBoxS
End Property

Public Property Let Style(ByVal NewStyle As TCheckStyleA)
    m_CheckBoxS = NewStyle
    Call Render
    PropertyChanged "Style"
End Property

Public Property Get HighLight() As OLE_COLOR
    HighLight = m_HighLight
End Property

Public Property Let HighLight(ByVal NewColor As OLE_COLOR)
    m_HighLight = NewColor
    Call Render
    PropertyChanged "HighLight"
End Property

Public Property Get ShowHighLight() As Boolean
    ShowHighLight = m_ShowHighLight
End Property

Public Property Let ShowHighLight(ByVal vNewValue As Boolean)
    m_ShowHighLight = vNewValue
    Call Render
    PropertyChanged "ShowHighLight"
End Property
