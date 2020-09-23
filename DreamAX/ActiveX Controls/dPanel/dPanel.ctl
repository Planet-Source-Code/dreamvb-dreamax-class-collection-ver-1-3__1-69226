VERSION 5.00
Begin VB.UserControl dPanel 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "dPanel.ctx":0000
End
Attribute VB_Name = "dPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Const BF_RECT As Long = &HF

Enum TPanelBevel
    bvNone = 0
    bvLowered = 1
    bvRaised = 2
End Enum

Enum PanelCapAlign
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

Private m_BevOuter As TPanelBevel
Private m_BevInner As TPanelBevel
Private m_CapPAlign As PanelCapAlign
Private m_Caption As String

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub RenderPanel()
Dim rc1 As RECT
Dim xAlignOff As Long

    With UserControl
        .Cls
        'Outer Bevel
        SetRect rc1, 0, 0, .ScaleWidth, .ScaleHeight
        DrawEdge .hdc, rc1, (m_BevOuter * 2), BF_RECT
        'inner Bevel
        SetRect rc1, 1, 1, .ScaleWidth - 1, .ScaleHeight - 1
        DrawEdge .hdc, rc1, (m_BevInner * 2), BF_RECT
        
        'Caption Alignments
        Select Case m_CapPAlign
            Case aLeft
                xAlignOff = (rc1.Left + 1)
            Case aCenter
                xAlignOff = (rc1.Right - .TextWidth(m_Caption) - 1) \ 2
            Case aRight
                xAlignOff = (rc1.Right - 1) - .TextWidth(m_Caption) - 2
        End Select
        
        'Set Caption Alignment.
        .CurrentX = xAlignOff
        .CurrentY = (rc1.Bottom \ 2) - .TextHeight(m_Caption) \ 2
        'Print the caption.
        TextOut .hdc, .CurrentX, .CurrentY, m_Caption, Len(m_Caption)

        .Refresh
    End With
End Sub

Private Sub UserControl_InitProperties()
    BevelInner = bvNone
    BevelOuter = bvRaised
    CaptionAlign = aCenter
    Caption = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BevelInner = PropBag.ReadProperty("BevelInner", 0)
    BevelOuter = PropBag.ReadProperty("BevelOuter", 1)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    CaptionAlign = PropBag.ReadProperty("CaptionAlign", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BevelInner", BevelInner, 0)
    Call PropBag.WriteProperty("BevelOuter", BevelOuter, 1)
    Call PropBag.WriteProperty("Caption", Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("CaptionAlign", CaptionAlign, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Private Sub UserControl_Resize()
    Call RenderPanel
End Sub

Private Sub UserControl_Show()
    Call RenderPanel
End Sub

Public Property Get BevelOuter() As TPanelBevel
    BevelOuter = m_BevOuter
    Call RenderPanel
    PropertyChanged "BevelOuter"
End Property

Public Property Let BevelOuter(ByVal NewBevel As TPanelBevel)
    m_BevOuter = NewBevel
End Property

Public Property Get BevelInner() As TPanelBevel
    BevelInner = m_BevInner
End Property

Public Property Let BevelInner(ByVal NewBevel As TPanelBevel)
    m_BevInner = NewBevel
    Call RenderPanel
    PropertyChanged "BevelInner"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call RenderPanel
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    m_Caption = NewCaption
    Call RenderPanel
    PropertyChanged "Caption"
End Property

Public Property Get CaptionAlign() As PanelCapAlign
    CaptionAlign = m_CapPAlign
End Property

Public Property Let CaptionAlign(ByVal NewAlign As PanelCapAlign)
    m_CapPAlign = NewAlign
    Call RenderPanel
    PropertyChanged "CaptionAlign"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call RenderPanel
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    Call RenderPanel
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
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
    RaiseEvent MouseMove(Button, Shift, X, Y)
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

