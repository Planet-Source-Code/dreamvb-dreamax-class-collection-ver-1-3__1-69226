VERSION 5.00
Begin VB.UserControl dStatusbar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   76
   ToolboxBitmap   =   "dStatusbar.ctx":0000
   Begin VB.PictureBox pGripSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   135
      Picture         =   "dStatusbar.ctx":00FA
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "dStatusbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Enum sBarStyle
    bNone = 0
    bRaised = 1
    bLowered = 2
    bFrame = 3
End Enum

Enum sGripStyle
    gsDefault = 0
    gsNewLook = 1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private m_BarStyle As sBarStyle
Private m_GripStyle As sGripStyle
Private m_SimpleText As String
Private m_ForeColor As OLE_COLOR

'Flat Buton Style.
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BF_RECT As Long = &HF

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()
Event DblClick()

Private Sub DrawLoweredBevel()
Dim rc As RECT
Dim mStyle As Long
Dim xOff As Integer
Dim yOff As Integer
Dim gStyle As Integer

    With UserControl
        xOff = 1
        rc.Top = 0
        rc.Bottom = .ScaleHeight
        rc.Left = 0
        rc.Right = .ScaleWidth
        'Clear the control
        .Cls
        'Bar Styles
        Select Case m_BarStyle
            Case bNone
            Case bRaised
                mStyle = BDR_RAISEDINNER
            Case bLowered
                mStyle = BDR_SUNKENOUTER
            Case bFrame
                xOff = 2
                mStyle = 6
        End Select
        
        'Grip Styles
        If (GripStyle = gsNewLook) Then
            gStyle = 12
        Else
            gStyle = 0
        End If
        
        'Draw the statusbar.
        DrawEdge .hDC, rc, mStyle, BF_RECT
        
        'Enable/Disabled code
        If (Not Enabled) Then
            .ForeColor = &H80000011
        Else
            .ForeColor = m_ForeColor
        End If
        
        'Caption offset
        yOff = (.ScaleHeight - .TextHeight(m_SimpleText)) \ 2
        TextOut .hDC, 3, yOff, m_SimpleText, Len(m_SimpleText)
        'Position the resize grip.
        TransparentBlt .hDC, (rc.Right - 12) - xOff, (rc.Bottom - 12) - xOff, 12, 12, pGripSrc.hDC, gStyle, 0, 12, 12, RGB(255, 0, 255)
        'Update Drawing.
        .Refresh
    End With
    
End Sub

Private Sub UserControl_InitProperties()
    BarStyle = bNone
    GripStyle = gsDefault
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BarStyle = PropBag.ReadProperty("BarStyle", 0)
    SimpleText = PropBag.ReadProperty("SimpleText", "")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    ForeColor = PropBag.ReadProperty("ForeColor", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    GripStyle = PropBag.ReadProperty("GripStyle", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BarStyle", BarStyle, 0)
    Call PropBag.WriteProperty("SimpleText", SimpleText, "")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", ForeColor, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("GripStyle", GripStyle, 0)
End Sub

Private Sub UserControl_Resize()
    DrawLoweredBevel
End Sub

Public Property Get BarStyle() As sBarStyle
    BarStyle = m_BarStyle
End Property

Public Property Let BarStyle(ByVal NewStyle As sBarStyle)
    m_BarStyle = NewStyle
    Call DrawLoweredBevel
    PropertyChanged "BarStyle"
End Property

Public Property Get SimpleText() As String
    SimpleText = m_SimpleText
End Property

Public Property Let SimpleText(ByVal NewText As String)
    m_SimpleText = NewText
    Call DrawLoweredBevel
    PropertyChanged "BarStyle"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call DrawLoweredBevel
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call DrawLoweredBevel
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    Call DrawLoweredBevel
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call DrawLoweredBevel
    PropertyChanged "BackColor"
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

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get GripStyle() As sGripStyle
    GripStyle = m_GripStyle
End Property

Public Property Let GripStyle(ByVal NewStyle As sGripStyle)
    m_GripStyle = NewStyle
    Call DrawLoweredBevel
    PropertyChanged "GripStyle"
End Property
