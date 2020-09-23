VERSION 5.00
Begin VB.UserControl dLinkLabel 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   HasDC           =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   1635
   ToolboxBitmap   =   "dLinkLabel.ctx":0000
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1140
   End
End
Attribute VB_Name = "dLinkLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private m_LinkUrl As String
Private m_HoverCol As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_ActiveColor As OLE_COLOR
Private m_TextColor As OLE_COLOR
Private m_VisitedColor As OLE_COLOR
Private m_UrlOpenSate As VbAppWinStyle
Private m_HoverState As Boolean
Private m_MouseDown As Boolean
Private m_ShowUnderLine As Boolean
Private m_underline As Boolean
Private m_MousePointer As MouseP
Private m_IsVisited As Boolean

Dim hCur As Long

Enum MouseP
    IDC_WAIT = 32514&
    IDC_UPARROW = 32516&
    IDC_SIZEWE = 32644&
    IDC_SIZENWSE = 32642&
    IDC_SIZENS = 32645&
    IDC_SIZENESW = 32643&
    IDC_SIZEALL = 32646&
    IDC_SIZE = 32640&
    IDC_NO = 32648&
    IDC_IBEAM = 32513&
    IDC_CROSS = 32515&
    IDC_ARROW = 32512&
    IDC_HAND = 32649&
End Enum

Event HoverIn()
Event HoverOut()

Event MouseDown(shift As Integer, X As Single, Y As Single)
Event MouseMove(shift As Integer, X As Single, Y As Single)
Event MouseUp(shift As Integer, X As Single, Y As Single)

Private Sub ExecuteUrl()
    If Len(Trim(m_LinkUrl)) <> 0 Then
        ShellExecute 0, "open", m_LinkUrl, vbNullString, vbNullString, m_UrlOpenSate
    End If
End Sub

Private Sub SetLink()
    'Set up the mouse pointer for the control
    Call SetCursorHand(m_MousePointer)

    If (m_HoverState) Then
        'Check if show underline is enabled
        If Not ShowUnderLine Then
            lblA.Font.Underline = m_underline
        Else
            lblA.Font.Underline = ShowUnderLine
        End If
        'Set the hover color
        lblA.ForeColor = m_HoverCol
        RaiseEvent HoverIn
    Else
        lblA.Font.Underline = m_underline
        'Check if Url has been Visited
        If (m_IsVisited) Then
            'Set text color toVisited color
            lblA.ForeColor = m_VisitedColor
        Else
            'Not Visited, so set normal color
            lblA.ForeColor = m_TextColor
        End If
        m_MouseDown = False
        RaiseEvent HoverOut
    End If
    
    'Check if the mouse button was pressed
    If (m_MouseDown) Then
        'Set the Textcolor to the active color
        lblA.ForeColor = m_ActiveColor
        'Turn on Visted
        m_IsVisited = True
    End If

End Sub

Private Sub SetCursorHand(dCurType As Long)
    'Load the cursor from the system
    hCur = LoadCursor(0, dCurType)
    
    If (hCur) Then
        'Set the cursor
        SetCursor hCur
    End If
End Sub

Private Sub lblA_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, shift, X, Y
End Sub

Private Sub lblA_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, shift, X, Y
End Sub

Private Sub lblA_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, shift, X, Y
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    ShowUnderLine = True
    IsVisited = False
    HoverColor = vbBlue
    TextColor = vbBlack
    ActiveColor = vbRed
    VisitedColor = &H970080
    lblA.Font.Name = "Verdana"
    lblA.Font.Size = 8
    MousePointer = IDC_ARROW
    UrlOpenSate = vbNormalFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    If (Button <> vbLeftButton) Then Exit Sub
    m_MouseDown = True
    UserControl_MouseMove Button, shift, X, Y
    RaiseEvent MouseDown(shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(shift, X, Y)

    If (X < 0) Or (X > UserControl.ScaleWidth) _
    Or (Y < 0) Or (Y > UserControl.ScaleHeight) Then
        ReleaseCapture
        m_HoverState = False
    ElseIf GetCapture() <> UserControl.hwnd Then
        m_HoverState = True
        SetCapture UserControl.hwnd
    End If
    
    Call SetLink
End Sub

Private Sub UserControl_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
    If (Button <> vbLeftButton) Then Exit Sub
    'Execure the url
    Call ExecuteUrl
    m_MouseDown = False
    UserControl_MouseMove Button, shift, X, Y
    RaiseEvent MouseUp(shift, X, Y)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Size lblA.Width, lblA.Height
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblA.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblA.Caption() = New_Caption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblA.Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Url = PropBag.ReadProperty("Url", "")
    HoverColor = PropBag.ReadProperty("HoverColor", vbBlue)
    TextColor = PropBag.ReadProperty("TextColor", vbBlack)
    ActiveColor = PropBag.ReadProperty("ActiveColor", vbRed)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblA.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblA.Font = PropBag.ReadProperty("Font", Ambient.Font)
    ShowUnderLine = PropBag.ReadProperty("ShowUnderLine", True)
    MousePointer = PropBag.ReadProperty("MousePointer", IDC_ARROW)
    VisitedColor = PropBag.ReadProperty("VisitedColor", &H970080)
    m_IsVisited = PropBag.ReadProperty("IsVisited", False)
    UrlOpenSate = PropBag.ReadProperty("UrlOpenSate", vbNormalFocus)
End Sub

Private Sub UserControl_Show()
    m_underline = lblA.Font.Underline
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    Call DestroyCursor(hCur)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblA.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Url", Url, "")
    Call PropBag.WriteProperty("HoverColor", HoverColor, vbBlue)
    Call PropBag.WriteProperty("TextColor", TextColor, vbBlack)
    Call PropBag.WriteProperty("ActiveColor", ActiveColor, vbRed)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", lblA.Enabled, True)
    Call PropBag.WriteProperty("Font", lblA.Font, Ambient.Font)
    Call PropBag.WriteProperty("ShowUnderLine", ShowUnderLine, True)
    Call PropBag.WriteProperty("MousePointer", MousePointer, IDC_ARROW)
    Call PropBag.WriteProperty("VisitedColor", VisitedColor, &H970080)
    Call PropBag.WriteProperty("IsVisited", IsVisited, False)
    Call PropBag.WriteProperty("UrlOpenSate", UrlOpenSate, vbNormalFocus)
End Sub

Public Property Get Url() As String
    Url = m_LinkUrl
End Property

Public Property Let Url(ByVal NewUrl As String)
    m_LinkUrl = NewUrl
    PropertyChanged "Url"
End Property

Public Property Get HoverColor() As OLE_COLOR
    HoverColor = m_HoverCol
End Property

Public Property Let HoverColor(ByVal NewHoverC As OLE_COLOR)
    m_HoverCol = NewHoverC
    PropertyChanged "HoverColor"
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal NewTextC As OLE_COLOR)
    m_TextColor = NewTextC
    Call SetLink
    PropertyChanged "TextColor"
End Property

Public Property Get ActiveColor() As OLE_COLOR
    ActiveColor = m_ActiveColor
End Property

Public Property Let ActiveColor(ByVal NewActiveC As OLE_COLOR)
    m_ActiveColor = NewActiveC
    PropertyChanged "NewActiveC"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = lblA.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lblA.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblA.Font
    UserControl_Resize
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblA.Font = New_Font
    m_underline = lblA.Font.Underline
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get ShowUnderLine() As Boolean
    ShowUnderLine = m_ShowUnderLine
End Property

Public Property Let ShowUnderLine(ByVal vNewShowU As Boolean)
    m_ShowUnderLine = vNewShowU
    PropertyChanged "ShowUnderLine"
End Property

Public Property Get MousePointer() As MouseP
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal NewMouseP As MouseP)
    m_MousePointer = NewMouseP
    PropertyChanged "MousePointer"
End Property

Public Property Get VisitedColor() As OLE_COLOR
    VisitedColor = m_VisitedColor
End Property

Public Property Let VisitedColor(ByVal NewVisC As OLE_COLOR)
    m_VisitedColor = NewVisC
    PropertyChanged "VisitedColor"
End Property

Public Property Get IsVisited() As Boolean
    IsVisited = m_IsVisited
End Property

Public Property Let IsVisited(ByVal NewVis As Boolean)
Dim tmp As Boolean
    m_IsVisited = NewVis
    Call SetLink
    PropertyChanged "IsVisited"
End Property

Public Property Get UrlOpenSate() As VbAppWinStyle
    UrlOpenSate = m_UrlOpenSate
End Property

Public Property Let UrlOpenSate(ByVal vNewWinState As VbAppWinStyle)
    m_UrlOpenSate = vNewWinState
    PropertyChanged "UrlOpenSate"
End Property

