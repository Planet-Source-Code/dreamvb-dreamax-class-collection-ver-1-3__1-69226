VERSION 5.00
Begin VB.UserControl dAniCursor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   37
   ToolboxBitmap   =   "dAniCursor.ctx":0000
End
Attribute VB_Name = "dAniCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function LoadCursorFromFile Lib "user32.dll" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum TFrameStyle
    fNone = 0
    fRaised = 1
    fLowered = 2
    fFrame = 3
End Enum

Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000

Private Const SS_ICON As Long = &H3&
Private Const STM_SETIMAGE As Long = &H172
Private Const IMAGE_CURSOR As Long = 2
'
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8

Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private m_Filename As String
Private m_hCursor As Long
Private m_IsBusy As Boolean
Private m_CursorHwnd As Long
Private m_PanelBorder As TFrameStyle

Private Sub DrawFrame()
Dim m_Rect As RECT
    'Function used to draw a Panel effects.
    
    With UserControl
        .Cls
        'Set up the rect
        m_Rect.Left = 0
        m_Rect.Top = 0
        m_Rect.Right = (.ScaleWidth)
        m_Rect.Bottom = (.ScaleHeight)

        If (m_PanelBorder = fRaised) Then
            DrawEdge .hdc, m_Rect, &H4, BF_RECT
        End If
    
        If (m_PanelBorder = fLowered) Then
            DrawEdge .hdc, m_Rect, &H2, BF_RECT
        End If
    
        If (m_PanelBorder = fFrame) Then
            DrawEdge .hdc, m_Rect, &H6, BF_RECT
        End If
    End With
    
End Sub

Public Sub cPlay()
Dim xPos As Long
Dim yPos As Long
    
    'Center the icon
    xPos = (UserControl.ScaleWidth - 32) \ 2
    yPos = (UserControl.ScaleHeight - 32) \ 2
    
    Call cStop
    
    m_hCursor = LoadCursorFromFile(m_Filename)

    If (m_hCursor <> 0) Then
        m_CursorHwnd = CreateWindowEx(0, "Static", "", WS_CHILD Or WS_VISIBLE Or SS_ICON, xPos, yPos, 0, 0, UserControl.hwnd, 0, App.hInstance, ByVal 0)
    End If
    
    If (m_CursorHwnd = 0) Then
        DestroyCursor m_hCursor
        Exit Sub
    Else
        SendMessage m_CursorHwnd, STM_SETIMAGE, IMAGE_CURSOR, ByVal m_hCursor
        m_IsBusy = True
    End If
End Sub

Public Sub cStop()

    If DestroyCursor(m_hCursor) Then
        m_hCursor = 0
    End If
    
    If IsWindow(m_CursorHwnd) Then
        DestroyWindow (m_CursorHwnd)
        m_CursorHwnd = 0
    End If
    
    m_IsBusy = False
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    Call DrawFrame
End Sub

Public Property Get Filename() As String
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal vNewFile As String)
    m_Filename = vNewFile
    PropertyChanged "Filename"
End Property

Private Sub UserControl_InitProperties()
    m_Filename = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Filename = PropBag.ReadProperty("Filename", "")
    m_PanelBorder = PropBag.ReadProperty("PanelBorder", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Show()
    Call DrawFrame
End Sub

Private Sub UserControl_Terminate()
    If (m_IsBusy) Then Call cStop
    m_Filename = vbNullString
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Filename", m_Filename, "")
    Call PropBag.WriteProperty("PanelBorder", m_PanelBorder, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

Public Property Get IsBusy() As Boolean
    IsBusy = m_IsBusy
End Property

Public Property Get PanelBorder() As TFrameStyle
    PanelBorder = m_PanelBorder
End Property

Public Property Let PanelBorder(ByVal vNewBorder As TFrameStyle)
    m_PanelBorder = vNewBorder
    Call DrawFrame
    PropertyChanged "PanelBorder"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call DrawFrame
    PropertyChanged "BackColor"
End Property

