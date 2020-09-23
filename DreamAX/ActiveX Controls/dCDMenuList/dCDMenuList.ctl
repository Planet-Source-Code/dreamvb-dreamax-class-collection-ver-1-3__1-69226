VERSION 5.00
Begin VB.UserControl dCDMenuList 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   136
   Begin VB.PictureBox PicBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   0
      Top             =   0
      Width           =   1305
      Begin VB.PictureBox pIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   540
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   150
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.VScrollBar vBar 
         Height          =   540
         Left            =   210
         TabIndex        =   1
         Top             =   165
         Width           =   270
      End
   End
End
Attribute VB_Name = "dCDMenuList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Item
    iCaption As String
    iText As String
    iKey As String
    iImage As Variant
End Type

'Text Style Consts
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
'Item Max size
Private Const m_ItemSize = 40
'Item Width
Private m_ItemWidth As Long
'Item Count
Private m_ListCount As Long
'Holds the item data
Private m_ItemData() As Item
'Last visable Item index.
Private m_LastItem As Long
'Item Index
Private m_Index As Long

Private m_SelectBkColor As OLE_COLOR
Private m_SelectFColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR

Private c_Rect As RECT

Private m_Button As MouseButtonConstants
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event DblClick()
Event Click()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Selected(Index As Long)
Event Change()

Private Sub PicBase_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyHome
            ListIndex = 1
        Case vbKeyEnd
            ListIndex = ListCount
        Case vbKeyDown, vbKeyRight
            If (ListIndex >= ListCount) Then Exit Sub
            ListIndex = ListIndex + 1
        Case vbKeyUp, vbKeyLeft
            If (ListIndex <= 1) Then Exit Sub
            ListIndex = ListIndex - 1
    End Select
    
End Sub

Private Sub PicBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Pre_Idx As Long
On Error Resume Next
    
    m_Button = Button
    'Check that the left button was pressed.
    If Button <> vbLeftButton Then Exit Sub
    'Get Prev item.
    Pre_Idx = ListIndex
    'Get Item Index
    ListIndex = ((Y \ m_ItemSize) + vBar.Value) + 1
    'Check that index is not greator than the count.
    If (ListIndex > ListCount) Then ListIndex = Pre_Idx
    'Render Items
    Call RenderLB
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub PicBase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PicBase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub PicBase_DblClick()
    If (m_Button = vbLeftButton) Then
        RaiseEvent DblClick
    End If
End Sub

Private Sub PicBase_Click()
    If (m_Button = vbLeftButton) Then
        RaiseEvent Click
        RaiseEvent Selected(ListIndex)
    End If
End Sub

Private Sub UserControl_InitProperties()
    SelectBackColor = vbBlue
    SelectForeColor = vbWhite
    ForeColor = vbBlack
End Sub

Private Sub UserControl_Resize()
    Call ResizeAll
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SelectBackColor", SelectBackColor, vbBlue)
    Call PropBag.WriteProperty("SelectForeColor", SelectForeColor, vbWhite)
    Call PropBag.WriteProperty("BackColor", PicBase.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", PicBase.MousePointer, 0)
    Call PropBag.WriteProperty("ForeColor", ForeColor, vbBlack)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    SelectBackColor = PropBag.ReadProperty("SelectBackColor", vbBlue)
    SelectForeColor = PropBag.ReadProperty("SelectForeColor", vbWhite)
    PicBase.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    PicBase.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
End Sub

Private Sub vBar_Change()
    RenderLB
End Sub

Private Sub vBar_GotFocus()
    PicBase.SetFocus
End Sub

Private Sub vBar_Scroll()
    Call vBar_Change
End Sub

'Control Properties stuff
Public Property Get ListCount() As Long
    ListCount = m_ListCount
End Property

Public Property Get ListIndex() As Long
    'Return ListItem Index
    ListIndex = m_Index
End Property

Public Property Let ListIndex(ByVal NewIndex As Long)
    m_Index = NewIndex
    Call RenderLB
    RaiseEvent Change
    PropertyChanged "ListIndex"
End Property

Public Property Get SelectBackColor() As OLE_COLOR
    SelectBackColor = m_SelectBkColor
End Property

Public Property Let SelectBackColor(ByVal NewColor As OLE_COLOR)
    m_SelectBkColor = NewColor
    Call RenderLB
    PropertyChanged "SelectBackColor"
End Property

Public Property Get SelectForeColor() As OLE_COLOR
    SelectForeColor = m_SelectFColor
End Property

Public Property Let SelectForeColor(ByVal NewColor As OLE_COLOR)
    m_SelectFColor = NewColor
    Call RenderLB
    PropertyChanged "SelectForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = PicBase.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    PicBase.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = PicBase.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set PicBase.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = PicBase.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    PicBase.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    Call RenderLB
    PropertyChanged "ForeColor"
End Property

Public Sub Refresh()
    PicBase.Refresh
End Sub

Public Property Get ItemKey(Index As Long) As String
On Error Resume Next
    ItemKey = m_ItemData(Index).iKey
End Property

Public Property Let ItemKey(Index As Long, NewKey As String)
    m_ItemData(Index).iKey = NewKey
End Property

'Public Listbox Subs
Public Sub AddItem(sCaption As String, sText As String, Key As Variant, nPicture)
    'INC Listcount
    m_ListCount = (m_ListCount + 1)
    'Resize array to hold items
    ReDim Preserve m_ItemData(ListCount)
    'Fill in Item Info
    With m_ItemData(ListCount)
        Set .iImage = nPicture
        .iCaption = sCaption
        .iText = sText
        .iKey = Key
    End With

    'Render the items
    Call RenderItem(ListCount)
End Sub

Public Sub Clear()
    PicBase.Cls
    m_ListCount = 0
    Erase m_ItemData
    Call RenderLB
    vBar.Visible = ListCount
End Sub

'Private Listbox Subs, Functions, Tools
Private Function SelectItem(ItemY As Long, hdc As Long, oColor As OLE_COLOR)
Dim hBrush As Long
    'Setup the rect
    SetRect c_Rect, 0, (ItemY + m_ItemSize) + 1, m_ItemWidth, ItemY + 1
    'Create the Brush
    hBrush = CreateSolidBrush(TranslateColor(oColor))
    'Fill selection area,
    FillRect hdc, c_Rect, hBrush
    'Delete brush object.
    DeleteObject hBrush
End Function

Private Function PrintText(ItemY As Long, PicBox As PictureBox, lText As String, Align As Integer, Optional TextBold As Boolean = False)
    'Set font bold prop.
    PicBox.FontBold = TextBold
    'Setup the rect
    SetRect c_Rect, (m_ItemSize + 2), (ItemY + m_ItemSize), m_ItemWidth, ItemY
    'Draw the Text
    DrawText PicBox.hdc, lText, Len(lText), c_Rect, DT_SINGLELINE Or Align Or DT_VCENTER
End Function

Private Sub RenderItem(Item As Long)
On Error Resume Next
Dim ItemY As Long

    'Render a single item.
    If (ListCount > m_LastItem) Then
        vBar.Max = (ListCount - m_LastItem)
        vBar.Enabled = True
    Else
        vBar.Max = 0
        vBar.Enabled = False
    End If
    
    ItemY = (Item - vBar.Value - 1) * (m_ItemSize - 1)
    
    With m_ItemData(Item)
        'Item is selected.
        If (Item = ListIndex) Then
            'Item forecolor.
            PicBase.ForeColor = SelectForeColor
            'Item selection backcolor.
            SelectItem ItemY, PicBase.hdc, SelectBackColor
        Else
            'Item Forecolor
            PicBase.ForeColor = ForeColor
        End If
        'Set the icon picture
        Set pIcon.Picture = .iImage
        'Draw the icon on the left side of the selected item
        TransparentBlt PicBase.hdc, 2, ItemY + (32 \ 2) - 11, 32, _
        32, pIcon.hdc, 0, 0, 32, 32, RGB(255, 0, 255)
        'Print bold Text
        Call PrintText(ItemY - 6, PicBase, .iCaption, 0, True)
        'Print normal Text
        Call PrintText(ItemY + 6, PicBase, .iText, 0)
        'Destroy the picture.
        Set pIcon.Picture = Nothing
    End With
    
    vBar.Visible = ListCount
    
End Sub

Private Sub RenderLB()
On Error Resume Next
Dim Counter As Long

    PicBase.Cls
    'Last item offset to draw to
    m_LastItem = Fix(PicBase.ScaleHeight / m_ItemSize)
    'Render each of the list items
    For Counter = (vBar.Value + 1) To m_ListCount
        Call RenderItem(Counter)
    Next Counter
    
End Sub

Private Sub ResizeAll()
On Error Resume Next
    'Position the picture canvas and scrollbar,
    With PicBase
        .Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        m_ItemWidth = (.ScaleWidth - vBar.Width)
        vBar.Move m_ItemWidth, 0, vBar.Width, .ScaleHeight
        
        vBar.Visible = ListCount
    End With
    
    RenderLB
End Sub

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

