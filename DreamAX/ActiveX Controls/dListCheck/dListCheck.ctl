VERSION 5.00
Begin VB.UserControl dListCheck 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   DrawWidth       =   53
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "dListCheck.ctx":0000
   Begin VB.PictureBox PicIcons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   255
      Picture         =   "dListCheck.ctx":00FA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      Begin VB.VScrollBar vBar 
         Height          =   270
         Left            =   480
         Max             =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   90
         Width           =   270
      End
   End
End
Attribute VB_Name = "dListCheck"
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
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Enum LItemTypeA
    LItem = 0
    LHeader = 1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Item
    Item As String
    Key As String
    mItemType As LItemTypeA
    mValue As Boolean
End Type

'Text Style Consts
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4

Private m_ItemData() As Item
Private m_ListCount As Long

Private m_ItemHeight As Long
Private m_ItemWidth As Long

Private Const m_ImageSize = 16

Private m_TextAlignment As Integer

Private m_LastItem As Long
Private m_Index As Long

Private m_HeadFColor As OLE_COLOR
Private m_HeadBkColor As OLE_COLOR
Private m_SelectBkColor As OLE_COLOR
Private m_SelectFColor As OLE_COLOR
Private m_HeaderBold As Boolean
Private m_ForeColor As OLE_COLOR

Private m_Button As MouseButtonConstants

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event DblClick()
Event Click()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Selected(Index As Long)
Event Change()

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Function SelectItem(ItemY As Long, hdc As Long, oColor As OLE_COLOR)
Dim rc As RECT
Dim hBrush As Long

    With rc
        .Left = 0
        .Right = m_ItemWidth
        .Top = (ItemY + m_ItemHeight) + 1
        .Bottom = (1 + ItemY)
    End With
    
    hBrush = CreateSolidBrush(TranslateColor(oColor))
    FillRect hdc, rc, hBrush
    DeleteObject hBrush

End Function

Private Function PrintText(ItemY As Long, hdc As Long, lText As String, Align As Integer)
Dim rc As RECT
    'Prints the item Text.
    With rc
        .Left = m_ImageSize + 2
        .Right = m_ItemWidth
        .Top = (ItemY + m_ItemHeight)
        .Bottom = ItemY
    End With
    
    DrawText hdc, lText, Len(lText), rc, DT_SINGLELINE Or Align Or DT_VCENTER

End Function

Private Sub RenderItem(Item As Long)
On Error Resume Next
Dim IconX As Integer
Dim ItemY_Pos As Long
    
    'Renders the items.
    If (ListCount) > (m_LastItem) Then
        vBar.Max = (ListCount - m_LastItem)
    Else
        vBar.Max = 0
    End If
    
    ItemY_Pos = (Item - vBar.Value - 1) * m_ItemHeight - 1
    IconX = 0
    
    PicBase.FontBold = False
    
    With m_ItemData(Item)
        'Item is selected.
        If (Item = m_Index) Then
            'Item forecolor.
            PicBase.ForeColor = SelectForeColor
            'Item selection backcolor.
            SelectItem ItemY_Pos, PicBase.hdc, SelectBackColor
        Else
            'Item Forecolor
            PicBase.ForeColor = ForeColor
        End If
        'Do Item Header
        If (.mItemType = LHeader) Then
            PicBase.FontBold = BoldHeaders
            PicBase.ForeColor = HeaderForeColor
            SelectItem ItemY_Pos, PicBase.hdc, HeaderBackColor
            Call PrintText(ItemY_Pos, PicBase.hdc, .Item, 1)
        Else
            'Do Check Items
            'PicBase.FontBold = False
            'Check if the item is checked.
            If (.mValue) Then IconX = 16
            'Blt on the check image.
            TransparentBlt PicBase.hdc, 0, ItemY_Pos + (m_ImageSize \ 2) - 7, m_ImageSize, _
            m_ImageSize, PicIcons.hdc, IconX, 0, m_ImageSize, m_ImageSize, RGB(255, 0, 255)
            'Print Item
            Call PrintText(ItemY_Pos, PicBase.hdc, .Item, 0)
        End If
    End With
End Sub

Private Sub RenderLB()
On Error Resume Next
Dim sPos As Long
Dim Counter As Long

    PicBase.Cls
    'Last item offset to draw to
    m_LastItem = Fix(PicBase.ScaleHeight / m_ItemHeight)
    sPos = (vBar.Value + 1)
    'Render each of the list items
    For Counter = 1 To m_ListCount
        Call RenderItem(Counter)
    Next Counter
    
    Counter = 0
    sPos = 0
End Sub

Private Sub ResizeAll()
On Error Resume Next
    'Position the picture canvas and scrollbar,
    With PicBase
        .Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        vBar.Move (.ScaleWidth - vBar.Width), 0, vBar.Width, .ScaleHeight
        m_ItemWidth = (.ScaleWidth - vBar.Width)
    End With
End Sub

'Listbox Subs
Public Sub AddItem(Item As String, Optional Key As Variant = "", Optional ItemType As LItemTypeA = LItem)
    'INC Listcount
    m_ListCount = m_ListCount + 1
    'Resize array to hold items
    ReDim Preserve m_ItemData(ListCount)
    'Fill in Item Info
    With m_ItemData(ListCount)
        .Item = Item
        .Key = Key
        .mItemType = ItemType
    End With
    'Render the items
    Call RenderItem(ListCount)
End Sub

Public Sub Clear()
    m_ListCount = 0
    Erase m_ItemData
    PicBase.Cls
    Call RenderLB
End Sub

Public Sub Delete(Index As Long)
On Error GoTo ErrFlag:
Dim iSize As Long
    iSize = UBound(m_ItemData)
    
    'Deletes an list item
    If (iSize = 0) Then
        Call Clear
        Exit Sub
    End If
    
    If (Index > ListCount) Then
        Err.Raise 9
        Exit Sub
    End If
    
    While (iSize > Index)
        m_ItemData(Index).Item = m_ItemData(Index + 1).Item
        m_ItemData(Index).Key = m_ItemData(Index + 1).Key
        m_ItemData(Index).mItemType = m_ItemData(Index + 1).mItemType
        m_ItemData(Index).mValue = m_ItemData(Index + 1).mValue
        
        Index = Index + 1
    Wend
    
    ReDim Preserve m_ItemData(iSize - 1)
    m_ListCount = m_ListCount - 1
    iSize = 0
    
    Call RenderLB
    Exit Sub
    
ErrFlag:
If Err Then Err.Raise 9 + vbObjectError
End Sub

Private Sub PicBase_Click()
    If (m_Button = vbLeftButton) Then
        RaiseEvent Click
        RaiseEvent Selected(ListIndex)
    End If
End Sub

Private Sub UserControl_InitProperties()
    HeaderBackColor = vbButtonFace
    HeaderForeColor = vbBlack
    SelectBackColor = vbBlue
    SelectForeColor = vbWhite
    ForeColor = vbBlack
    BoldHeaders = True
End Sub

Private Sub UserControl_Resize()
    Call ResizeAll
End Sub

Private Sub UserControl_Show()
    'Item Height.
    m_ItemHeight = m_ImageSize
    Call RenderLB
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HeaderBackColor", HeaderBackColor, vbButtonFace)
    Call PropBag.WriteProperty("HeaderForeColor", HeaderForeColor, vbBlack)
    Call PropBag.WriteProperty("SelectBackColor", SelectBackColor, vbBlue)
    Call PropBag.WriteProperty("SelectForeColor", SelectForeColor, vbWhite)
    Call PropBag.WriteProperty("BoldHeaders", BoldHeaders, True)
    Call PropBag.WriteProperty("BackColor", PicBase.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", PicBase.MousePointer, 0)
    Call PropBag.WriteProperty("ForeColor", ForeColor, vbBlack)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    HeaderBackColor = PropBag.ReadProperty("HeaderBackColor", vbButtonFace)
    HeaderForeColor = PropBag.ReadProperty("HeaderForeColor", vbBlack)
    SelectBackColor = PropBag.ReadProperty("SelectBackColor", vbBlue)
    SelectForeColor = PropBag.ReadProperty("SelectForeColor", vbWhite)
    BoldHeaders = PropBag.ReadProperty("BoldHeaders", True)
    PicBase.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    PicBase.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
End Sub

Private Sub vBar_Change()
    RenderLB
End Sub

Private Sub vBar_Scroll()
    Call vBar_Change
End Sub

'Control Properties stuff
Public Property Get ListCount() As Long
    ListCount = m_ListCount
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    'Return ListItem Index
    ListIndex = m_Index
End Property

Public Property Let ListIndex(ByVal NewIndex As Long)
    m_Index = NewIndex
    Call RenderLB
    RaiseEvent Change
    PropertyChanged "ListIndex"
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
    HeaderBackColor = m_HeadBkColor
End Property

Public Property Let HeaderBackColor(ByVal NewColor As OLE_COLOR)
    m_HeadBkColor = NewColor
    Call RenderLB
    PropertyChanged "HeaderBackColor"
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
    HeaderForeColor = m_HeadFColor
End Property

Public Property Let HeaderForeColor(ByVal NewColor As OLE_COLOR)
    m_HeadFColor = NewColor
    Call RenderLB
    PropertyChanged "HeaderForeColor"
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

Public Property Get BoldHeaders() As Boolean
    BoldHeaders = m_HeaderBold
End Property

Public Property Let BoldHeaders(ByVal vNewValue As Boolean)
    m_HeaderBold = vNewValue
    Call RenderLB
    PropertyChanged "BoldHeaders"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = PicBase.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    PicBase.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get SelectedText() As String
    If (ListIndex <= 0) Then
        SelectedText = m_ItemData(1).Item
    Else
        SelectedText = m_ItemData(ListIndex).Item
    End If
End Property

Public Property Get IsHeader(Index As Long) As Boolean
    IsHeader = m_ItemData(Index).mItemType = LHeader
End Property

Public Property Let IsHeader(Index As Long, NewValue As Boolean)
    m_ItemData(Index).mItemType = Abs(NewValue)
    Call RenderLB
    PropertyChanged "IsHeader"
End Property

Public Property Get IsChecked(Index As Long) As Boolean
    IsChecked = Not m_ItemData(Index).mValue
End Property

Public Property Let IsChecked(Index As Long, Checked As Boolean)

    m_ItemData(Index).mValue = Not Checked
    
    Call RenderLB
    PropertyChanged "IsChecked"
End Property

Public Property Get Item(Index As Long) As String
    Item = m_ItemData(Index).Item
End Property

Public Property Let Item(Index As Long, NewItem As String)
    m_ItemData(Index).Item = NewItem
    Call RenderLB
    PropertyChanged "Item"
End Property

Private Sub PicBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Pre_Idx As Long
On Error Resume Next
    m_Button = Button
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button <> vbLeftButton Then Exit Sub
    'Get Prev item.
    Pre_Idx = m_Index
    'Get Item Index
    m_Index = ((Y \ m_ItemHeight) + vBar.Value) + 1
    'Check that index is not greator than the count.
    If (m_Index > ListCount) Then m_Index = Pre_Idx
    'Exit if we click on a header.
    If m_ItemData(m_Index).mItemType = LHeader Then Exit Sub
    'Check that we are of the left side of the item for the checking.
    If (X <= m_ImageSize) And (m_ItemData(m_Index).mItemType <> LHeader) Then
        m_ItemData(m_Index).mValue = (Not m_ItemData(m_Index).mValue)
    End If
    'Render Items
    Call RenderLB
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

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = PicBase.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set PicBase.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = PicBase.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    PicBase.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    Call RenderLB
    PropertyChanged "ForeColor"
End Property

