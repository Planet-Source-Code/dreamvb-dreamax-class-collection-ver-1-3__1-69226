VERSION 5.00
Begin VB.UserControl dTabStrip 
   AutoRedraw      =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   132
   ToolboxBitmap   =   "dTabStrip.ctx":0000
   Begin VB.PictureBox pButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   60
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "dTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private m_TabKeys As New Collection
Private m_Captions() As String
Private m_OldIndex As Integer

Private mHeight As Long
Private mWidth As Long
'Bottom Line Color
Private m_BottomLnCol As OLE_COLOR
Private m_TabForeCol As OLE_COLOR
Private m_TabBackCol As OLE_COLOR
Private m_TabSelectCol As OLE_COLOR
Private m_TabOutLineCol As OLE_COLOR
Private m_BoldTab As Boolean

'Props
Dim m_TabIndex As Integer

Event TabChange(Index As Integer, Key As String, Caption As String)
Event Error(Number As Integer, ErrStr As String)

Private Sub RenderBottomLine()
    UserControl.Cls
    UserControl.Line (0, UserControl.ScaleHeight - 1)- _
    (UserControl.ScaleWidth, UserControl.ScaleHeight - 1), BottomLineColor
    UserControl.Refresh
End Sub

Private Sub ArrangePicBoxs(Index As Long)
    'Load Picturebox
    Load pButton(Index)
    'Set Font size
    pButton(Index - 1).Font = UserControl.Font
    'pButton(Index - 1).FontSize = UserControl.FontSize
    'Set Tab Width
    mWidth = pButton(Index - 1).TextWidth(m_Captions(Index)) + Screen.TwipsPerPixelX
    pButton(Index - 1).Width = mWidth
    'Set Tab Position
    pButton(Index).Left = pButton(Index - 1).Left + mWidth + 5
    'Show the Tab
    pButton(Index - 1).Visible = True
    'Set Tab Height
    pButton(Index - 1).Height = mHeight
    'Set Tab Top Position.
    pButton(Index - 1).Top = (UserControl.ScaleHeight - mHeight)
    'Draw the Tabs
    Call DrawTab(Index - 1)
    'm_TabIndex = -1
End Sub

Private Sub DrawTab(ByVal Index As Long, Optional IsPressed As Boolean = False)
On Error Resume Next
Dim pBox As PictureBox
Dim lnColor As OLE_COLOR
Dim bkColor As OLE_COLOR
Dim bLine As OLE_COLOR
Dim m_Caption As String

    'Get the Tab Caption.
    m_Caption = m_Captions(Index + 1)
    'Create new PictureBox
    Set pBox = pButton(Index)
    'Tab Click State
    If (IsPressed) Then
        'Down State
        bkColor = TabSelectedColor
        bLine = bkColor
        lnColor = TabOutLineColor
        'Bold selected tab.
        If (BoldSelected) Then
            pBox.FontBold = True
        Else
            pBox.FontBold = UserControl.FontBold
        End If
    Else
        'Unbold or Bold the tab.
        If (BoldSelected) Then
            pBox.FontBold = False
        Else
            pBox.FontBold = UserControl.FontBold
        End If
        'Up State
        lnColor = TabOutLineColor
        bLine = lnColor
        bkColor = TabBackColor
    End If
    
    'Tab Drawing Code.
    With pBox
        .Cls
        .ForeColor = TextColor
        'Set Tab BackColor
        .BackColor = bkColor
        'Draw the Tab Lines, Left, Top, Right
        pBox.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), lnColor, B
        'Draw the Tab Bottom Line
        pBox.Line (0, .ScaleHeight - 1)-(.ScaleWidth - 1, .ScaleHeight - 1), bLine
        'Set Tab Caption Positions
        .CurrentX = (.ScaleWidth - .TextWidth(m_Caption)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(m_Caption)) \ 2
        'Set TabCaption
        TextOut .hdc, .CurrentX, .CurrentY, m_Caption, Len(m_Caption)
        .Refresh
    End With

End Sub

Public Sub RefreshTabs(Optional EraseTabs As Boolean = False)
On Error Resume Next
Dim Count As Long
    
    If (EraseTabs) Then
        'Unload the Pictureboxes.
        For Count = pButton.Count - 1 To 1 Step -1
            Unload pButton(Count)
            pButton(Count - 1).Visible = False
        Next Count
    End If
    
    For Count = 1 To m_TabKeys.Count
        'Arrange the PicturebBoxes.
        Call ArrangePicBoxs(Count)
    Next Count
    
End Sub

Public Sub AddTab(Caption As String, Key)
Dim ButCnt As Long

    'Add New Tab Item
    m_TabKeys.Add Key
    ButCnt = m_TabKeys.Count
    'Resize the array to hold captions
    ReDim Preserve m_Captions(1 To ButCnt)
    m_Captions(ButCnt) = Caption
    'Arrange the PicturebBoxes.
    Call ArrangePicBoxs(ButCnt)
End Sub

Public Sub DeleteTab(Index As Long)
On Error Resume Next
Dim nTop As Long
Dim X As Integer

    'Remove the item from the collection.
    nTop = UBound(m_Captions)
    
    m_TabKeys.Remove Index

    For X = Index To nTop - 1
        'Shift up the Items in the array
        m_Captions(X) = m_Captions(X + 1)
    Next X
    'Resize the array
    ReDim Preserve m_Captions(1 To nTop - 1)
    
    'Update Tabs
    Call RefreshTabs(True)
    Exit Sub
End Sub

Public Sub ClearTabs()
    Set m_TabKeys = Nothing
    Erase m_Captions
    Call RefreshTabs(True)
End Sub

Private Sub pButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo TErr:
    'Set Tab State
    If (Button = vbLeftButton) Then

        If Index <> m_TabIndex Then
            Call DrawTab(Index, True)
            Call DrawTab(m_TabIndex)
        Else
            Call DrawTab(Index, True)
        End If
    
        m_TabIndex = Index
        '
        RaiseEvent TabChange(Index + 1, m_TabKeys(Index + 1), m_Captions(Index + 1))
    End If
    
    Exit Sub
TErr:
    RaiseEvent TabChange(Index, m_TabKeys(Index), m_Captions(Index))

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mHeight = UserControl.TextHeight("H") + 7
    BottomLineColor = PropBag.ReadProperty("BottomLineColor", &H998877)
    TextColor = PropBag.ReadProperty("TextColor", vbBlack)
    TabBackColor = PropBag.ReadProperty("TabBackColor", &HBBAA99)
    TabSelectedColor = PropBag.ReadProperty("TabSelectedColor", vbWhite)
    TabOutLineColor = PropBag.ReadProperty("TabOutLineColor", &H998877)
    BoldSelected = PropBag.ReadProperty("BoldSelected", False)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BottomLineColor", BottomLineColor, &H998877)
    Call PropBag.WriteProperty("TextColor", TextColor, vbBlack)
    Call PropBag.WriteProperty("TabBackColor", TabBackColor, &HBBAA99)
    Call PropBag.WriteProperty("TabSelectedColor", TabSelectedColor, vbWhite)
    Call PropBag.WriteProperty("TabOutLineColor", TabOutLineColor, &H998877)
    Call PropBag.WriteProperty("BoldSelected", BoldSelected, False)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    BottomLineColor = &H998877
    TextColor = vbBlack
    TabBackColor = &HBBAA99
    TabSelectedColor = vbWhite
    TabOutLineColor = &H998877
    BoldSelected = False
End Sub

Private Sub UserControl_Show()
    Call RenderBottomLine
End Sub

'UserControl.Width = pButton(ButCnt).Left * Screen.TwipsPerPixelX
Private Sub UserControl_Resize()
    Call RenderBottomLine
End Sub

Public Property Get TabCaption(Index As Integer) As String
    TabCaption = m_Captions(Index)
End Property

Public Property Let TabCaption(Index As Integer, ByVal NewCaption As String)
    m_Captions(Index) = NewCaption
    Call RefreshTabs(True)
    Call DrawTab(m_TabIndex, True)
    PropertyChanged "TabCaption"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call RenderBottomLine
    PropertyChanged "BackColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Let TabSelect(ByVal NewIndex As Integer)
Dim Idx As Integer
    'Minues 1 from the index
    Idx = (NewIndex - 1)
    'Select the Tab
    pButton_MouseDown Idx, 1, 0, 1, 1
    PropertyChanged "TabSelect"
End Property

Public Property Get BottomLineColor() As OLE_COLOR
    BottomLineColor = m_BottomLnCol
End Property

Public Property Let BottomLineColor(ByVal NewColor As OLE_COLOR)
    m_BottomLnCol = NewColor
    Call RenderBottomLine
    PropertyChanged "BottomLineColor"
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TabForeCol
End Property

Public Property Let TextColor(ByVal NewColor As OLE_COLOR)
    m_TabForeCol = NewColor
    Call RefreshTabs(True)
    Call DrawTab(m_TabIndex, True)
    PropertyChanged "TextColor"
End Property

Public Property Get TabBackColor() As OLE_COLOR
    TabBackColor = m_TabBackCol
End Property

Public Property Let TabBackColor(ByVal NewColor As OLE_COLOR)
    m_TabBackCol = NewColor
    Call RefreshTabs(True)
    Call DrawTab(m_TabIndex, True)
    PropertyChanged "TabBackColor"
End Property

Public Property Get TabSelectedColor() As OLE_COLOR
    TabSelectedColor = m_TabSelectCol
End Property

Public Property Let TabSelectedColor(ByVal NewColor As OLE_COLOR)
    m_TabSelectCol = NewColor
    Call RefreshTabs(True)
    Call DrawTab(m_TabIndex, True)
    PropertyChanged "TabSelectedColor"
End Property

Public Property Get TabOutLineColor() As OLE_COLOR
    TabOutLineColor = m_TabOutLineCol
End Property

Public Property Let TabOutLineColor(ByVal NewColor As OLE_COLOR)
    m_TabOutLineCol = NewColor
    Call RefreshTabs(True)
    Call DrawTab(m_TabIndex, True)
    PropertyChanged "TabOutLineColor"
End Property

Public Property Get TabCount() As Integer
    TabCount = m_TabKeys.Count
End Property


Public Property Get BoldSelected() As Boolean
    BoldSelected = m_BoldTab
End Property

Public Property Let BoldSelected(ByVal NewValue As Boolean)
    m_BoldTab = NewValue
    Call RefreshTabs(True)
    Call DrawTab(m_TabIndex, True)
    PropertyChanged "BoldSelected"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

