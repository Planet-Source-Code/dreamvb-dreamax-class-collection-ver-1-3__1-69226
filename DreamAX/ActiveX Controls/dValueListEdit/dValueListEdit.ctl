VERSION 5.00
Begin VB.UserControl dValueListEdit 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   HasDC           =   0   'False
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   Begin VB.PictureBox pBase 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2835
      Begin VB.PictureBox pHolder 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   163
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   -15
         Width           =   2445
         Begin VB.TextBox TxtA 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   0
            Left            =   1305
            TabIndex        =   3
            Top             =   15
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LblA 
            BackColor       =   &H00FFFFFF&
            Caption         =   "#0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   15
            TabIndex        =   4
            Tag             =   "0"
            Top             =   15
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.VScrollBar vBar1 
         Height          =   645
         Left            =   2460
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   -15
         Width           =   255
      End
   End
End
Attribute VB_Name = "dValueListEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mItemIndex As Integer
Private mSelectColor As OLE_COLOR

Event ItemChanged(sKey As String, Value As String)
Event ItemClick(Index As Integer)

Private Sub ResizeControls()
On Error Resume Next
Dim vMax As Integer
Dim c As Control

    'Resize eveything
    pBase.Width = UserControl.ScaleWidth
    pBase.Height = UserControl.ScaleHeight
    pHolder.Width = (pBase.ScaleWidth - vBar1.Width)
    '
    vBar1.Left = (pBase.ScaleWidth - vBar1.Width)
    vBar1.Height = (pBase.ScaleHeight) + 1
    vMax = (pHolder.Height - pBase.Height)
    
    If (vMax < 0) Then vMax = 0
    vBar1.Max = vMax
    '
    'Resize the controls
    For Each c In UserControl.Controls
        If (c.Name = "TxtA") Then
            If (c.Index > 0) Then
                c.Width = (pHolder.ScaleWidth - c.Left)
            End If
        End If
    Next c
    
    Set c = Nothing
End Sub

Private Function IndexOf(ByVal Index) As Integer
Dim Count As Integer
Dim idx As Integer
    
    idx = -1
    
    If IsNumeric(Index) Then
        idx = Index
    Else
        For Count = 0 To LblA.Count - 1
            If (TxtA(Count).Tag = Index) Then
                idx = Count
                Exit For
            End If
        Next Count
    End If
    
    IndexOf = idx

End Function

Public Sub Clear()
Dim c As Control

    'Unload Object arrays.
    For Each c In UserControl.Controls
        If (c.Name = "LblA") Or (c.Name = "TxtA") Then
            If (c.Index > 0) Then
                'Unload all objects except the first one.
                Unload c
            End If
        End If
    Next c
    
    'pHolder.Visible = False
    Set c = Nothing
    
End Sub

Public Sub AddItem(Caption As String, Optional Key As String = "", Optional Value As String = "")
Dim LCount As Integer
Dim ItemTop As Long

    LCount = LblA.Count
    
    'Load Label and Textbox Objects
    Load LblA(LCount)
    Load TxtA(LCount)
    
    If (LCount = 1) Then
        'Top for first Label
        ItemTop = LblA(0).Top
    Else
        'All the rest.
        ItemTop = LblA(LCount - 1).Top + LblA(LCount).Height + 1
    End If
    
    'Set Item Name properties.
    With LblA(LCount)
        .Top = ItemTop
        .Width = (TxtA(0).Left - .Left) - 1
        .Caption = Caption
        .ForeColor = ForeColor
        .BackColor = BackColor
        .Font = Font
        .Visible = True
    End With
    
    'Set Value Properties.
    With TxtA(LCount)
        .Top = ItemTop
        '.Tag = Key
        .Text = Value
        .Tag = Key
        .ForeColor = ForeColor
        .BackColor = BackColor
        .Font = Font
        .Visible = True
    End With
    
    LCount = 0
    
    pHolder.Height = (ItemTop + TxtA(0).Height) + 1
    UserControl_Resize
End Sub

Private Sub LblA_Click(Index As Integer)
    mItemIndex = Index
    'Highlight Value Label.
    Call HighLight(ItemIndex)
    'Send foucs to the Picturebox.
    pHolder.SetFocus
End Sub

Private Sub HighLight(Index As Integer)
Dim Count As Integer
    
    For Count = 1 To LblA.Count - 1
        LblA(Count).ForeColor = ForeColor
        LblA(Count).BackColor = BackColor
    Next Count
    
    LblA(Index).BackColor = SelectedColor
    LblA(Index).ForeColor = vbWhite
    'Raise Click Event
    RaiseEvent ItemClick(Index)
End Sub

Private Sub LblA_DblClick(Index As Integer)
    'Set focus on the Tetbox
    TxtA(Index).SelStart = 0
    TxtA(Index).SelLength = Len(TxtA(Index).Text)
    TxtA(Index).SetFocus
End Sub

Private Sub TxtA_Click(Index As Integer)
    mItemIndex = Index
    Call HighLight(ItemIndex)
End Sub

Private Sub TxtA_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = 13) Then
        RaiseEvent ItemChanged(TxtA(Index).Tag, TxtA(Index).Text)
        KeyAscii = 0
    End If
End Sub

Private Sub UserControl_InitProperties()
    BoldValueNames = True
    SelectedColor = vbHighlight
    Set UserControl.Font = Ambient.Font
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_Resize()
    Call ResizeControls
End Sub

Private Sub UserControl_Show()
    mItemIndex = 1
    Call ResizeControls
End Sub

Private Sub vBar1_Change()
    pHolder.Top = (-vBar1.Value)
End Sub

Private Sub vBar1_Scroll()
    vBar1_Change
End Sub

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    GridColor = pHolder.BackColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
    pHolder.BackColor() = New_GridColor
    PropertyChanged "GridColor"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pHolder.BackColor = PropBag.ReadProperty("GridColor", &HC0C0C0)
    BoldValueNames = PropBag.ReadProperty("BoldValueNames", True)
    pBase.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    SelectedColor = PropBag.ReadProperty("SelectedColor", vbHighlight)
    LblA(0).ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GridColor", pHolder.BackColor, &HC0C0C0)
    Call PropBag.WriteProperty("BoldValueNames", BoldValueNames, True)
    Call PropBag.WriteProperty("BackColor", pBase.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("SelectedColor", SelectedColor, vbHighlight)
    Call PropBag.WriteProperty("ForeColor", LblA(0).ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, Ambient.Font.Name)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
End Sub

Public Property Get BoldValueNames() As Boolean
    BoldValueNames = LblA(0).FontBold
End Property

Public Property Let BoldValueNames(ByVal vNewValue As Boolean)
Dim Count As Integer

    For Count = 0 To (LblA.Count - 1)
        LblA(Count).FontBold = vNewValue
    Next Count
    
    PropertyChanged "BoldValueNames"
End Property

Public Property Get ItemValue(Index) As String
Dim idx As Integer
On Error GoTo IdxErr:
    idx = IndexOf(Index)
    'Return the value
    ItemValue = TxtA(idx).Text
    
    Exit Property
IdxErr:
End Property

Public Property Let ItemValue(Index, vNewValue As String)
Dim idx As Integer
    idx = IndexOf(Index)
    'Set the value
    TxtA(idx).Text = vNewValue
End Property

Public Property Get ItemCaption(Index) As String
Dim idx As Integer
On Error GoTo IdxErr:
    idx = IndexOf(Index)
    'Return the ItemCaption
    ItemCaption = LblA(idx).Caption
    
    Exit Property
IdxErr:
End Property

Public Property Let ItemCaption(Index, vNewCaption As String)
Dim idx As Integer
    idx = IndexOf(Index)
    'Set the ItemCaption
    LblA(idx).Caption = vNewCaption
End Property

Public Property Get ItemCount() As Integer
    ItemCount = (LblA.Count - 1)
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = pBase.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Dim Count As Integer
    pBase.BackColor() = New_BackColor
    
    For Count = 1 To (LblA.Count - 1)
        LblA(Count).BackColor = New_BackColor
        TxtA(Count).BackColor = New_BackColor
    Next Count
    
    PropertyChanged "BackColor"
End Property

Public Property Get SelectedColor() As OLE_COLOR
    SelectedColor = mSelectColor
End Property

Public Property Let SelectedColor(ByVal NewColor As OLE_COLOR)
    mSelectColor = NewColor
    PropertyChanged "SelectedColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = LblA(0).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Dim Count As Integer

    LblA(0).ForeColor() = New_ForeColor
    
    For Count = 1 To (LblA.Count - 1)
        LblA(Count).ForeColor = New_ForeColor
        TxtA(Count).ForeColor = New_ForeColor
    Next Count
    
    PropertyChanged "ForeColor"
End Property

Public Property Get ItemIndex() As Integer
    ItemIndex = mItemIndex
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Dim Count As Integer

    Set UserControl.Font = New_Font
    
    For Count = 1 To (LblA.Count - 1)
        LblA(Count).Font = New_Font
        TxtA(Count).Font = New_Font
    Next Count
    
    PropertyChanged "Font"
End Property

