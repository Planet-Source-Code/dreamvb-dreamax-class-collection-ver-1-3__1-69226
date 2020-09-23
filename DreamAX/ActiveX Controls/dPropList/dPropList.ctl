VERSION 5.00
Begin VB.UserControl dPropList 
   Alignable       =   -1  'True
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   ToolboxBitmap   =   "dPropList.ctx":0000
   Begin VB.PictureBox pBase 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   0
      Top             =   0
      Width           =   3360
      Begin VB.PictureBox pHolder 
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   0
         ScaleHeight     =   143
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   163
         TabIndex        =   2
         Top             =   -15
         Width           =   2445
         Begin VB.TextBox TxtA 
            Height          =   315
            Index           =   0
            Left            =   1155
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.PictureBox pSplit 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1095
            MousePointer    =   9  'Size W E
            ScaleHeight     =   285
            ScaleWidth      =   30
            TabIndex        =   4
            Top             =   0
            Width           =   30
         End
         Begin VB.ComboBox CboA 
            Height          =   315
            Index           =   0
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label LblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BackColor"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   6
            Top             =   90
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.VScrollBar vBar1 
         Height          =   285
         Left            =   2460
         TabIndex        =   1
         Top             =   -15
         Width           =   255
      End
   End
End
Attribute VB_Name = "dPropList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Enum CrtType
    tTextbox = 0
    tComboBox = 1
End Enum

Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2

Event PropChanged(sKey As String, ItemProp As CrtType, ItemValue As String)
Event PropClick(sKey As String, ItemProp As CrtType)

Private Sub FlatBorder(ByVal hwnd As Long, MakeControlFlat As Boolean)
Dim TFlat As Long
 Const WS_BORDER As Long = &H800000


    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (MakeControlFlat) Then
        TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE '
    Else
        TFlat = TFlat And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    End If
    
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Sub

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
    vBar1.Height = (pBase.ScaleHeight)
    vMax = (pHolder.Height - pBase.Height)
    
    If (vMax < 0) Then vMax = 0
    vBar1.Max = vMax
    
    'Resize the controls
    For Each c In UserControl.Controls
        If (c.Name = "CboA") Or (c.Name = "TxtA") Then
            If (c.Index > 0) Then
                c.Width = (pHolder.ScaleWidth - c.Left)
            End If
        End If
    Next c
    
    pSplit.Visible = (LblA.Count > 1)
    pSplit.Height = (pHolder.ScaleHeight - 1)
    pSplit.Line (0, 0)-(0, pSplit.ScaleHeight - 1), vb3DShadow
    pSplit.Line (15, 0)-(15, pSplit.ScaleHeight - 1), vbWhite

    pSplit.Refresh
    
    Set c = Nothing
End Sub

Private Function IndexOfControl(sKey As String, Optional ItemProp As CrtType = tTextbox) As Integer
Dim c As Control
Dim Idx As Integer
Dim sTypeName As String
    
    'This function is used to return the index of a control
    
    Idx = -1
    
    If (ItemProp = tComboBox) Then sTypeName = "COMBOBOX" 'Return the index of a combobox control
    If (ItemProp = tTextbox) Then sTypeName = "TEXTBOX" ' Return the index of a textbox control
    
    For Each c In UserControl.Controls
        'Only serach the controls for TxtA and CboA
        If (c.Name = "TxtA") Or (c.Name = "CboA") Then
            'Only index's greator than zero to be checked
            If (c.Index > 0) Then
                'Check to see if the typename matches sTypeName
                If UCase(TypeName(c)) = sTypeName Then
                    'Compare the Tag with the propery Key
                    If (StrComp(sKey, c.Tag, vbTextCompare) = 0) Then
                        'Return the controls index
                        Idx = c.Index
                        'Exit loop
                        Exit For
                    End If
                End If
            End If
        End If
    Next c
    'Return the found index
    IndexOfControl = Idx
    
    'Clear up
    sTypeName = vbNullString
    Set c = Nothing
    Idx = 0
End Function

Public Sub Clear()
Dim c As Control
    'Used to unload all the control arrays.
    For Each c In UserControl.Controls
        If (c.Name = "LblA") Or (c.Name = "TxtA") Or (c.Name = "CboA") Then
            If (c.Index > 0) Then
                'Unload all controls except the first one.
                Unload c
            End If
        End If
    Next c
    Set c = Nothing
End Sub

Public Function GetPropItemValue(sKey As String, Optional ItemProp As CrtType = tTextbox) As String
Dim cIdx As Integer
On Error GoTo TErr:
    'Get the index of the control
    cIdx = IndexOfControl(sKey, ItemProp)
    'Check for a vaild return index
    If (cIdx = -1) Then
        Err.Raise 9
        Exit Function
    Else
        If (ItemProp = tTextbox) Then
            'Return textbox value.
            GetPropItemValue = TxtA(cIdx).Text
        End If
        If (ItemProp = tComboBox) Then
            'Return comboxbox value.
            GetPropItemValue = CboA(cIdx).Text
        End If
    End If
    
    'Clear up
    cIdx = 0
    Exit Function
TErr:
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub SetPropItemValue(sKey As String, lValues, Optional ItemProp As CrtType = tTextbox)
Dim cIdx As Integer
Dim Item
On Error GoTo TErr:

    'Get the index of the control
    cIdx = IndexOfControl(sKey, ItemProp)
    'Check for a vaild return index
    If (cIdx = -1) Then
        Err.Raise 9
        Exit Sub
    Else
        'We are dealing with a textbox control
        'Set the textbox's data
        If (ItemProp = tTextbox) Then
            'Set up the control to assign the text to.
            TxtA(cIdx).Text = lValues
        End If
        
        'Set the items for the combobox
        If (ItemProp = tComboBox) Then
            'Clear the combobox
            CboA(cIdx).Clear
            'Add the items from the collection to the combobox
            For Each Item In lValues
                CboA(cIdx).AddItem Item
            Next Item
            'Set the first top item
            CboA(cIdx).ListIndex = 0
        End If
    End If
    
    'Clear up
    cIdx = 0
    Item = ""
    Exit Sub
    
TErr:
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AddProp(mCaption As String, sKey As String, ItemProp As CrtType)
Dim LCount As Integer
Dim ItemTop As Long

    'Used to add the property items.
    
    LCount = LblA.Count 'Return the number of labels
    Load LblA(LCount)
    
    If (LCount = 1) Then
        'Deafult top for the first label
        ItemTop = LblA(0).Top
    Else
        'All the rest.
        ItemTop = LblA(LCount - 1).Top + LblA(LCount).Height + 8
    End If
    
    With LblA(LCount)
        .Top = ItemTop
        .Caption = mCaption
        .Visible = True
    End With

    Select Case ItemProp
        Case tTextbox
            'Text Field
            LCount = TxtA.Count
            Load TxtA(LCount)
            TxtA(LCount).Top = ItemTop - 4
            TxtA(LCount).Visible = True
            TxtA(LCount).Tag = sKey
            FlatBorder TxtA(LCount).hwnd, True
        Case tComboBox
            'ComboBox
            LCount = CboA.Count
            Load CboA(LCount)
            CboA(LCount).Top = ItemTop - 4
            CboA(LCount).Visible = True
            CboA(LCount).Tag = sKey
    End Select
    
    LCount = 0
    
    pHolder.Height = (ItemTop + 20)
    UserControl_Resize
End Sub

Private Sub CboA_Click(Index As Integer)
    RaiseEvent PropClick(CboA(Index).Tag, tComboBox)
    RaiseEvent PropChanged(CboA(Index).Tag, tComboBox, CboA(Index).Text)
End Sub

Private Sub CboA_LostFocus(Index As Integer)
    RaiseEvent PropChanged(CboA(Index).Tag, tComboBox, CboA(Index).Text)
End Sub

Private Sub TxtA_Click(Index As Integer)
    RaiseEvent PropClick(TxtA(Index).Tag, tTextbox)
End Sub

Private Sub TxtA_GotFocus(Index As Integer)
    TxtA(Index).BackColor = &HFFF9F2
End Sub

Private Sub TxtA_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = 13) Then
        RaiseEvent PropChanged(TxtA(Index).Tag, tTextbox, TxtA(Index).Text)
        KeyAscii = 0
    End If
End Sub

Private Sub TxtA_LostFocus(Index As Integer)
    TxtA(Index).BackColor = vbWhite
    RaiseEvent PropChanged(TxtA(Index).Tag, tTextbox, TxtA(Index).Text)
End Sub

Private Sub UserControl_Resize()
    Call ResizeControls
End Sub

Private Sub UserControl_Show()
    FlatBorder UserControl.hwnd, True
    Call ResizeControls
End Sub

Private Sub vBar1_Change()
    pHolder.Top = -vBar1.Value
End Sub

Private Sub vBar1_Scroll()
    vBar1_Change
End Sub

