VERSION 5.00
Begin VB.UserControl dCheckPanel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   68
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   Begin VB.CheckBox chkOp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "dCheckPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Caption As String
Private m_CaptionH As Integer
Private m_AutoSize As Boolean
Private m_CapBackColor As OLE_COLOR
Private m_LineColor As OLE_COLOR
Private m_CheckForeColor As OLE_COLOR
Private m_Align As CPanelTxtAlign
Private m_CheckAlign As CCheckBoxAlign

Enum CPanelTxtAlign
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

Enum CCheckBoxAlign
    CLeft = 0
    CRight = 1
End Enum

Private Const AbsCheckHeight As Integer = 18
'
Event Change(Index As Integer, Key, Value As Integer)

Public Sub Clear()
Dim count As Integer
    'Destroy all the objects.
    For count = 1 To chkOp.count - 1
        Unload chkOp(count)
    Next count
End Sub

Private Sub RenderPanel()
    With UserControl
        .Cls
        
        UserControl.Line (0, 0)-(.ScaleWidth - 1, CaptionHeight), CaptionBackColor, BF
        UserControl.Line (0, CaptionHeight + 1)-(.ScaleWidth, CaptionHeight + 1), CaptionLineColor
        
        Select Case CaptionAlignment
            Case aLeft
                .CurrentX = 3
            Case aCenter
                .CurrentX = (.ScaleWidth - 3 - .TextWidth(m_Caption)) \ 2
            Case aRight
                .CurrentX = (.ScaleWidth - 3) - .TextWidth(m_Caption)
        End Select
        
        .CurrentY = (CaptionHeight / 2) - .TextHeight(m_Caption) / 2
        
        UserControl.Print m_Caption
        .Refresh
    End With
End Sub

Public Sub AddCheck(Caption As String, Optional Key, Optional Value As Boolean = False)
Dim xCount As Integer
Dim xTop As Integer

    xCount = chkOp.count
    
    'Load new object
    Load chkOp(xCount)
    
    If (xCount = 1) Then
        xTop = chkOp(0).Top
    Else
        xTop = chkOp(xCount - 1).Top + AbsCheckHeight
    End If
    
    With chkOp(xCount)
        .Top = xTop
        .Caption = Caption
        .Visible = True
        .Tag = Key
        .Value = Abs(Value)
        .BackColor = BackColor
        .Alignment = CheckAlignment
    End With
    
    AutoSize = m_AutoSize
    
End Sub

Private Sub chkOp_Click(Index As Integer)
    RaiseEvent Change(Index, chkOp(Index).Tag, chkOp(Index).Value)
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    CaptionForeColor = 0
    CaptionBackColor = &H80000001
    CaptionLineColor = &HFFFFFF
    CheckForeColor = 0
    CaptionHeight = 20
    CaptionAlignment = aLeft
    Set UserControl.Font = Ambient.Font
    FontBold = True
    AutoSize = False
    CheckAlignment = CLeft
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    CaptionHeight = PropBag.ReadProperty("CaptionHeight", 20)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    CaptionForeColor = PropBag.ReadProperty("CaptionForeColor", 0)
    CaptionBackColor = PropBag.ReadProperty("CaptionBackColor", &H80000001)
    CaptionLineColor = PropBag.ReadProperty("CaptionLineColor", &HFFFFFF)
    CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", aLeft)
    CheckForeColor = PropBag.ReadProperty("CheckForeColor", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    AutoSize = PropBag.ReadProperty("AutoSize", False)
    CheckAlignment = PropBag.ReadProperty("CheckAlignment", CLeft)
End Sub

Private Sub UserControl_Resize()
    chkOp(0).Width = (UserControl.ScaleWidth - 13)
    '
    AutoSize = m_AutoSize
    Call RenderPanel
End Sub

Private Sub UserControl_Show()
    Call RenderPanel
    chkOp(0).BackColor = BackColor
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("CaptionHeight", CaptionHeight, 20)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("CaptionForeColor", CaptionForeColor, 0)
    Call PropBag.WriteProperty("CaptionBackColor", CaptionBackColor, &H80000001)
    Call PropBag.WriteProperty("CaptionLineColor", CaptionLineColor, &HFFFFFF)
    Call PropBag.WriteProperty("CaptionAlignment", CaptionAlignment, aLeft)
    Call PropBag.WriteProperty("CheckForeColor", CheckForeColor, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("AutoSize", AutoSize, False)
    Call PropBag.WriteProperty("CheckAlignment", CheckAlignment, CLeft)
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    m_Caption = NewCaption
    Call RenderPanel
    PropertyChanged "Caption"
End Property

Public Property Get CaptionHeight() As Integer
    CaptionHeight = m_CaptionH
End Property

Public Property Let CaptionHeight(ByVal NewHeight As Integer)
    m_CaptionH = NewHeight
    Call RenderPanel
    PropertyChanged "CaptionHeight"
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

Public Property Get CaptionForeColor() As OLE_COLOR
Attribute CaptionForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    CaptionForeColor = UserControl.ForeColor
End Property

Public Property Let CaptionForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    Call RenderPanel
    PropertyChanged "CaptionForeColor"
End Property

Public Property Get CaptionBackColor() As OLE_COLOR
    CaptionBackColor = m_CapBackColor
End Property

Public Property Let CaptionBackColor(ByVal New_BackColor As OLE_COLOR)
    m_CapBackColor = New_BackColor
    Call RenderPanel
    PropertyChanged "CaptionBackColor"
End Property

Public Property Get CaptionLineColor() As OLE_COLOR
    CaptionLineColor = m_LineColor
End Property

Public Property Let CaptionLineColor(ByVal NewLineColor As OLE_COLOR)
    m_LineColor = NewLineColor
    Call RenderPanel
    PropertyChanged "CaptionLineColor"
End Property

Public Property Get CaptionAlignment() As CPanelTxtAlign
    CaptionAlignment = m_Align
End Property

Public Property Let CaptionAlignment(ByVal NewAlign As CPanelTxtAlign)
    m_Align = NewAlign
    Call RenderPanel
    PropertyChanged "CaptionAlignment"
End Property

Public Property Get CheckForeColor() As OLE_COLOR
    CheckForeColor = m_CheckForeColor
End Property

Public Property Let CheckForeColor(ByVal NewForeColor As OLE_COLOR)
Dim count As Integer

    m_CheckForeColor = NewForeColor
    
    For count = 0 To chkOp.count - 1
        chkOp(count).ForeColor = NewForeColor
    Next count
    
    PropertyChanged "CheckForeColor"
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

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal NewAutoSize As Boolean)
Dim lTop As Long

    m_AutoSize = NewAutoSize
    '
    If (m_AutoSize) Then
        If (chkOp.count = 1) Then
            UserControl.Height = (m_CaptionH * Screen.TwipsPerPixelY)
        Else
            lTop = chkOp(chkOp.count - 1).Top * Screen.TwipsPerPixelY + chkOp(0).Top * _
            Screen.TwipsPerPixelY - chkOp(0).Height * Screen.TwipsPerPixelY
            UserControl.Height = lTop + chkOp(0).Height + chkOp(0).Top * 2
        End If
    End If
    
    PropertyChanged "AutoSize"
End Property

Public Property Get CheckValue(Index As Integer) As Boolean
    CheckValue = chkOp(Index).Value
End Property

Public Property Let CheckValue(Index As Integer, ByVal vNewValue As Boolean)
    chkOp(Index).Value = Abs(vNewValue)
    PropertyChanged "CheckValue"
End Property

Public Property Get CheckCaption(Index As Integer) As String
    CheckCaption = chkOp(Index).Caption
End Property

Public Property Let CheckCaption(Index As Integer, ByVal vNewCaption As String)
    chkOp(Index).Caption = vNewCaption
    PropertyChanged "CheckCaption"
End Property
'
Public Property Get CheckKey(Index As Integer) As String
    CheckKey = chkOp(Index).Tag
End Property

Public Property Let CheckKey(Index As Integer, ByVal vNewKey As String)
    chkOp(Index).Tag = vNewKey
    PropertyChanged "CheckKey"
End Property

Public Property Get CheckBold(Index As Integer) As Boolean
    CheckBold = chkOp(Index).FontBold
End Property

Public Property Let CheckBold(Index As Integer, ByVal vNewBold As Boolean)
    chkOp(Index).FontBold = vNewBold
    PropertyChanged "CheckBold"
End Property

Public Property Get CheckColor(Index As Integer) As OLE_COLOR
    CheckColor = chkOp(Index).ForeColor
End Property

Public Property Let CheckColor(Index As Integer, ByVal vNewColor As OLE_COLOR)
    chkOp(Index).ForeColor = vNewColor
    PropertyChanged "CheckColor"
End Property

Public Property Get CheckEnabled(Index As Integer) As Boolean
    CheckEnabled = chkOp(Index).Enabled
End Property

Public Property Let CheckEnabled(Index As Integer, ByVal vNewEnabled As Boolean)
    chkOp(Index).Enabled = vNewEnabled
    PropertyChanged "CheckEnabled"
End Property

Public Property Get CheckCount() As Integer
    CheckCount = (chkOp.count - 1)
End Property

Public Property Get CheckAlignment() As CCheckBoxAlign
    CheckAlignment = m_CheckAlign
End Property

Public Property Let CheckAlignment(ByVal vNewAlignment As CCheckBoxAlign)
Dim x As Integer
Dim c As CheckBox
    m_CheckAlign = vNewAlignment
    '
    For Each c In UserControl.Controls
        c.Alignment = vNewAlignment
    Next c
    
    Set c = Nothing
    
    PropertyChanged "CheckAlignment"
End Property
