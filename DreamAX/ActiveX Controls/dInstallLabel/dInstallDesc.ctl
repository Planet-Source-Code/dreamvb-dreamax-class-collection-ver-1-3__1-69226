VERSION 5.00
Begin VB.UserControl dInstallDesc 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ToolboxBitmap   =   "dInstallDesc.ctx":0000
End
Attribute VB_Name = "dInstallDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum bStyle
    bNone = 0
    bFixed = 1
End Enum

Private vLst As New Collection
'Event Declarations:
Event DblClick()
Event Click()
'Variables
Private m_BlockColor As OLE_COLOR
Private m_HintColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_AutoSize As Boolean

Public Sub NextItem()
Static iCount As Integer
    iCount = iCount + 1
    If (iCount > vLst.Count) Then iCount = 1
    Call Render(True, iCount)
End Sub

Public Sub ClearItems()
    Set vLst = Nothing
End Sub

Public Sub SetItems(DescItems As Collection)
    Set vLst = DescItems
End Sub

Private Sub Render(Optional RenderBold As Boolean = False, Optional Index As Integer)
Dim X As Long
Dim Item
On Error Resume Next
    
    With UserControl
        .Cls
        .CurrentY = 10
        For Each Item In vLst
            X = X + 1
            'Check if we can render the level bold and we on the correct index
            If (RenderBold) And (Index = X) Then
                'Turn on bold
                .FontBold = True
                .ForeColor = m_HintColor
                'Draws a small block next to the item
                UserControl.Line (2, .TextHeight("Az") + .CurrentY)-(12, .CurrentY + 2), m_BlockColor, BF
            Else
                'Turn bold off
                .ForeColor = m_ForeColor
                .FontBold = False
            End If
            'Set the items left pos
            .CurrentX = 15
            'Print the Item
            UserControl.Print Item & vbCrLf
        Next Item
        .Refresh
    End With
End Sub

Private Sub UserControl_Resize()
    Call Render
End Sub

Private Sub UserControl_Show()
    Call Render
    If (m_AutoSize) Then
        UserControl.Height = (UserControl.CurrentY) * _
        Screen.TwipsPerPixelY + (UserControl.TextHeight("Az")) + 15
    End If
    
End Sub

Private Sub UserControl_Terminate()
    Set vLst = Nothing
End Sub

Public Property Get BorderStyle() As bStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bStyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_BlockColor = PropBag.ReadProperty("BlockColor", 0)
    m_AutoSize = PropBag.ReadProperty("AutoFit", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_HintColor = PropBag.ReadProperty("HintColor", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("BlockColor", m_BlockColor, 0)
    Call PropBag.WriteProperty("AutoFit", m_AutoSize, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("HintColor", m_HintColor, 0)
End Sub

Public Property Get BlockColor() As OLE_COLOR
    BlockColor = m_BlockColor
End Property

Public Property Let BlockColor(vNewColor As OLE_COLOR)
    m_BlockColor = vNewColor
    PropertyChanged "BlockColor"
    Call Render
End Property

Public Property Get AutoFit() As Boolean
    AutoFit = m_AutoSize
End Property

Public Property Let AutoFit(vNewFit As Boolean)
    m_AutoSize = vNewFit
    PropertyChanged "AutoFit"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get HintColor() As OLE_COLOR
    HintColor = m_HintColor
End Property

Public Property Let HintColor(vNewColor As OLE_COLOR)
    m_HintColor = vNewColor
    PropertyChanged "HintColor"
End Property
