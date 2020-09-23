VERSION 5.00
Begin VB.UserControl dCtrlColorSet 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dCtrlColorSet.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dCtrlColorSet.ctx":0100
End
Attribute VB_Name = "dCtrlColorSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private m_backcolor As OLE_COLOR
Private m_Forecolor As OLE_COLOR

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Public Sub Activate()
Attribute Activate.VB_UserMemId = -550
Dim c As Control
    'Don't apply the setting if we are in design mode.
    If (UserControl.Ambient.UserMode) Then
        For Each c In Parent.Controls
            'Supported controls.
            Select Case TypeName(c)
                Case "ListBox", "TextBox", "ComboBox", _
                    "Label", "CheckBox", "OptionButton" _
                    , "DirListBox", "DriveListBox", _
                    "FileListBox", "Frame"
                    'Set control backcolor and forecolor
                    c.BackColor = TranslateColor(CtrlBackColor)
                    c.ForeColor = TranslateColor(CtrlForeColor)
                Case "CommandButton"
                    'Set control backcolor.
                    c.BackColor = CtrlBackColor
            End Select
        Next c
    End If
End Sub

Private Sub UserControl_InitProperties()
    CtrlBackColor = UserControl.Parent.BackColor
    CtrlForeColor = UserControl.Parent.ForeColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    CtrlBackColor = PropBag.ReadProperty("CtrlBackColor", vbButtonFace)
    CtrlForeColor = PropBag.ReadProperty("CtrlForeColor", vbBlack)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CtrlBackColor", CtrlBackColor, vbButtonFace)
    Call PropBag.WriteProperty("CtrlForeColor", CtrlForeColor, vbBlack)
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

Public Property Get CtrlBackColor() As OLE_COLOR
    CtrlBackColor = m_backcolor
End Property

Public Property Let CtrlBackColor(ByVal NewColorVal As OLE_COLOR)
    m_backcolor = NewColorVal
    PropertyChanged "CtrlBackColor"
End Property

Public Property Get CtrlForeColor() As OLE_COLOR
    CtrlForeColor = m_Forecolor
End Property

Public Property Let CtrlForeColor(ByVal NewColorVal As OLE_COLOR)
    m_Forecolor = NewColorVal
    PropertyChanged "CtrlForeColor"
End Property
