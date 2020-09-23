VERSION 5.00
Begin VB.UserControl dColorButton 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   330
   HasDC           =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   330
   ToolboxBitmap   =   "dColorButton.ctx":0000
   Begin VB.Image Imgdown 
      Height          =   255
      Left            =   0
      Picture         =   "dColorButton.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImgSrc 
      Height          =   255
      Left            =   0
      Picture         =   "dColorButton.ctx":0390
      Top             =   0
      Width           =   270
   End
End
Attribute VB_Name = "dColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ButtonColorA As OLE_COLOR
Private Const m_defColor = &HFF8000
'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_InitProperties()
    ButtonColor = m_defColor
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Size ImgSrc.Width, ImgSrc.Height
End Sub

Public Property Get ButtonColor() As OLE_COLOR
    ButtonColor = m_ButtonColorA
End Property

Public Property Let ButtonColor(ByVal NewColor As OLE_COLOR)
    m_ButtonColorA = NewColor
    'Set the user controls Backcolor
    UserControl.BackColor = m_ButtonColorA
    PropertyChanged "ButtonColor"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ButtonColor = PropBag.ReadProperty("ButtonColor", m_defColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    ImgSrc.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
End Sub

Private Sub UserControl_Show()
    UserControl.BackColor = ButtonColor
    '
    If (Not Enabled) Then
        UserControl.BackColor = &H80000011
    Else
        UserControl.BackColor = ButtonColor
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ButtonColor", ButtonColor, m_defColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", ImgSrc.MousePointer, vbDefault)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    
    If (Not New_Enabled) Then
        UserControl.BackColor = &H80000011
    Else
        UserControl.BackColor = ButtonColor
    End If
    
    PropertyChanged "Enabled"
End Property

Private Sub ImgSrc_Click()
    RaiseEvent Click
End Sub

Private Sub ImgSrc_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub ImgSrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        Imgdown.Visible = True
        ImgSrc.Visible = False
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = ImgSrc.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set ImgSrc.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub ImgSrc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = ImgSrc.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ImgSrc.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub ImgSrc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (Button = vbLeftButton) Then
        Imgdown.Visible = False
        ImgSrc.Visible = True
    End If
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

