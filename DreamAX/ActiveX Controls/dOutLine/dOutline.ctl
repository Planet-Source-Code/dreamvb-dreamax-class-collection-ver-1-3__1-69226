VERSION 5.00
Begin VB.UserControl dOutline 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ControlContainer=   -1  'True
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "dOutline.ctx":0000
   Begin VB.PictureBox PicLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   3
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox PicLine 
      Align           =   4  'Align Right
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   2
      Left            =   1200
      ScaleHeight     =   465
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   15
      Width           =   15
   End
   Begin VB.PictureBox PicLine 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox PicLine 
      Align           =   3  'Align Left
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   0
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   15
      Width           =   15
   End
End
Attribute VB_Name = "dOutline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TBorderCol
    brTop As OLE_COLOR
    brLeft As OLE_COLOR
    brRight As OLE_COLOR
    brBottom As OLE_COLOR
End Type

Enum brStyle
    brTrans = 0
    brOpaque = 1
End Enum

Private m_BorderCol As TBorderCol
'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Private Sub Render()
    PicLine(0).BackColor = m_BorderCol.brLeft
    PicLine(1).BackColor = m_BorderCol.brTop
    PicLine(2).BackColor = m_BorderCol.brRight
    PicLine(3).BackColor = m_BorderCol.brBottom
End Sub

Private Sub UserControl_InitProperties()
    m_BorderCol.brTop = 0
    m_BorderCol.brLeft = 0
    m_BorderCol.brRight = 0
    m_BorderCol.brBottom = 0
End Sub

Private Sub UserControl_Show()
    Call Render
End Sub

Public Property Get TopLine() As OLE_COLOR
    TopLine = m_BorderCol.brTop
End Property

Public Property Let TopLine(ByVal NewColor As OLE_COLOR)
    m_BorderCol.brTop = NewColor
    Call Render
    PropertyChanged "TopLine"
End Property

Public Property Get LeftLine() As OLE_COLOR
    LeftLine = m_BorderCol.brLeft
End Property

Public Property Let LeftLine(ByVal NewColor As OLE_COLOR)
    m_BorderCol.brLeft = NewColor
    Call Render
    PropertyChanged "LeftLine"
End Property

Public Property Get RightLine() As OLE_COLOR
    RightLine = m_BorderCol.brRight
End Property

Public Property Let RightLine(ByVal NewColor As OLE_COLOR)
    m_BorderCol.brRight = NewColor
    Call Render
    PropertyChanged "RightLine"
End Property


Public Property Get BottomLine() As OLE_COLOR
    BottomLine = m_BorderCol.brBottom
End Property

Public Property Let BottomLine(ByVal NewColor As OLE_COLOR)
    m_BorderCol.brBottom = NewColor
    Call Render
    PropertyChanged "BottomLine"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BorderCol.brTop = PropBag.ReadProperty("TopLine", 0)
    m_BorderCol.brLeft = PropBag.ReadProperty("LeftLine", 0)
    m_BorderCol.brRight = PropBag.ReadProperty("RightLine", 0)
    m_BorderCol.brBottom = PropBag.ReadProperty("BottomLine", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TopLine", m_BorderCol.brTop, 0)
    Call PropBag.WriteProperty("LeftLine", m_BorderCol.brLeft, 0)
    Call PropBag.WriteProperty("RightLine", m_BorderCol.brRight, 0)
    Call PropBag.WriteProperty("BottomLine", m_BorderCol.brBottom, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BackStyle() As brStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As brStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

