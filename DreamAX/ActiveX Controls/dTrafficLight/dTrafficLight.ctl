VERSION 5.00
Begin VB.UserControl dTrafficLight 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "dTrafficLight.ctx":0000
   Begin VB.PictureBox PicRes 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   -135
      Picture         =   "dTrafficLight.ctx":0312
      ScaleHeight     =   930
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   3240
   End
End
Attribute VB_Name = "dTrafficLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Enum TShowLight
    LNone = 0
    LRed = 1
    LGreen = 2
    LYellow = 3
End Enum

Private m_ShowLight As TShowLight
'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub RenderLights()
Dim mWidth As Long
Dim mHeight As Long
Dim ImgXPos As Long
    
    mWidth = (PicRes.ScaleWidth \ 4) \ Screen.TwipsPerPixelX
    mHeight = (PicRes.ScaleHeight) \ Screen.TwipsPerPixelY
    
    'Draw the Lines.
    With UserControl
        .Cls
        
        Select Case m_ShowLight
            Case LNone
                ImgXPos = 3
            Case LRed
                ImgXPos = 2
            Case LGreen
                ImgXPos = 0
            Case LYellow
                ImgXPos = 1
        End Select
        
        TransparentBlt .hDC, 0, 0, mWidth, mHeight, PicRes.hDC, _
         (mWidth * ImgXPos), 0, mWidth, mHeight, RGB(255, 0, 255)
        
        .MaskPicture = .Image
        .Refresh
   End With
End Sub
Private Sub UserControl_Initialize()
    Call RenderLights
End Sub

Private Sub UserControl_InitProperties()
    m_ShowLight = LRed
End Sub

Private Sub UserControl_Resize()
 On Error Resume Next
    UserControl.Size (PicRes.ScaleWidth \ 4), PicRes.ScaleHeight
    Call RenderLights
    If Err Then Err.Clear
End Sub

Private Sub UserControl_Show()
   Call RenderLights
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ShowLight", ShowLight, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ShowLight = PropBag.ReadProperty("ShowLight", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Public Property Get ShowLight() As TShowLight
    ShowLight = m_ShowLight
End Property

Public Property Let ShowLight(ByVal NewLight As TShowLight)
    m_ShowLight = NewLight
    Call RenderLights
    PropertyChanged "ShowLight"
End Property
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

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

