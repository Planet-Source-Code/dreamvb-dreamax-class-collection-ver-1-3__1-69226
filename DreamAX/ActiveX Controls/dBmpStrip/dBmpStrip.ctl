VERSION 5.00
Begin VB.UserControl dBmpStrip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   HasDC           =   0   'False
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ToolboxBitmap   =   "dBmpStrip.ctx":0000
   Begin VB.PictureBox PixSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   2145
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   0
      Top             =   495
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "dBmpStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Enum FRAME_DIR
    Horizontal = 0
    Vertical = 1
End Enum

Enum FrameBStyleD
    bNone = 0
    bFixed = 1
End Enum

Private m_Frames As Long
Private m_FrameW As Long
Private m_FrameH As Long
Private m_FrameDir As FRAME_DIR
Private m_FrameIndex As Long

Private Sub RenderFrame()
Dim xOff As Long
Dim yOff As Long
Dim ImgX As Long
Dim ImgY As Long

On Error Resume Next

    With UserControl
        .Cls
            xOff = (.ScaleWidth - FrameWidth) \ 2
            yOff = (.ScaleHeight - FrameHeight) \ 2
            '
            
            If (m_FrameIndex < 1) Then
                m_FrameIndex = 1
            End If
            
            If (m_FrameIndex > Frames) Then
                m_FrameIndex = Frames
            End If
            
            If (m_FrameDir = Horizontal) Then
                ImgX = (m_FrameIndex * FrameWidth) - FrameWidth
                ImgY = 0
            Else
                ImgX = 0
                ImgY = (m_FrameIndex * FrameHeight) - FrameHeight
            End If
        
            BitBlt .hDC, xOff, yOff, FrameWidth, FrameHeight _
            , PixSrc.hDC, ImgX, ImgY, vbSrcCopy
            
        .Refresh
        
        .MaskPicture = .Image
    End With
End Sub

Public Property Get SrcPicture() As Picture
Attribute SrcPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set SrcPicture = PixSrc.Picture
End Property

Public Property Set SrcPicture(ByVal New_SrcPicture As Picture)
    Set PixSrc.Picture = New_SrcPicture
    Call RenderFrame
    PropertyChanged "SrcPicture"
End Property

Private Sub UserControl_InitProperties()
    Frames = 0
    FrameWidth = 0
    FrameHeight = 0
    FrameIndex = 1
    FrameDirection = Horizontal
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set SrcPicture = PropBag.ReadProperty("SrcPicture", Nothing)
    Frames = PropBag.ReadProperty("Frames", 0)
    FrameWidth = PropBag.ReadProperty("FrameWidth", 0)
    FrameHeight = PropBag.ReadProperty("FrameHeight", 0)
    FrameDirection = PropBag.ReadProperty("FrameDirection", 0)
    FrameIndex = PropBag.ReadProperty("FrameIndex", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

Private Sub UserControl_Resize()
    Call RenderFrame
End Sub

Private Sub UserControl_Show()
   Call RenderFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SrcPicture", SrcPicture, Nothing)
    Call PropBag.WriteProperty("Frames", Frames, 0)
    Call PropBag.WriteProperty("FrameWidth", FrameWidth, 0)
    Call PropBag.WriteProperty("FrameHeight", FrameHeight, 0)
    Call PropBag.WriteProperty("FrameDirection", FrameDirection, 0)
    Call PropBag.WriteProperty("FrameIndex", FrameIndex, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
End Sub

Public Property Get Frames() As Integer
    Frames = m_Frames
End Property

Public Property Let Frames(ByVal NewFCount As Integer)
    m_Frames = NewFCount
    PropertyChanged "Property"
End Property

Public Property Get FrameWidth() As Long
    FrameWidth = m_FrameW
End Property

Public Property Let FrameWidth(ByVal NewWidth As Long)
    m_FrameW = NewWidth
    Call RenderFrame
    PropertyChanged "FrameWidth"
End Property

Public Property Get FrameHeight() As Long
    FrameHeight = m_FrameH
End Property

Public Property Let FrameHeight(ByVal NewHeight As Long)
    m_FrameH = NewHeight
    Call RenderFrame
    PropertyChanged "FrameHeight"
End Property

Public Property Get FrameDirection() As FRAME_DIR
    FrameDirection = m_FrameDir
End Property

Public Property Let FrameDirection(ByVal NewfDir As FRAME_DIR)
    m_FrameDir = NewfDir
    Call RenderFrame
    PropertyChanged "FrameDirection"
End Property

Public Property Get FrameIndex() As Long
    FrameIndex = m_FrameIndex
End Property

Public Property Let FrameIndex(ByVal vNewIndex As Long)
    m_FrameIndex = vNewIndex
    Call RenderFrame
    PropertyChanged "FrameIndex"
End Property

Public Property Get BorderStyle() As FrameBStyleD
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As FrameBStyleD)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

