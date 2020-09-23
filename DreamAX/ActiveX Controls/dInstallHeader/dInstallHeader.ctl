VERSION 5.00
Begin VB.UserControl dInstallheader 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ToolboxBitmap   =   "dInstallHeader.ctx":0000
   Begin VB.Image ImgPic 
      Height          =   720
      Left            =   5100
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   525
      Width           =   45
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "dInstallheader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_HeadLnCol As OLE_COLOR
Private Const m_defMsg = "Put your message here."
Private Const m_defCaption = "Caption"

Private Sub SetupHeader()
    With UserControl
        .Cls
        'Align the picture.
        ImgPic.Left = (.ScaleWidth - ImgPic.Width) - 2
        ImgPic.Top = (.ScaleHeight - ImgPic.Height) \ 2
        'Align the mesage caption
        lblMsg.Top = ((.ScaleHeight - lblMsg.Height) \ 2) + .TextHeight(lblMsg.Caption) - 4
        'Draw the bottom line
        UserControl.Line (0, .ScaleHeight - 2)-(.ScaleWidth, .ScaleHeight - 2), HeaderLineColor
        .Refresh
    End With
End Sub

Private Sub UserControl_InitProperties()
    HeaderLineColor = vb3DShadow
    lblCaption.Caption = m_defCaption
    lblMsg.Caption = m_defMsg
End Sub

Private Sub UserControl_Resize()
    SetupHeader
End Sub

Private Sub UserControl_Show()
    SetupHeader
End Sub

Public Property Get CaptionFont() As Font
Attribute CaptionFont.VB_Description = "Returns a Font object."
    Set CaptionFont = lblCaption.Font
End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)
    Set lblCaption.Font = New_CaptionFont
    PropertyChanged "CaptionFont"
End Property

Public Property Get MsgFont() As Font
Attribute MsgFont.VB_Description = "Returns a Font object."
    Set MsgFont = lblMsg.Font
End Property

Public Property Set MsgFont(ByVal New_MsgFont As Font)
    Set lblMsg.Font = New_MsgFont
    PropertyChanged "MsgFont"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set lblCaption.Font = PropBag.ReadProperty("CaptionFont", Ambient.Font)
    Set lblMsg.Font = PropBag.ReadProperty("MsgFont", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    HeaderLineColor = PropBag.ReadProperty("HeaderLineColor", vb3DShadow)
    lblCaption.Caption = PropBag.ReadProperty("Caption", m_defCaption)
    lblMsg.Caption = PropBag.ReadProperty("Message", m_defMsg)
    lblCaption.ForeColor = PropBag.ReadProperty("CaptionForeColor", &H80000012)
    lblMsg.ForeColor = PropBag.ReadProperty("MessageForeColor", &H80000012)
    ShowPicture = PropBag.ReadProperty("ShowPicture", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CaptionFont", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("MsgFont", lblMsg.Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("HeaderLineColor", HeaderLineColor, vb3DShadow)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, m_defCaption)
    Call PropBag.WriteProperty("Message", lblMsg.Caption, m_defMsg)
    Call PropBag.WriteProperty("CaptionForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MessageForeColor", lblMsg.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ShowPicture", ShowPicture, True)
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = ImgPic.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set ImgPic.Picture = New_Picture
    Call SetupHeader
    PropertyChanged "Picture"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call SetupHeader
    PropertyChanged "BackColor"
End Property

Public Property Get HeaderLineColor() As OLE_COLOR
    HeaderLineColor = m_HeadLnCol
End Property

Public Property Let HeaderLineColor(ByVal NewColor As OLE_COLOR)
    m_HeadLnCol = NewColor
    Call SetupHeader
    PropertyChanged "HeaderLineColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Message() As String
Attribute Message.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Message = lblMsg.Caption
End Property

Public Property Let Message(ByVal New_Message As String)
    lblMsg.Caption() = New_Message
    PropertyChanged "Message"
End Property

Public Property Get CaptionForeColor() As OLE_COLOR
Attribute CaptionForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    CaptionForeColor = lblCaption.ForeColor
End Property

Public Property Let CaptionForeColor(ByVal New_CaptionForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_CaptionForeColor
    PropertyChanged "CaptionForeColor"
End Property

Public Property Get MessageForeColor() As OLE_COLOR
Attribute MessageForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    MessageForeColor = lblMsg.ForeColor
End Property

Public Property Let MessageForeColor(ByVal New_MessageForeColor As OLE_COLOR)
    lblMsg.ForeColor() = New_MessageForeColor
    PropertyChanged "MessageForeColor"
End Property

Public Property Get ShowPicture() As Boolean
    ShowPicture = ImgPic.Visible
End Property

Public Property Let ShowPicture(ByVal vNewValue As Boolean)
    ImgPic.Visible() = vNewValue
    PropertyChanged "ShowPicture"
End Property
