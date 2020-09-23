VERSION 5.00
Begin VB.UserControl dColorPanel 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   HasDC           =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   59
   ToolboxBitmap   =   "dColorBar.ctx":0000
   Begin VB.PictureBox pPal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   75
      MousePointer    =   2  'Cross
      Picture         =   "dColorBar.ctx":0312
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   90
      Width           =   735
   End
   Begin VB.Frame Fra1 
      Height          =   675
      Left            =   75
      TabIndex        =   1
      Top             =   1515
      Width           =   735
      Begin VB.PictureBox c1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   2
         Top             =   135
         Width           =   435
      End
      Begin VB.PictureBox c2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   255
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   3
         Top             =   300
         Width           =   435
      End
      Begin VB.Image ImgArrow 
         Height          =   135
         Left            =   60
         Picture         =   "dColorBar.ctx":3D24
         Top             =   480
         Width           =   180
      End
   End
   Begin VB.Frame Fra2 
      Height          =   1440
      Left            =   75
      TabIndex        =   4
      Top             =   2115
      Width           =   735
      Begin VB.PictureBox PicSelColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   90
         ScaleHeight     =   390
         ScaleWidth      =   510
         TabIndex        =   8
         Top             =   915
         Width           =   540
         Begin VB.Line Ln1 
            Index           =   1
            X1              =   0
            X2              =   510
            Y1              =   0
            Y2              =   390
         End
         Begin VB.Line Ln1 
            Index           =   0
            X1              =   495
            X2              =   15
            Y1              =   -15
            Y2              =   390
         End
      End
      Begin VB.Label lblRGB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B .."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   105
         TabIndex        =   7
         Top             =   600
         Width           =   240
      End
      Begin VB.Label lblRGB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G .."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   375
         Width           =   255
      End
      Begin VB.Label lblRGB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R .."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   150
         Width           =   240
      End
   End
End
Attribute VB_Name = "dColorPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (ByRef pChoosecolor As ChooseColor) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type RGBTRIPLE
    rgbtBlue As Byte
    rgbtGreen As Byte
    rgbtRed As Byte
End Type

Enum TCFrameStyle
    None = 0
    Raised = 1
    Lowered = 2
End Enum

Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF

Private m_PanelBorder As TCFrameStyle
Private m_ColorDLGFullOpen As Boolean

Private Sub DrawFrame(Optional FrameStyle As TCFrameStyle)
Dim Cols(3) As OLE_COLOR
    'Function used draw a Panel effect around the control
    
    UserControl.Cls
    If FrameStyle = None Then
        Exit Sub
    End If
    
    If FrameStyle = Raised Then
        Cols(0) = vbWhite
        Cols(1) = Cols(0)
        Cols(2) = &H80000010
        Cols(3) = Cols(2)
    End If
    
    If FrameStyle = Lowered Then
        Cols(0) = &H80000010
        Cols(1) = Cols(0)
        Cols(2) = vbWhite
        Cols(3) = Cols(2)
    End If
    
    
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, 0), Cols(0) 'Top
    UserControl.Line (0, 0)-(0, UserControl.ScaleHeight - 1), Cols(1) 'Left
    UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), Cols(2) 'bottom
    UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), Cols(3) 'right
End Sub

Private Function GetColorFromDLG(Optional InitColor As OLE_COLOR, Optional cFullOpen As Boolean = True) As Long
Dim cc As ChooseColor
Dim m_InitColor As Long
Dim aColorRef(15) As Long
Dim Cnt As Integer
Dim j As RGBTRIPLE
    
    Call LongToRGB(InitColor, j)
    
    'Fill in the custom colors with shaded color
    For Cnt = 240 To 15 Step -15
        aColorRef((Cnt \ 15) - 1) = RGB(j.rgbtRed + Cnt, j.rgbtGreen + Cnt, j.rgbtBlue + Cnt)
    Next Cnt
    
    ' Translate the initial OLE color to a long value
    If (InitColor <> 0) And OleTranslateColor(InitColor, 0, m_InitColor) Then
        m_InitColor = CLR_INVALID
    End If
    
    'Fill ChooseColor Type
    With cc
        .lStructSize = Len(cc)
        .hwndOwner = UserControl.hWnd
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = m_InitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(cFullOpen, CC_FULLOPEN, 0)
        'Show the color Dialogbox
        If ChooseColor(cc) Then
            'Return choosen color
            GetColorFromDLG = .rgbResult
        Else
            'Cancel button was pressed by user.
            GetColorFromDLG = -1
        End If
    End With
    'Clear up
    Cnt = 0
    m_InitColor = 0
    'Free up used types
    ZeroMemory cc, Len(cc)
    ZeroMemory j, Len(j)

End Function

Private Sub LongToRGB(LngColor As Long, RgbType As RGBTRIPLE)
On Error Resume Next
    'Convert Long Color To RGB
    RgbType.rgbtRed = (LngColor Mod 256)
    RgbType.rgbtGreen = ((LngColor And &HFF00) / 256) Mod 256
    RgbType.rgbtBlue = ((LngColor And &HFF0000) / 65536)
End Sub

Private Sub c1_Click()
Dim cRet As Long
    cRet = GetColorFromDLG(c1.BackColor, m_ColorDLGFullOpen)
    If (cRet = -1) Then Exit Sub 'Cancel button was pressed
    'Set the color
    c1.BackColor = cRet
End Sub

Private Sub c2_Click()
Dim cRet As Long
    cRet = GetColorFromDLG(c2.BackColor, m_ColorDLGFullOpen)
    If (cRet = -1) Then Exit Sub 'Cancel button was pressed
    'Set the color
    c2.BackColor = cRet
End Sub

Private Sub ImgArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A As OLE_COLOR
Dim B As OLE_COLOR
    
    A = c1.BackColor
    B = c2.BackColor
    
    If (Button = vbLeftButton) Then
        c1.BackColor = B
        c2.BackColor = A
    End If
End Sub

Private Sub pPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pPal_MouseMove Button, 0, X, Y
End Sub

Private Sub pPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m_InRect As Boolean
    'Make sure we stay within the pallates hight and width values
    m_InRect = (X < 2) Or (X > 45) Or (Y < 2) Or (Y > 96)
    'Hide or show the crossed line
    Ln1(0).Visible = m_InRect
    Ln1(1).Visible = m_InRect
    
    If (m_InRect) Then
        PicSelColor.BackColor = vbButtonShadow
        'Set the defaulr RGB Lables
        lblRGB(0).Caption = "R .."
        lblRGB(1).Caption = "G .."
        lblRGB(2).Caption = "B .."
        Exit Sub
    End If
    
    'Keep in the pallate rect size
    If (X < 2) Then X = 2
    If (X > 45) Then X = 45
    If (Y < 2) Then Y = 2
    If (Y > 96) Then Y = 96
    
    'Show the selected color
    PicSelColor.BackColor = pPal.Point(X, Y)
    'This sets color1
    If (Button = vbLeftButton) Then
        c1.BackColor = PicSelColor.BackColor
    End If
    'This sets color2
    If (Button = vbRightButton) Then
        c2.BackColor = PicSelColor.BackColor
    End If
    'Display the RGB values
    Call UpdateRGBLabels(PicSelColor.BackColor)
End Sub

Private Sub UpdateRGBLabels(This As Long)
Dim TRgb As RGBTRIPLE
    Call LongToRGB(This, TRgb)
    'Update the labels
    lblRGB(0).Caption = "R " & TRgb.rgbtRed
    lblRGB(1).Caption = "G " & TRgb.rgbtGreen
    lblRGB(2).Caption = "B " & TRgb.rgbtBlue
End Sub

Public Property Get ForegroundColor() As OLE_COLOR
Attribute ForegroundColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ForegroundColor = c1.BackColor
End Property

Public Property Let ForegroundColor(ByVal New_ForegroundColor As OLE_COLOR)
    c1.BackColor() = New_ForegroundColor
    PropertyChanged "ForegroundColor"
End Property

Public Property Get BackgroundColor() As OLE_COLOR
Attribute BackgroundColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackgroundColor = c2.BackColor
End Property

Public Property Let BackgroundColor(ByVal New_BackgroundColor As OLE_COLOR)
    c2.BackColor() = New_BackgroundColor
    PropertyChanged "BackgroundColor"
End Property

Private Sub UserControl_InitProperties()
    m_PanelBorder = Raised
    OpenFullColorDLG = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    c1.BackColor = PropBag.ReadProperty("ForegroundColor", &H0&)
    c2.BackColor = PropBag.ReadProperty("BackgroundColor", &H80000005)
    m_PanelBorder = PropBag.ReadProperty("PanelBorder", 1)
    UserControl.BackColor = PropBag.ReadProperty("PanelBackColor", &H8000000F)
    m_ColorDLGFullOpen = PropBag.ReadProperty("OpenFullColorDLG", True)
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 885, 3645
End Sub

Private Sub UserControl_Show()
    Call DrawFrame(m_PanelBorder)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForegroundColor", c1.BackColor, &H0&)
    Call PropBag.WriteProperty("BackgroundColor", c2.BackColor, &H80000005)
    Call PropBag.WriteProperty("PanelBorder", m_PanelBorder, 1)
    Call PropBag.WriteProperty("PanelBackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("OpenFullColorDLG", m_ColorDLGFullOpen, True)
End Sub

Public Property Get PanelBorder() As TCFrameStyle
    PanelBorder = m_PanelBorder
End Property

Public Property Let PanelBorder(ByVal vNewBorder As TCFrameStyle)
    m_PanelBorder = vNewBorder
    Call DrawFrame(m_PanelBorder)
    PropertyChanged "PanelBorder"
End Property

Public Property Get PanelBackColor() As OLE_COLOR
Attribute PanelBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    PanelBackColor = UserControl.BackColor
End Property

Public Property Let PanelBackColor(ByVal New_PanelBackColor As OLE_COLOR)
    UserControl.BackColor() = New_PanelBackColor
    PropertyChanged "PanelBackColor"
End Property

Public Property Get OpenFullColorDLG() As Boolean
    OpenFullColorDLG = m_ColorDLGFullOpen
End Property

Public Property Let OpenFullColorDLG(ByVal vNewValue As Boolean)
    m_ColorDLGFullOpen = vNewValue
    PropertyChanged "OpenFullColorDLG"
End Property
