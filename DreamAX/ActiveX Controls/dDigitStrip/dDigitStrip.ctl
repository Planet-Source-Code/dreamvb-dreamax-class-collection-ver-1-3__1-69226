VERSION 5.00
Begin VB.UserControl dDigitStrip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   975
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   65
   ToolboxBitmap   =   "dDigitStrip.ctx":0000
   Begin VB.PictureBox PicSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "dDigitStrip.ctx":0532
      ScaleHeight     =   315
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "dDigitStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Const DigitWidth As Integer = 16
Private Const DigitHeight As Integer = 21

Private mValue As Long
Private mAutoSize As Boolean
Private mDigitBStyle As DigitBStyle
Private mDigitColor As OLE_COLOR
Private mDigitDimmedColor As OLE_COLOR

Enum DigitBStyle
    bNone = 0
    bFixedSingle = 1
End Enum

Private Sub SetColor(ByVal X As Integer, ByVal Y As Integer, ByVal oColor As OLE_COLOR)
Dim Ret As Long
Dim lColor As Long

    With PicSrc
        .FillStyle = vbFSSolid
        .FillColor = oColor
        'Locate color to replace
        lColor = GetPixel(.hdc, X, Y)
        Ret = ExtFloodFill(.hdc, X, Y, lColor, 1)
    End With
End Sub

Private Sub CreateDigit(Digit As Integer)
    'Reset the picturebox.
    PicSrc.Cls
    PicSrc.BackColor = BackColor
    
    Call SetColor(3, 2, mDigitDimmedColor) 'Top
    Call SetColor(2, 3, mDigitDimmedColor) 'TopLeft
    Call SetColor(13, 3, mDigitDimmedColor) 'TopRight
    Call SetColor(4, 10, mDigitDimmedColor) 'Center
    Call SetColor(2, 11, mDigitDimmedColor) 'BottomLeft
    Call SetColor(11, 12, mDigitDimmedColor) 'BottomRight
    Call SetColor(3, 18, mDigitDimmedColor) 'Bottom
    '
    Select Case Digit
        Case 0
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(2, 3, mDigitColor) 'TopLeft
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(2, 11, mDigitColor) 'BottomLeft
            Call SetColor(11, 12, mDigitColor) 'BottomRight
            Call SetColor(3, 18, mDigitColor) 'Bottom
        Case 1
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(11, 12, mDigitColor) 'BottomRight
        Case 2
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(4, 10, mDigitColor) 'Center
            Call SetColor(2, 11, mDigitColor) 'BottomLeft
            Call SetColor(3, 18, mDigitColor) 'Bottom
        Case 3
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(4, 10, mDigitColor) 'Center
            Call SetColor(11, 12, mDigitColor) 'BottomRight
            Call SetColor(3, 18, mDigitColor) 'Bottom
        Case 4
            Call SetColor(2, 3, mDigitColor) 'TopLeft
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(4, 10, mDigitColor) 'Center
            Call SetColor(11, 12, mDigitColor) 'BottomRight
        Case 5
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(2, 3, mDigitColor) 'TopLeft
            Call SetColor(4, 10, mDigitColor) 'Center
            Call SetColor(11, 12, mDigitColor) 'BottomRight
            Call SetColor(3, 18, mDigitColor) 'Bottom
        Case 6
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(2, 3, mDigitColor) 'TopLeft
            Call SetColor(4, 10, mDigitColor) 'Center
            Call SetColor(2, 11, mDigitColor) 'BottomLeft
            Call SetColor(11, 12, mDigitColor) 'BottomRight
            Call SetColor(3, 18, mDigitColor) 'Bottom
        Case 7
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(11, 12, mDigitColor) 'BottomRight
        Case 8
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(2, 3, mDigitColor) 'TopLeft
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(4, 10, mDigitColor) 'Center
            Call SetColor(2, 11, mDigitColor) 'BottomLeft
            Call SetColor(11, 12, mDigitColor) 'BottomRight
            Call SetColor(3, 18, mDigitColor) 'Bottom
        Case 9
            Call SetColor(3, 2, mDigitColor) 'Top
            Call SetColor(2, 3, mDigitColor) 'TopLeft
            Call SetColor(13, 3, mDigitColor) 'TopRight
            Call SetColor(4, 10, mDigitColor) 'Center
            Call SetColor(11, 12, mDigitColor) 'BottomRight
            Call SetColor(3, 18, mDigitColor) 'Bottom
    End Select
    
    'Refresh the Picturebox
    PicSrc.Refresh
    
End Sub

Private Sub RenderDisplay()
Dim StrValue As String
Dim ValueLen As Integer
Dim c As Integer
Dim Count As Integer
Dim bExtra As Integer

    'Digit Value
    StrValue = Value
    'Digit value length
    ValueLen = Len(StrValue)
    bExtra = 0
    
    If (BorderStyle = bFixedSingle) Then
        bExtra = 60
    End If
    
    If (AutoSize) Then
        UserControl.Width = (ValueLen * DigitWidth) * Screen.TwipsPerPixelX + bExtra
    End If
    
    With UserControl
        'Fix width if it goes below the Digit's width.
        If (UserControl.Width < (DigitWidth * Screen.TwipsPerPixelX)) Then
            UserControl.Width = (DigitWidth * Screen.TwipsPerPixelX)
        End If
        'Fix height to the digit's height
        .Height = (DigitHeight * Screen.TwipsPerPixelY) + bExtra
        
        For Count = 1 To ValueLen
            c = Val(Mid$(StrValue, Count, 1))
            'Draw the single digits
            Call CreateDigit(c)
            'Copy the new digit to the control.
            BitBlt UserControl.hdc, (Count * DigitWidth) - DigitWidth, 0, DigitWidth, _
            DigitHeight, PicSrc.hdc, 0, 0, vbSrcCopy
        Next Count
        'Refresh the Control.
        .Refresh
    End With
    
    StrValue = vbNullString
    ValueLen = 0
    Count = 0
End Sub

Private Sub UserControl_InitProperties()
    Value = 0
    BorderStyle = bFixedSingle
    AutoSize = True
    DigitColor = &HFFFF00
    DigitDimmedColor = &H808080
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H0&)
    DigitColor = PropBag.ReadProperty("DigitColor", &HFFFF00)
    DigitDimmedColor = PropBag.ReadProperty("DigitDimmedColor", &H808080)
    AutoSize = PropBag.ReadProperty("AutoSize", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", bFixedSingle)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", Value, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H0&)
    Call PropBag.WriteProperty("DigitColor", DigitColor, &HFFFF00)
    Call PropBag.WriteProperty("DigitDimmedColor", DigitDimmedColor, &H808080)
    Call PropBag.WriteProperty("AutoSize", AutoSize, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, bFixedSingle)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    Call RenderDisplay
End Sub

Private Sub UserControl_Show()
    Call RenderDisplay
End Sub

Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let Value(ByVal vNewValue As Long)
    mValue = vNewValue
    Call RenderDisplay
    PropertyChanged "Value"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call RenderDisplay
    PropertyChanged "BackColor"
End Property

Public Property Get DigitColor() As OLE_COLOR
    DigitColor = mDigitColor
End Property

Public Property Let DigitColor(ByVal oColor As OLE_COLOR)
    mDigitColor = oColor
    Call RenderDisplay
    PropertyChanged "DigitColor"
End Property

Public Property Get DigitDimmedColor() As OLE_COLOR
    DigitDimmedColor = mDigitDimmedColor
End Property

Public Property Let DigitDimmedColor(ByVal vNewDimColor As OLE_COLOR)
    mDigitDimmedColor = vNewDimColor
    Call RenderDisplay
    PropertyChanged "DigitDimmedColor"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = mAutoSize
End Property

Public Property Let AutoSize(ByVal vNewAutoSize As Boolean)
    mAutoSize = vNewAutoSize
    Call RenderDisplay
    PropertyChanged "AutoSize"
End Property

Public Property Get BorderStyle() As DigitBStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As DigitBStyle)
    UserControl.BorderStyle() = New_BorderStyle
    Call RenderDisplay
    PropertyChanged "BorderStyle"
End Property

