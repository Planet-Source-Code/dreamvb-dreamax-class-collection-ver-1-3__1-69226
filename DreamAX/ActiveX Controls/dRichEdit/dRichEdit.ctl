VERSION 5.00
Begin VB.UserControl dRichEdit 
   Alignable       =   -1  'True
   BackColor       =   &H000080FF&
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   11.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1395
   ScaleWidth      =   1755
End
Attribute VB_Name = "dRichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal Hwnd As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Const WM_USER As Long = &H400
Private Const WS_CHILD As Long = &H40000000
Private Const WS_BORDER As Long = &H800000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_EX_STATICEDGE As Long = &H20000
'Edit Messages
Private Const WM_CUT As Long = &H300
Private Const WM_COPY As Long = &H301
Private Const WM_PASTE As Long = &H302
Private Const WM_CLEAR As Long = &H303
Private Const EM_UNDO As Long = &HC7
Private Const EM_REDO As Long = (WM_USER + 84)
Private Const EM_SETSEL As Long = &HB1
'Scrollbars
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
'RichEdit styles
Private Const ES_AUTOVSCROLL As Long = &H40&
Private Const ES_NOHIDESEL As Long = &H100&
Private Const ES_WANTRETURN As Long = &H1000&
Private Const ES_MULTILINE As Long = &H4&
Private Const EM_LIMITTEXT As Long = &HC5
Private Const WM_SETFONT As Long = &H30

Private Type MyWindowType
    wParentHwnd As Long
    wHwnd As Long
    wWidth As Long
    wHeight As Long
End Type

Private mRTFWinType As MyWindowType
'
Private mLength As Long
Private mText As String

Private mScrollBar As ScrollBarConstants

Public Sub Cut()
    If (mRTFWinType.wHwnd <> 0) Then
        SendMessage mRTFWinType.wHwnd, WM_CUT, 0, 0
    End If
End Sub

Public Sub Copy()
    If (mRTFWinType.wHwnd <> 0) Then
        SendMessage mRTFWinType.wHwnd, WM_COPY, 0, 0
    End If
End Sub

Public Sub Paste()
    If (mRTFWinType.wHwnd <> 0) Then
        SendMessage mRTFWinType.wHwnd, WM_PASTE, 0, 0
    End If
End Sub

Public Sub Delete()
    If (mRTFWinType.wHwnd <> 0) Then
        SendMessage mRTFWinType.wHwnd, WM_CLEAR, 0, 0
    End If
End Sub

Public Sub Undo()
    If (mRTFWinType.wHwnd <> 0) Then
        SendMessage mRTFWinType.wHwnd, EM_UNDO, 0, 0
    End If
End Sub

Public Sub Redo()
    If (mRTFWinType.wHwnd <> 0) Then
        SendMessage mRTFWinType.wHwnd, EM_REDO, 0, 0
    End If
End Sub

Public Sub SelectAll()
    If (mRTFWinType.wHwnd <> 0) Then
        SendMessage mRTFWinType.wHwnd, EM_SETSEL, 0, 0
    End If
End Sub

Private Sub DestroyRTFWindow()
    'Destroy the window
    DestroyWindow mRTFWinType.wHwnd
    ZeroMemory mRTFWinType, Len(mRTFWinType)
End Sub

Public Sub LoadFromFile(ByVal Filename As String)
Dim fp As Long
Dim Bytes() As Byte
    fp = FreeFile
    
    Open Filename For Binary As #fp
        If LOF(fp) > 0 Then
            ReDim Bytes(0 To LOF(fp) - 1)
        End If
        
        Get #fp, , Bytes
    Close #fp
    
    Text = StrConv(Bytes, vbUnicode)
    
    Erase Bytes
    
End Sub

Public Sub SaveToFile(ByVal Filename As String)
Dim fp As Long
    fp = FreeFile
    
    Open Filename For Output As #fp
        Print #fp, Text
    Close #fp
    
End Sub

Private Sub CreateRichEditWnd()
Dim wStyle As Long
Dim sBars As Long

    'scrolbars
    If (Scrollbars = vbHorizontal) Then
        sBars = WS_HSCROLL
    ElseIf (Scrollbars = vbVertical) Then
        sBars = WS_VSCROLL
    ElseIf (Scrollbars = vbBoth) Then
        sBars = (WS_HSCROLL Or WS_VSCROLL)
    Else
        sBars = 0
    End If

    wStyle = (WS_CHILD Or WS_BORDER Or WS_VISIBLE Or ES_MULTILINE Or sBars _
    Or ES_AUTOVSCROLL Or ES_NOHIDESEL Or ES_WANTRETURN Or WS_TABSTOP)
 
    If LoadLibrary("riched20.dll") = 0 Then
        Exit Sub
    Else
        With mRTFWinType
            .wHeight = FixHeight
            .wWidth = FixWidth
            .wParentHwnd = UserControl.Hwnd
            'Create the Window.
            .wHwnd = CreateWindowEx(&H200&, "RichEdit20A", "", wStyle, _
            0, 0, .wWidth, .wHeight, .wParentHwnd, 0, App.hInstance, WS_EX_STATICEDGE)
            'Set Text Max Length
            SendMessage .wHwnd, EM_LIMITTEXT, MaxLength, ByVal 0&

            'Set Default Text
            SetWindowText .wHwnd, Text
            
            Call CreateFontObj
        
        End With
    End If
    
End Sub

Function CreateFontObj() As Long
Dim fs As Long
Dim hFont As Long

    With UserControl
        If (mRTFWinType.wHwnd <> 0) Then
            'Get Fontsize.
            fs = .TextHeight("Xy") \ Screen.TwipsPerPixelY
            'Create Font Object.
            hFont = CreateFont(fs, 0, 0, 0, .Font.Weight, .FontItalic, .FontUnderline, _
            .FontStrikethru, .Font.Charset, 4, 0, 0, 0, .FontName)
            'Set the Controls font.
            SendMessage mRTFWinType.wHwnd, WM_SETFONT, hFont, 0
            'Destroy the font object
            DeleteObject CreateFontObj
        End If
        
    End With
End Function

Private Function GetTextFromWnd() As String
Dim sLen As Long
Dim sBuff As String
Dim Ret As Long

    If (mRTFWinType.wHwnd <> 0) Then
        'Get Windows Text Length
        sLen = GetWindowTextLength(mRTFWinType.wHwnd) + 1
        'Create buffer to hold Window Text
        sBuff = Space(sLen)
        'Get the Text from the Window
        Ret = GetWindowText(mRTFWinType.wHwnd, sBuff, sLen)
        
        If (Ret > 0) Then
            'Return the Text
            GetTextFromWnd = Left$(sBuff, Ret)
        End If
        
        sBuff = vbNullString
    End If
End Function

Private Sub UserControl_InitProperties()
    MaxLength = 0
    Text = Ambient.DisplayName
    Scrollbars = vbSBNone
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
    UserControl.FontSize = PropBag.ReadProperty("FontSize", Font.Size)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
    Text = PropBag.ReadProperty("Text", Ambient.DisplayName)
    Scrollbars = PropBag.ReadProperty("Scrollbars", vbSBNone)
End Sub

Private Sub UserControl_Show()
    Call CreateRichEditWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MaxLength", MaxLength, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "MS Sans Serif")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, Font.Size)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, False)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, False)
    Call PropBag.WriteProperty("Text", Text, Ambient.DisplayName)
    Call PropBag.WriteProperty("Scrollbars", Scrollbars, vbSBNone)
End Sub

Private Sub UserControl_Resize()
    
    If (mRTFWinType.wHwnd) <> 0 Then
        MoveWindow mRTFWinType.wHwnd, 0, 0, FixWidth, FixHeight, 1
    End If
End Sub

Private Property Get FixWidth() As Long
    FixWidth = (UserControl.ScaleWidth \ Screen.TwipsPerPixelX)
End Property

Private Property Get FixHeight() As Long
    FixHeight = (UserControl.ScaleHeight \ Screen.TwipsPerPixelY)
End Property

Private Sub UserControl_Terminate()
    Call DestroyRTFWindow
End Sub

Public Property Get MaxLength() As Long
    MaxLength = mLength
End Property

Public Property Let MaxLength(ByVal vNewLength As Long)
    mLength = vNewLength
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call CreateFontObj
    PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    'Update Controls Font
    Call CreateFontObj
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    'Update Controls Font
    Call CreateFontObj
    PropertyChanged "FontItalic"
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    'Update Controls Font
    Call CreateFontObj
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    'Update Controls Font
    Call CreateFontObj
    PropertyChanged "FontSize"
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    'Update Controls Font
    Call CreateFontObj
    PropertyChanged "FontStrikethru"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    'Update Controls Font
    Call CreateFontObj
    PropertyChanged "FontUnderline"
End Property

Public Property Get Text() As String
    Text = GetTextFromWnd
End Property

Public Property Let Text(ByVal vNewText As String)
    If (mRTFWinType.wHwnd <> 0) Then
        SetWindowText mRTFWinType.wHwnd, vNewText
    End If
    
    PropertyChanged "Text"
End Property

Public Property Get Hwnd() As Long
    Hwnd = mRTFWinType.wHwnd
End Property


Public Property Get Scrollbars() As ScrollBarConstants
    Scrollbars = mScrollBar
End Property

Public Property Let Scrollbars(ByVal vNewBars As ScrollBarConstants)
    mScrollBar = vNewBars
    PropertyChanged "Scrollbars"
End Property
