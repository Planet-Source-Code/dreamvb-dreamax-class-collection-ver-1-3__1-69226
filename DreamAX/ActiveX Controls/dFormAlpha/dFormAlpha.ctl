VERSION 5.00
Begin VB.UserControl dFormAlpha 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dFormAlpha.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dFormAlpha.ctx":00D8
End
Attribute VB_Name = "dFormAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)

Private m_WndHangle As Long
Private m_Alpha As Byte
Private m_Enabled As Boolean

Private Sub UpdateAlpha(Hwnd As Long)
Dim OrgWinProc As Long
    'Old window style
    OrgWinProc = GetWindowLong(Hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    'Set the new window style
    SetWindowLong Hwnd, GWL_EXSTYLE, OrgWinProc
    
    If (Enabled) Then
        'Update alpha
        SetLayeredWindowAttributes Hwnd, 0, Alpha, &H2
    Else
        'Sets the alpha value to the default 255
        SetLayeredWindowAttributes Hwnd, 0, 255, &H2
    End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Size 450, 405
End Sub

Public Property Get Alpha() As Byte
    'Get the alpha value
    Alpha = m_Alpha
End Property

Public Property Let Alpha(ByVal New_Alpha As Byte)
    ' set the new alpha value
    m_Alpha = New_Alpha
    
    If (Ambient.UserMode) Then
        'Only do the alpha if the form is not in design time.
        Call UpdateAlpha(Parent.Hwnd)
    End If
    
    PropertyChanged "Alpha"
End Property

Private Sub UserControl_InitProperties()
    Alpha = 128
    Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Alpha = PropBag.ReadProperty("Alpha", 128)
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alpha", Alpha, 128)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vEnabled As Boolean)
    m_Enabled = vEnabled
    Call UpdateAlpha(Parent.Hwnd)
    PropertyChanged "Enabled"
End Property
