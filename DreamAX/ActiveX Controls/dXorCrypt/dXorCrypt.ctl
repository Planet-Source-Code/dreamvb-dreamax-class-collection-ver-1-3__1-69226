VERSION 5.00
Begin VB.UserControl dXorCrypt 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dXorCrypt.ctx":0000
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   30
   ToolboxBitmap   =   "dXorCrypt.ctx":00E0
End
Attribute VB_Name = "dXorCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Key As String

Private Sub UserControl_InitProperties()
    Key = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Key = PropBag.ReadProperty("Key", "")
End Sub

Private Sub UserControl_Terminate()
    m_Key = vbNullString
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Key", Key, "")
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

Public Function XorCrypt(ByVal SrcData As String) As String
Dim mBytes() As Byte
Dim KeyBytes() As Byte
Dim Count As Long

Dim KeyLen As Long
Dim KeyIdx As Integer

On Error Resume Next

    mBytes = StrConv(SrcData, vbFromUnicode)
    KeyBytes = StrConv(m_Key, vbFromUnicode)
    
    'Key Size
    KeyLen = UBound(KeyBytes)
    'Loop tho mBytes and Encrypt
    For Count = 0 To UBound(mBytes)
        'Keep the index pos of the Key
        If (KeyIdx > KeyLen) Then KeyIdx = 0
        'Xor each byte with the each byte of the key.
        mBytes(Count) = mBytes(Count) Xor KeyBytes(KeyIdx)
        'INC Key Counter
        KeyIdx = (KeyIdx + 1)
    Next Count
    
    'Return the encrypted byte array.
    XorCrypt = StrConv(mBytes, vbUnicode)
    
    'Clean up
    Erase mBytes
    Erase KeyBytes
    KeyIdx = 0
    KeyLen = 0
    Count = 0
End Function

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal NewKey As String)
    m_Key = NewKey
    PropertyChanged "Key"
End Property

