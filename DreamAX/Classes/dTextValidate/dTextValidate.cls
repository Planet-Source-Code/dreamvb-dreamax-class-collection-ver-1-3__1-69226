VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dTextValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum TFormatTypes
    fNumeric = 0
    fAplha = 1
    fAlphaNumeric = 2
    fCustom = 3
End Enum

Private WithEvents obj_TxtBox As TextBox
Attribute obj_TxtBox.VB_VarHelpID = -1
Private m_Format As String
Private m_CustFormat As String
Private m_FormatT As TFormatTypes
Private m_AllowDelete As Boolean
Private m_AllowSpace As Boolean

Private Sub Class_Terminate()
    'Clear up.
    m_Format = vbNullString
    m_CustFormat = vbNullString
    m_AllowDelete = False
    m_AllowSpace = False
End Sub

Private Sub obj_TxtBox_KeyPress(KeyAscii As Integer)
    
    'Check if delete key can be used.
    If (AllowDelete) Then
        If (KeyAscii = 8) Then
            Exit Sub
        End If
    End If
    'Check if space key can be used.
    If (AllowSpace) And (m_FormatT <> fCustom) Then
        If (KeyAscii = 32) Then
            Exit Sub
        End If
    End If
    'Vaild formats allowed.
    Select Case m_FormatT
        Case fNumeric
            m_Format = "*[0-9]*"
        Case fAplha
            m_Format = "*[a-zA-Z]*"
        Case fAlphaNumeric
            m_Format = "*[a-zA-Z0-9]*"
        Case fCustom
            m_Format = m_CustFormat
    End Select
    
    'Check if format is vaild.
    If Not Chr(KeyAscii) Like (m_Format) Then
        'set key to zero not vaild format.
        KeyAscii = 0
    End If
    
End Sub

'Class Propertys
Public Property Get Format() As TFormatTypes
    Format = m_FormatT
End Property

Public Property Let Format(ByVal NewFormat As TFormatTypes)
    m_FormatT = NewFormat
End Property

Public Property Get AllowDelete() As Boolean
    AllowDelete = m_AllowDelete
End Property

Public Property Let AllowDelete(ByVal NewDelete As Boolean)
    m_AllowDelete = NewDelete
End Property

Public Property Get AllowSpace() As Boolean
    AllowSpace = m_AllowSpace
End Property

Public Property Let AllowSpace(ByVal NewSpace As Boolean)
    m_AllowSpace = NewSpace
End Property

Public Property Get CustomFormat() As String
    CustomFormat = m_CustFormat
End Property

Public Property Let CustomFormat(ByVal NewFormat As String)
    m_CustFormat = NewFormat
End Property

Public Property Get TextBox() As TextBox
    TextBox = obj_TxtBox
End Property

Public Property Let TextBox(ByVal NewTxtBox As TextBox)
    Set obj_TxtBox = NewTxtBox
End Property
