VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Key As String
Private m_Value As String

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal NewKey As String)
    m_Key = NewKey
End Property

Public Property Get Value() As String
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As String)
    m_Value = NewValue
End Property
