VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dDateLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_MaxDate As String
Private m_Active As Boolean
Public Event DateUp()

Private Sub DoDateCheck()
Dim TmpDate As Date
    If IsDate(m_MaxDate) Then
        'Store the date.
        TmpDate = m_MaxDate
        'Check the userdate with the systems date.
        If (Date > TmpDate) Then
            If (m_Active) Then
                RaiseEvent DateUp
            End If
        End If
    End If
End Sub

Public Property Let MaxDate(ByVal vNewDate As String)
    m_MaxDate = Trim(vNewDate)
    'Only check the date when running in UserMode
    Call DoDateCheck
End Property

Public Property Get MaxDate() As String
    MaxDate = m_MaxDate
End Property

Public Property Get Active() As Boolean
    Active = m_Active
End Property

Public Property Let Active(ByVal mEnabled As Boolean)
    m_Active = mEnabled
End Property

Private Sub Class_Initialize()
    m_Active = True
End Sub
