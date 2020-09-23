VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dStringContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Items() As String

Public Sub Add(ByVal lpStr As String)
Dim cSize As Long
    'Get array size.
    cSize = SafeUBound
    'Resize the array to hold the item.
    ReDim Preserve m_Items(cSize + 1)
    'Set the array item.
    m_Items(count) = lpStr
    'Clear up
    cSize = 0
End Sub

Public Sub Delete(ByVal Index As Long)
Dim count As Long
Dim cTop As Long

    'Checks that the index is in the correct array bounds.
    If Not InRange(Index) Then
        Err.Raise 9, "Item"
        Exit Sub
    End If
    'Get array upper bound.
    cTop = (SafeUBound - 1)
    'Loop tho the array and shift all the items up
    For count = Index To cTop
        m_Items(count) = m_Items(count + 1)
    Next count
    'if top less then zero clear the array.
    If (cTop < 0) Then
        Call Clear
    Else
        'Resize the array remoevng the last index.
        ReDim Preserve m_Items(cTop)
    End If
    
    'Clear up
    count = 0
    cTop = 0
End Sub

Public Property Get count() As Long
    count = SafeUBound
End Property

Public Sub Clear()
    'Clear all items from the array.
    Erase m_Items
End Sub

Public Property Get Item(ByVal Index As Long) As String
    'Checks that the index is in the correct array bounds.
    If Not InRange(Index) Then
        Err.Raise 9, "Item"
        Exit Property
    End If
    'Return item.
    Item = m_Items(Index)
End Property

Public Property Let Item(ByVal Index As Long, ByVal NewItem As String)
    'Checks that the index is in the correct array bounds.
    If Not InRange(Index) Then
        Err.Raise 9, "Item"
        Exit Property
    End If
    'Set the new item.
    m_Items(Index) = NewItem
End Property

Public Function ToString(Optional ByVal Seperator As String = vbCrLf) As String
    'Returns the array as a whole string.
    ToString = Join(m_Items, Seperator)
End Function

Public Sub FromString(ByVal lpStr As String, Optional ByVal Seperator As String = vbCrLf)
    'Converts an input string to the array
    'First clear the array
    Call Clear
    'Store the string as the array.
    m_Items = Split(lpStr, Seperator, , vbTextCompare)
End Sub

Public Function IndexOf(ByVal lpStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
Dim pos As Long
Dim x As Long
    'Returns the index position of the item found in the array.
    '-1 is returned if the item is not found.
    pos = -1
    For x = 0 To SafeUBound
        'Compare lpStr with the items in the array
        If StrComp(lpStr, m_Items(x), Compare) = 0 Then
            'Return position found.
            pos = x
            'Skip out of the loop
            Exit For
        End If
    Next x
    'clear var.
    x = 0
    'Return the index
    IndexOf = pos
    
End Function

Public Function Equals(ByVal lpStr As String) As Boolean
    'Returns true of false if an item in the array is equal to lpStr
    Equals = IndexOf(lpStr, vbBinaryCompare) <> -1
End Function

Public Sub Swap(ByVal FromIndex As Long, ByVal ToIndex As Long)
Dim Tmp1 As String

    'Checks that the index is in the correct array bounds.
    If Not InRange(FromIndex) Or Not InRange(ToIndex) Then
        Err.Raise 9, "Item"
        Exit Sub
    End If
    'Not much point in swaping the same item.
    If (FromIndex = ToIndex) Then Exit Sub
    'Swap the items
    Tmp1 = m_Items(ToIndex)
    m_Items(ToIndex) = m_Items(FromIndex)
    m_Items(FromIndex) = Tmp1
    'Clear var
    Tmp1 = vbNullString
End Sub

Private Sub Class_Terminate()
    'Clear things up
    Call Clear
End Sub

'Private tools used in this class
'
'
Private Function SafeUBound() As Long
On Error GoTo ErrFlag
    'Return safe size of an error with out the error.
    SafeUBound = UBound(m_Items)
    Exit Function
ErrFlag:
    'Empty array
    SafeUBound = -1
End Function

Private Function InRange(Index As Long) As Boolean
    'Returns true if we are in a vaild range
    If (Index < 0) Or (Index > UBound(m_Items)) Then
        'We are out of range.
        InRange = False
    Else
        'Rnage ok.
        InRange = True
    End If
End Function


