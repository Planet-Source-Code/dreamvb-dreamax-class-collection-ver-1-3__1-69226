VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dRecentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_RecentFiles() As String
Private m_Count As Integer
Private m_Max As Integer
Private m_AppName As String

Event Error()

Public Sub AddItem(Item As String)
On Error GoTo TEmptyErr:
    
    'Check if items array has been resized.
    If (Not IsCreated) Then
        Err.Raise 9, "AddItem", "Not Created"
    End If
    'Current Item Counter
    m_Count = m_Count + 1
    'Check that we not adding more items than the max number.
    If (m_Count <= MaxItems) Then
        m_RecentFiles(m_Count) = Item
    Else
        'Add item to first item.
        m_RecentFiles(1) = Item
    End If
    
    Exit Sub
TEmptyErr:
    RaiseEvent Error
End Sub

Public Sub Clear()
On Error Resume Next
    'Clear all Items
    m_Count = 0
    Erase m_RecentFiles
    MaxItems = m_Max
    
    Call DeleteOldItems
End Sub

Public Property Get Item(Index As Integer) As String
    'Return Item Value.
    Item = m_RecentFiles(Index)
End Property

Public Property Let Item(Index As Integer, ByVal NewItem As String)
    'Set a Items value.
    m_RecentFiles(Index) = NewItem
End Property

Public Property Get MaxItems() As Integer
    'Return MaxItems.
    MaxItems = m_Max
End Property

Public Property Let MaxItems(ByVal NewMax As Integer)
    'Sets the max number of items.
    m_Max = NewMax
    
    'If (IsCreated) Then
        'Erase old items.
     '   m_Count = 0
      '  Erase m_RecentFiles
   ' End If
    
    'Resize the array to fit the items.
    ReDim Preserve m_RecentFiles(1 To NewMax)
End Property

Public Property Get AppName() As String
    'Return AppName
    AppName = m_AppName
End Property

Public Property Let AppName(ByVal NewName As String)
    'AppName to save items.
    m_AppName = NewName
End Property

Private Sub DeleteOldItems()
On Error Resume Next
    'Delete the old settings
    DeleteSetting m_AppName, "File List"
End Sub


Public Sub SaveItems()
On Error GoTo tErr:
Dim Counter As Integer
    Call DeleteOldItems
    'Save Items to Registry.
    For Counter = 1 To MaxItems
       SaveSetting AppName, "File List", Counter, Item(Counter)
    Next Counter
    
    Exit Sub
    'Error Trap
tErr:
    RaiseEvent Error
End Sub

Public Sub LoadItems()
On Error GoTo tErr:
Dim Counter As Integer
    'Check if items array has been resized.
    If Not IsCreated Then
        Err.Raise 9, "AddItem", "Not Created"
    Else
        'Load in the Items
        For Counter = 1 To MaxItems
            Item(Counter) = GetSetting(AppName, "File List", Counter)
        Next Counter
    End If
    
    'm_Count = Counter
    
    Exit Sub
    'Error Trap
tErr:
    RaiseEvent Error
End Sub

Private Function IsCreated() As Boolean
On Error GoTo tErr:
    IsCreated = UBound(m_RecentFiles) <> 0
    Exit Function
    'Error Trap
tErr:
    IsCreated = False
End Function

Private Sub Class_Terminate()
    'Clear up used vars.
    Erase m_RecentFiles
    m_Count = 0
    m_Max = 0
    m_AppName = vbNullString
End Sub
