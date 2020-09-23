VERSION 5.00
Begin VB.Form FrnEx6 
   Caption         =   "dValueList - Example"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   750
      Left            =   4530
      TabIndex        =   11
      Top             =   255
      Width           =   705
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   750
      Left            =   5295
      TabIndex        =   10
      Top             =   255
      Width           =   705
   End
   Begin VB.TextBox txtVal 
      Height          =   315
      Left            =   2355
      TabIndex        =   9
      Top             =   690
      Width           =   1320
   End
   Begin VB.TextBox txtKey2 
      Height          =   315
      Left            =   2355
      TabIndex        =   7
      Top             =   255
      Width           =   1320
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   750
      Left            =   3765
      TabIndex        =   3
      Top             =   255
      Width           =   705
   End
   Begin VB.TextBox txtKey1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2415
      Width           =   1620
   End
   Begin VB.ListBox LstKeys 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   315
      Width           =   1605
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   4995
      TabIndex        =   0
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   195
      Left            =   1875
      TabIndex        =   8
      Top             =   735
      Width           =   405
   End
   Begin VB.Label lblKey1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      Height          =   195
      Left            =   1875
      TabIndex        =   6
      Top             =   330
      Width           =   270
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2190
      Width           =   405
   End
   Begin VB.Label lblKey0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   75
      Width           =   270
   End
End
Attribute VB_Name = "FrnEx6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TValueList As dValueList

Private Sub Refresh1()
Dim X As Integer
On Error Resume Next
    
    With LstKeys
        .Clear
        For X = 0 To TValueList.Count
            .AddItem TValueList.Item(X).Key
        Next X
        .ListIndex = 0
    End With
    
End Sub

Private Sub cmdAdd_Click()
    'Add new key and value.
    TValueList.Add txtKey2.Text, txtVal.Text
    'Update Listbox
    Call Refresh1
End Sub

Private Sub cmdDelete_Click()
Dim InKey As String
    'Ask user for Key to delete.
    InKey = Trim(InputBox("Enter the values Key you want to delete.", "Delete"))
    If Len(InKey) = 0 Then Exit Sub
    'Check if the Key was found.
    If (Not TValueList.KeyEquals(InKey)) Then
        MsgBox "The key '" & InKey & "' was not found.", vbInformation, "Key Not Found"
        Exit Sub
    Else
        'Delete
        TValueList.Delete InKey
        Call Refresh1
    End If
    
End Sub

Private Sub cmdExit_Click()
    Unload FrnEx6
End Sub

Private Sub cmdUpdate_Click()
Dim oItem As New cItem
Dim sKey As String

    sKey = txtKey2.Text
    If TValueList.KeyEquals(sKey) = False Then
        MsgBox "The Key '" & sKey & "' Was not found.", vbInformation, "Key Not Found"
        Exit Sub
    End If
    'Store Item
    Set oItem = TValueList.Item(sKey)
    'Sets the new data
    oItem.Key = sKey
    oItem.Value = txtVal.Text
    'Update the item
    TValueList.Item(sKey) = oItem
    Call Refresh1
    
    Set oItem = Nothing
    sKey = vbNullString
    
End Sub

Private Sub Form_Load()
    'Create Value List
    Set TValueList = New dValueList
    'Add some Test Items
    With TValueList
        .Add "Drive", "C:\"
        .Add "Memory", 512
        .Add "A", 2
        .Add "Test", "C:\Windows\Temp"
        'Refresh Listbox.
        Call Refresh1
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set TValueList = Nothing
    Set FrnEx6 = Nothing
End Sub

Private Sub LstKeys_Click()
    txtKey1.Text = TValueList.Item(LstKeys.Text).Value
End Sub

Private Sub txtKey2_Change()
    cmdAdd.Enabled = Len(Trim(txtKey2.Text)) > 0
End Sub
