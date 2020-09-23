VERSION 5.00
Begin VB.UserControl dChangeIconDialog 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dChangeIconDialog.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dChangeIconDialog.ctx":0103
End
Attribute VB_Name = "dChangeIconDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_IconIndex As Long
Private m_Filename As String

Private Declare Function SHChangeIconDialog Lib "shell32" Alias "#62" (ByVal hOwner As Long, _
ByVal szFilename As String, ByVal Reserved As Long, lpIconIndex As Long) As Long

Private Function StripNulls(ByVal LzStr As String) As String
Dim Buffer As String
Dim Count As Integer
    'Used to strip away any NULL chars in a string
    For Count = 1 To Len(LzStr)
        If (Asc(Mid(LzStr, Count, 1)) = 0) Then
            Exit For
        Else
            Buffer = Buffer & Mid(LzStr, Count, 1)
        End If
    Next Count
    
    StripNulls = Trim(Buffer)
    Buffer = vbNullString
End Function

Public Sub ShowDialog()
Dim ico_path As String
Dim sBuff As String
Dim iRet As Long

    'Displays the change icon dialogbox
    'Create a buffer string also needs NULLchar added to end of string
    sBuff = sBuff & String(260, Chr(0))
    'Convert the filename to Unicode
    ico_path = StrConv(m_Filename, vbUnicode)
    'Show the icon dialog
    iRet = SHChangeIconDialog(hWnd, ico_path, 260, m_IconIndex)
    
    'If the return value was zero we exit
    If (iRet = 0) Then
        ico_path = ""
        sBuff = ""
        Exit Sub
    Else
        sBuff = StrConv(ico_path, vbFromUnicode)
        m_Filename = StripNulls(sBuff)
    End If
    
    'Clear used variables
    sBuff = ""
    ico_path = ""
    iRet = 0
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

Public Property Get IconIndex() As Long
   IconIndex = m_IconIndex
End Property

Public Property Let IconIndex(ByVal NewIndex As Long)
    m_IconIndex = NewIndex
    PropertyChanged "IconIndex"
End Property

Public Property Get Filename() As String
   Filename = m_Filename
End Property

Public Property Let Filename(ByVal NewFilename As String)
    m_Filename = NewFilename
    PropertyChanged "Filename"
End Property

Private Sub UserControl_Terminate()
    m_Filename = ""
    m_IconIndex = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("IconIndex", m_IconIndex, 0)
    Call PropBag.WriteProperty("Filename", m_Filename, "")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_IconIndex = PropBag.ReadProperty("IconIndex", 0)
    m_Filename = PropBag.ReadProperty("Filename", "")
End Sub
