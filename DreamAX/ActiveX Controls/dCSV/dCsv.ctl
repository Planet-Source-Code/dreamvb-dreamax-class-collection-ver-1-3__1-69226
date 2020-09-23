VERSION 5.00
Begin VB.UserControl dCsv 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dCsv.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dCsv.ctx":00D6
End
Attribute VB_Name = "dCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CsvFilename As String

Private m_CsvQ As New Collection

Public Sub OpenCSV()
Dim fp As Long
Dim sLine As String
    
    Set m_CsvQ = Nothing
    
    fp = FreeFile
    '
    Open m_CsvFilename For Input As #fp
        Do Until EOF(fp)
            Line Input #fp, sLine
            If Trim(sLine) > 0 Then

            End If
        Loop
    Close #fp

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Filename = PropBag.ReadProperty("Filename", "")
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

Public Property Get Filename() As String
    Filename = m_CsvFilename
End Property

Public Property Let Filename(ByVal NewFilename As String)
    m_CsvFilename = NewFilename
    PropertyChanged "Filename"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Filename", Filename, "")
End Sub
