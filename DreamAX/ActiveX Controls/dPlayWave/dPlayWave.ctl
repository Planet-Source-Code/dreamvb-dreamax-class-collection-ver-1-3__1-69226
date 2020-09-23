VERSION 5.00
Begin VB.UserControl dPlayWave 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dPlayWave.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dPlayWave.ctx":0112
End
Attribute VB_Name = "dPlayWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Enum wPlaySrcT
    tFilename = 0
    tResource = 1
    tMemory = 2
End Enum

Private m_Filename As String
Private m_Loop As Boolean
Private m_Asynchronous As Boolean
Private m_PlaySrc As wPlaySrcT
Private m_DataPtr As Long

Private Function GetFlags() As Long
Dim lnLoop As Long
Dim mPlayType As Long

    'Wave Loop
    If (m_Loop) Then
        lnLoop = &H8
    Else
        lnLoop = 0
    End If
    
    'Play Type
    Select Case m_PlaySrc
        Case tFilename
            'Playing from Filename.
            mPlayType = &H20000
        Case tResource
            'Playing from Resource File.
            mPlayType = &H40004
        Case tMemory
            'Playing from memory.
            mPlayType = &H4
    End Select
    
    GetFlags = mPlayType Or lnLoop Or Abs(m_Asynchronous)
End Function

Public Sub wPlay()
Dim iRet As Long
    'Paly the wave.
    If (m_PlaySrc = tMemory) Then
        'Play from a data Ptr
        PlaySound wDataPtr, App.hInstance, GetFlags
    Else
        'Play from a filename or resource ID
        iRet = PlaySound(wFilename, App.hInstance, GetFlags)
    End If
End Sub

Public Sub wStop()
    'Stop the wave
    PlaySound ByVal 0&, App.hInstance, 0
End Sub

Private Sub UserControl_InitProperties()
    wFilename = ""
    wLoop = False
    wAsynchronous = True
    wPlaySource = tFilename
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("wFilename", wFilename, "")
    Call PropBag.WriteProperty("wLoop", wLoop, False)
    Call PropBag.WriteProperty("wAsynchronous", wAsynchronous, True)
    Call PropBag.WriteProperty("wPlaySource", wPlaySource, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    wFilename = PropBag.ReadProperty("wFilename", "")
    wLoop = PropBag.ReadProperty("wLoop", False)
    wAsynchronous = PropBag.ReadProperty("wAsynchronous", True)
    wPlaySource = PropBag.ReadProperty("wPlaySource", 0)
End Sub

Public Property Get wFilename() As String
    wFilename = m_Filename
End Property

Public Property Let wFilename(ByVal vNewFName As String)
    m_Filename = vNewFName
    PropertyChanged "wFilename"
End Property

Public Property Get wLoop() As Boolean
    wLoop = m_Loop
End Property

Public Property Let wLoop(ByVal NewwLoop As Boolean)
    m_Loop = NewwLoop
    PropertyChanged "wLoop"
End Property

Public Property Get wAsynchronous() As Boolean
    wAsynchronous = m_Asynchronous
End Property

Public Property Let wAsynchronous(ByVal vNewValue As Boolean)
    m_Asynchronous = vNewValue
    PropertyChanged "wAsynchronous"
End Property

Public Property Get wPlaySource() As wPlaySrcT
    wPlaySource = m_PlaySrc
End Property

Public Property Let wPlaySource(ByVal NewSrc As wPlaySrcT)
    m_PlaySrc = NewSrc
    PropertyChanged "wPlaySource"
End Property

Public Property Get wDataPtr() As Long
    wDataPtr = m_DataPtr
End Property

Public Property Let wDataPtr(ByVal vNewPtr As Long)
    m_DataPtr = vNewPtr
    PropertyChanged "wDataPtr"
End Property
