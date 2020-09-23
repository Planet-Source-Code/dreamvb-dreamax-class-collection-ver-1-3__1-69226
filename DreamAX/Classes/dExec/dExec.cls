VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Enum Operation
    oNone = 0
    oEdit = 1
    oExplore = 2
    oFind = 3
    oOpen = 4
    oPrint = 5
End Enum

Enum RetCode
    ERROR_FILE_NOT_FOUND = 2&
    ERROR_PATH_NOT_FOUND = 3&
    ERROR_BAD_FORMAT = 11&
    SE_ERR_ACCESSDENIED = 5
    SE_ERR_ASSOCINCOMPLETE = 27
    SE_ERR_DDEBUSY = 30
    SE_ERR_DDEFAIL = 29
    SE_ERR_DLLNOTFOUND = 32
    SE_ERR_DDETIMEOUT = 28
    SE_ERR_FNF = 2
    SE_ERR_NOASSOC = 31
    SE_ERR_OOM = 8
    SE_ERR_PNF = 3
    SE_ERR_SHARE = 26
End Enum

Enum WinShow
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum

Private m_WinShow As WinShow
Private mHwnd As Long
Private m_RetCode As RetCode

Public Property Get ShowCmd() As WinShow
    ShowCmd = m_WinShow
End Property

Public Property Let ShowCmd(ByVal Value As WinShow)
    m_WinShow = Value
End Property

Public Sub Exec(lpOperation As Operation, ByVal Filename As String, Optional ByVal Parameters As String = "", Optional ByVal Directory As String = "")
Dim Op As Variant
    Op = Array(vbNullString, "edit", "explore", "find", "open", "print")
    'Execute the given file.
    m_RetCode = ShellExecute(mHwnd, Op(lpOperation), Filename, Parameters, Directory, ShowCmd)
    'Clear array
    Erase Op
End Sub

Public Property Get ReturnCode() As RetCode
    ReturnCode = m_RetCode
End Property

Public Property Let ReturnCode(ByVal Value As RetCode)
    m_RetCode = Value
End Property

Public Property Get Hwnd() As Long
    Hwnd = mHwnd
End Property

Public Property Let Hwnd(ByVal Value As Long)
    mHwnd = Value
End Property
