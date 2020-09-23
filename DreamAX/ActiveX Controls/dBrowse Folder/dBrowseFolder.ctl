VERSION 5.00
Begin VB.UserControl dBrowseFolder 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dBrowseFolder.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dBrowseFolder.ctx":01F4
End
Attribute VB_Name = "dBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowBrowseForFolder()
    Call GetBrowseForFolder
End Sub

Private Sub UserControl_InitProperties()
    m_Title = ""
    m_InitDirectory = ""
    m_Flags = 4096
    m_Hwnd = 0
    m_RootF = NoSpecialFolder
    m_BackSlash = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Title = PropBag.ReadProperty("DialogTitle", "")
    m_Flags = PropBag.ReadProperty("Flags", 4096)
    m_RootF = PropBag.ReadProperty("RootFolder", NoSpecialFolder)
    m_BackSlash = PropBag.ReadProperty("AppendBackslash", True)
    m_InitDirectory = PropBag.ReadProperty("StartDirectory", "")
End Sub

Private Sub UserControl_Terminate()
    m_Title = ""
    m_Flags = 0
    m_Hwnd = 0
    m_RootF = 0
    m_Directory = ""
    m_BackSlash = False
    m_InitDirectory = ""
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DialogTitle", m_Title, "")
    Call PropBag.WriteProperty("Flags", m_Flags, 4096)
    Call PropBag.WriteProperty("RootFolder", m_RootF, NoSpecialFolder)
    Call PropBag.WriteProperty("AppendBackslash", m_BackSlash, True)
    Call PropBag.WriteProperty("StartDirectory", m_InitDirectory, "")
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

'Properties
Public Property Get DialogTitle() As String
    DialogTitle = m_Title
End Property

Public Property Let DialogTitle(ByVal vNewTitle As String)
    m_Title = vNewTitle
    PropertyChanged "DialogTitle"
End Property

Public Property Get Flags() As BROWSEFLAGS
    Flags = m_Flags
End Property

Public Property Let Flags(ByVal vNewFlags As BROWSEFLAGS)
    m_Flags = vNewFlags
    PropertyChanged "Title"
End Property

Public Property Let HwndOwner(ByVal vNewHwnd As Long)
    m_Hwnd = vNewHwnd
    PropertyChanged "HwndOwner"
End Property

Public Property Get RootFolder() As TRootFolder
    RootFolder = m_RootF
End Property

Public Property Let RootFolder(ByVal vNewRoot As TRootFolder)
    m_RootF = vNewRoot
    PropertyChanged "RootFolder"
End Property

Public Property Get Directory() As String
    Directory = m_Directory
End Property

Public Property Get AppendBackslash() As Boolean
    AppendBackslash = m_BackSlash
End Property

Public Property Let AppendBackslash(ByVal vNewBackSlash As Boolean)
    m_BackSlash = vNewBackSlash
    PropertyChanged "AppendBackslash"
End Property

Public Property Get StartDirectory() As String
    StartDirectory = m_InitDirectory
End Property

Public Property Let StartDirectory(ByVal vNewDir As String)
    m_InitDirectory = vNewDir
    PropertyChanged "StartDirectory"
End Property
