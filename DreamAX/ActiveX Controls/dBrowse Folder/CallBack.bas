Attribute VB_Name = "BrowseFolder"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal HwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
        
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Enum BROWSEFLAGS
    bBROWSEFORCOMPUTER = &H1000
    bBROWSEFORPRINTER = &H2000
    bBROWSEINCLUDEFILES = &H4000
    bBROWSEINCLUDEURLS = &H80
    bDONTGOBELOWDOMAIN = &H2
    bEDITBOX = &H10
    bNEWDIALOGSTYLE = &H40
    bRETURNFSANCESTORS = &H8
    bRETURNONLYFSDIRS = &H1
    bSHAREABLE = &H8000
    bSTATUSTEXT = &H4
    bUSENEWUI = &H40
    bVALIDATE = &H20
End Enum

Public Enum TRootFolder
    NoSpecialFolder = -2
    RecycleBin = &HA
    ControlPanel = &H3
    Desktop = &H0
    DesktopDirectory = &H10
    MyComputer = &H11
    Fonts = &H14
    NetHood = &H13
    Network = &H12
    Personal = &H5
    Printers = &H4
    Programs = &H2
    Recent = &H8
    SendTo = &H9
    StartMenu = &HB
    Startup = &H7
    Templates = &H15
    StartUpNonLocalized = &H1D
    CommonStartUpNonLocalized = &H1E
    CommonDocuments = &H2E
    CommonFavorites = &H1F
    CommonPrograms = &H17
    CommonStartUp = &H18
    CommonTemplates = &H2D
    Cookies = &H21
    Favorites = &H6
    History = &H22
    Internet = &H1
    MyMusic = &HD
    Printhood = &H1B
    Connections = &H31
End Enum

Private Const BFFM_SETSELECTIONA As Long = &H466
Private Const BFFM_INITIALIZED As Long = 1

Public m_Title As String
Public m_Flags As BROWSEFLAGS
Public m_Hwnd As Long
Public m_RootF As TRootFolder
Public m_Directory As String
Public m_BackSlash As Boolean
Public m_InitDirectory As String

Private Function GetFuncAddr(Func As Long) As Long
    'Used to Return a pointer to a function
    GetFuncAddr = Func
End Function

Public Function GetBrowseForFolder()
Dim bInf As BROWSEINFO
Dim PathID As Long
Dim RetPath As String
Dim PathRootID As Long

    If (m_RootF = NoSpecialFolder) Then
        'If no Special folder is selected use the user defind one
        PathRootID = 0
        m_RootF = 0
    Else
        SHGetSpecialFolderLocation m_Hwnd, m_RootF, PathRootID
    End If
    
    'Fill BrowseFolder Type
    With bInf
        .hOwner = m_Hwnd
        .lpszTitle = m_Title
        .ulFlags = m_Flags
        .pidlRoot = PathRootID
        .lpfn = GetFuncAddr(AddressOf BrowseCallBackFunc)
        .lParam = SHSimpleIDListFromPath(StrConv(m_InitDirectory, vbUnicode))
    End With
    
    'Get Path ID
    PathID = SHBrowseForFolder(bInf)
    
    If (PathID) Then
        'Create Buffer
        RetPath = Space$(512)
        'Get Folder Path
        If SHGetPathFromIDList(ByVal PathID, ByVal RetPath) Then
            'Strip nullchars
            m_Directory = Left$(RetPath, InStr(RetPath, Chr$(0)) - 1)
            'Add a backslash to the path if required
            If (m_BackSlash) Then m_Directory = FixPath(m_Directory)
            CoTaskMemFree PathID
        End If
    End If
    
    'Clean up
    ZeroMemory bInf, Len(bInf)
    RetPath = vbNullString
    If (PathRootID) Then CoTaskMemFree PathRootID
End Function

Private Function BrowseCallBackFunc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If (uMsg = BFFM_INITIALIZED) Then
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, 0, ByVal lpData)
    End If
End Function

Private Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

