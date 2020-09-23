VERSION 5.00
Begin VB.UserControl dID3v1 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dID3v1.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "dID3v1.ctx":0132
End
Attribute VB_Name = "dID3v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ID3v1
    ID As String * 3
    vTitle As String * 30
    vArtist As String * 30
    vAlbum As String * 30
    vYear As String * 4
    vComment As String * 28
    vFiller As Byte
    vTrack As Byte
    vGenre As Byte
End Type

Private m_Tag As String
Private m_Title As String
Private m_Artist As String
Private m_Album As String
Private m_Year As String
Private m_Comment As String
Private m_Genre As Byte
Private m_Track As Byte
Private m_HasTag As Boolean
'
Private m_ID3V1 As ID3v1
Private m_GenreList() As String
Private m_IsOpen As Boolean
Private m_FilePtr As Long

Public Sub CloseMP3()
    If (m_IsOpen <> False) Then
        Close #m_FilePtr
    End If
    m_IsOpen = False
End Sub

Private Function FindFile(lzFile As String) As Boolean
On Error Resume Next
    FindFile = LenB(Dir(lzFile)) <> 0
End Function

Private Sub InitGenres()
Dim s_Genre As String

    'Genres List
    s_Genre = "Blues,Classic Rock,Country,Dance,Disco,Funk,Grunge,Hip-Hop,Jazz," _
    & "Metal,New Age,Oldies,Other,Pop,R&B,Rap,Reggae,Rock,Techno,Industrial," _
    & "Alternative,Ska,Death Metal,Pranks,Soundtrack,Euro-Techno,Ambient," _
    & "Trip-Hop,Vocal,Jazz+Funk,Fusion,Trance,Classical,Instrumental,Acid," _
    & "House,Game,Sound Clip,Gospel,Noise,AlternRock,Bass,Soul,Punk,Space," _
    & "Meditative,Instrumental Pop,Instrumental Rock,Ethnic,Gothic,Darkwave," _
    & "Techno-Industrial,Electronic,Pop-Folk,Eurodance,Dream,Southern Rock," _
    & "Comedy,Cult,Gangsta,Top 40 ,Christian Rap,Pop/Funk,Jungle,Native American," _
    & "Cabaret,New Wave,Psychadelic,Rave,Showtunes,Trailer,Lo-Fi,Tribal,Acid Punk," _
    & "Acid Jazz,Polka,Retro,Musical,Rock & Roll,Hard Rock,Folk,Folk-Rock,National Folk," _
    & "Swing,Fast Fusion,Bebob,Latin,Revival,Celtic,Bluegrass,Avantgarde,Gothic Rock," _
    & "Progressive Rock,Psychedelic Rock,Symphonic Rock,Slow Rock,Big Band,Chorus," _
    & "Easy Listening,Acoustic,Humour,Speech,Chanson,Opera,Chamber Music,Sonata," _
    & "Symphony,Booty Bass,Primus,Porn Groove,Satire,Slow Jam,Club,Tango,Samba," _
    & "Folklore,Ballad,Power Ballad,Rhythmic Soul,Freestyle,Duet,Punk Rock," _
    & "Drum Solo,A Capella,Euro-House,Dance Hall,Goa,Drum & Bass,Club-House," _
    & "Hardcore,Terror,Indie,BritPop,Negerpunk,Polsk Punk,Beat,Christian Gangsta Rap," _
    & "Heavy Metal,Black Metal,Crossover,Contemporary Christian,Christian Rock,Merengue," _
    & "Salsa,Thrash Metal,Anime,JPop,Synthpop,Unknown"
    
    m_GenreList = Split(s_Genre, ",")
    s_Genre = vbNullString
    
End Sub

Public Sub OpenMP3(Filename As String)
    'Close the File
    Call CloseMP3
    
    m_FilePtr = FreeFile 'File Pointer
    
    'Check that the file is found.
    If Not FindFile(Filename) Then
        Err.Raise 53, "dID3v1.OpenMP3", "File Not Found:" & vbCrLf & Filename
        Exit Sub
    End If
    
    'Open the file
    Open Filename For Binary As #m_FilePtr
        Seek #m_FilePtr, LOF(m_FilePtr) - 127
        Get #m_FilePtr, , m_ID3V1
        'Check for vaild TAG
        If (m_ID3V1.ID <> "TAG") Then
            'Close the MP3 File
            Call CloseMP3
            Exit Sub
        Else
            With m_ID3V1
                m_Title = StripNulls(.vTitle)
                m_Artist = StripNulls(m_ID3V1.vArtist)
                m_Album = StripNulls(m_ID3V1.vAlbum)
                m_Year = StripNulls(m_ID3V1.vYear)
                m_Comment = StripNulls(m_ID3V1.vComment)
                m_Genre = StripNulls(m_ID3V1.vGenre)
                m_Track = StripNulls(m_ID3V1.vTrack)
            End With
        End If
        
        m_IsOpen = True
        m_HasTag = True
End Sub

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

Public Sub UpdateMP3()
Dim Tag As String * 3
On Error GoTo SaveErr:
    
    'Check that the MP3 File is Open
    If (m_IsOpen <> True) Then
        Err.Raise 9, "dID3v1.UpdateMP3", "File Not Open"
        Exit Sub
    End If
    
    Seek #m_FilePtr, LOF(m_FilePtr) - 127
    'Get the ID Tag
    Get #m_FilePtr, , Tag
    'No Tag so we need to start at the end of the file
    If (Tag <> "TAG") Then
        m_ID3V1.ID = "TAG"
        Seek #m_FilePtr, LOF(m_FilePtr)
    End If
    'Set up the TAG Info
    With m_ID3V1
        .vTitle = m_Title
        .vArtist = m_Artist
        .vAlbum = m_Album
        .vYear = m_Year
        .vComment = m_Comment
        .vFiller = 0
        .vTrack = m_Track
        .vGenre = m_Genre
    End With
    'Write th etag information to the
    Put #m_FilePtr, , m_ID3V1
    '
    Exit Sub
SaveErr:
    Err.Raise 53, "dID3v1.Update", Err.Description
End Sub

Private Sub UserControl_Terminate()
    m_Tag = vbNullString
    m_Title = vbNullString
    m_Artist = vbNullString
    m_Album = vbNullString
    m_Year = vbNullString
    m_Comment = vbNullString
    m_Genre = 0
    m_Track = 0
    '
    m_ID3V1.vAlbum = vbNullString
    m_ID3V1.vArtist = vbNullString
    m_ID3V1.vComment = vbNullString
    m_ID3V1.vFiller = 0
    m_ID3V1.vGenre = 0
    m_ID3V1.vTitle = vbNullString
    m_ID3V1.vTrack = 0
    m_ID3V1.vYear = vbNullString
    Erase m_GenreList
End Sub

Private Sub UserControl_Initialize()
    Call InitGenres
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 480, 480
End Sub

'MP3 Tag Propertys
Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal NewTitle As String)
    m_Title = NewTitle
End Property

Public Property Get Artist() As String
    Artist = m_Artist
End Property

Public Property Let Artist(ByVal NewArtist As String)
    m_Artist = NewArtist
End Property

Public Property Get Album() As String
    Album = m_Album
End Property

Public Property Let Album(ByVal NewAlbum As String)
    m_Album = NewAlbum
End Property

Public Property Get mYear() As String
    mYear = m_Year
End Property

Public Property Let mYear(ByVal NewYear As String)
    m_Year = NewYear
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(ByVal NewComment As String)
    m_Comment = NewComment
End Property

Public Property Get Genre() As Byte
    Genre = m_Genre
End Property

Public Property Let Genre(ByVal NewGenre As Byte)
    m_Genre = NewGenre
End Property

Public Property Get Track() As Byte
    Track = m_Track
End Property

Public Property Let Track(ByVal NewTrack As Byte)
    m_Track = NewTrack
End Property

Public Property Get Genres(ByVal Index As Integer) As String
On Error GoTo TErr:
    Genres = m_GenreList(Index)
TErr:
    If Err Then Err.Raise 9 + vbObject, "cID3v1.cls", Err.Description
End Property

Public Property Get GenresCount() As Integer
On Error GoTo TErr:
    GenresCount = UBound(m_GenreList)
    Exit Property
TErr:
    If Err Then Err.Raise 9, "dID3v1.GenresCount", Err.Description
End Property

Public Property Get HasTag() As Boolean
    HasTag = m_HasTag
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property




