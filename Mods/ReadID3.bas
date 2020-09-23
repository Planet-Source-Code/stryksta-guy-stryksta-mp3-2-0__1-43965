Attribute VB_Name = "ReadID3"
Option Explicit

Private strData As String * 127
Private strPath As String

Private strArtist As String
Private strTitle As String
Private strAlbum As String
Private strYear As String
Private strComment As String
Private hyph As String

Private TagCreated As Boolean

Public Sub LoadMp3File(valPath As String)
If UCase$(right$(valPath, 3)) = "MP3" Then
    strPath = valPath
            
    Open strPath For Binary As #1
    Get #1, FileLen(valPath) - 127, strData
    Close #1
    
    TagCreated = False
    
    If TagExists = True Then
        strArtist = Mid(strData, 34, 30)
        strTitle = Mid(strData, 4, 30)
        strAlbum = Mid(strData, 64, 30)
        strComment = Mid(strData, 98, 30)
        strYear = Mid(strData, 94, 4)
        hyph = " - "
    Else
        strArtist = "Untitled"
        strTitle = "Untitled"
        strAlbum = "Untitled"
        strYear = ""
        strComment = "Untitled"
    End If
    End If
    End Sub
                
Property Get Artist() As String
    
    Artist = RTrim(strArtist)
        
End Property

Property Get Title() As String

    Title = RTrim(strTitle)
        
End Property
Property Get FullSong() As String

    FullSong = RTrim(strArtist) & hyph & RTrim(strTitle)
        
End Property
Property Get Album() As String

    Album = RTrim(strAlbum)
            
End Property

Property Get Year() As String

    Year = RTrim(strYear)
    
End Property

Property Get Comment() As String
        
    Comment = RTrim(strComment)
          
End Property

Public Sub CloseMp3File()
    On Error GoTo errfOUNDx:

    Dim ToBeWritten As String
    
    SetAttr strPath, vbNormal
    
    Open strPath For Binary As #1
    
    FileLen (strPath)
    ToBeWritten = "TAG"
    Put #1, FileLen(strPath) - 127, ToBeWritten
   '
    ToBeWritten = strTitle & String(30 - Len(strTitle), " ")
    Put #1, FileLen(strPath) - 124, ToBeWritten
    
    ToBeWritten = strArtist & String(30 - Len(strArtist), " ")
    Put #1, FileLen(strPath) - 94, ToBeWritten
    
    ToBeWritten = strAlbum & String(30 - Len(strAlbum), " ")
    Put #1, FileLen(strPath) - 64, ToBeWritten
    
    ToBeWritten = strYear & String(4 - Len(strYear), " ")
    Put #1, FileLen(strPath) - 34, ToBeWritten
    
    ToBeWritten = strComment & String(30 - Len(strComment), " ")
    Put #1, FileLen(strPath) - 30, ToBeWritten
    
    Close #1
    
    TagCreated = True
errfOUNDx:
End Sub

Public Function TagExists() As Boolean
    
    If InStr(strData, "TAG") >= 1 Or TagCreated = True Then
        If right(strData, Len(strData) - 3) <> String(Len(strData) - 3, " ") Then
            TagExists = True
            Exit Function
        End If
    End If
    
    TagExists = False
    
End Function

Property Let Artist(valArtist As String)
     
    strArtist = valArtist
    
End Property

Property Let Title(valTitle As String)
     
    strTitle = valTitle
        
End Property

Property Let Album(valAlbum As String)
     
    strAlbum = valAlbum
    
End Property

Property Let Year(valYear As String)
     
    strYear = valYear
    
End Property

Property Let Comment(valComment As String)
     
    strComment = valComment
    
End Property
