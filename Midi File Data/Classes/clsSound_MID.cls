VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSound_MID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Sound Module re-coded from my ATL C++ dll MME.dll''
'Copyright © Steven Paterson, 2002 - 2003'''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
'MID VERSION: Taken SetSpeed Out; Converted to only Mids
'not MultiFormat Since then it uses Millisecond time vals
'not normal midi positions, lengths.

Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Getshortpathname is for passing the file to be played to the mci api, as it's old
'and requires short path
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Variable to find out if being played
Private BlnPlaying As Boolean

Public Function PlayFile(FileName As String, Volume As Integer)
Dim ShortFileName As String * 100
'ShortFileName = Space(100)

'Mci doesn't like long filename's, it's old.
GetShortPathName FileName, ShortFileName, 100
ShortFileName = StripNullChars(ShortFileName)

'open as mid
mciSendString "open " & ShortFileName & " type sequencer alias mid", 0, 0, 0

'set it so when we enquire where it's upto it tell's us in milliseconds.
mciSendString "set mid time format tmsf", 0, 0, 0

'set volume,
'MCI Volume spans from 0 (Quietest) to 1000 (Loudest)
mciSendString "setaudio mid volume to" & Volume, 0, 0, 0

'play the file
mciSendString "play mid", 0, 0, 0

BlnPlaying = True

End Function

Public Function PlayFileFrom(FileName As String, Position As Long, Volume As Integer)
Dim ShortFileName As String * 100
'ShortFileName = Space(100)
'Mci doesn't like long filename's, it's old.
GetShortPathName FileName, ShortFileName, 100
ShortFileName = StripNullChars(ShortFileName)

'open as mid
mciSendString "open " & ShortFileName & " type sequencer alias mid", 0, 0, 0

'set it so when we enquire where it's upto it tell's us in milliseconds.
mciSendString "set mid time format tmsf", 0, 0, 0

'set volume,
'MCI Volume spans from 0 (Quietest) to 1000 (Loudest)
mciSendString "setaudio mid volume to" & Volume, 0, 0, 0

'Set the value that makes this Playfrom
mciSendString "play mid from " & Position, 0, 0, 0

'play the file
mciSendString "play mid", 0, 0, 0

BlnPlaying = True

End Function

Property Get Playing() As Boolean
Playing = BlnPlaying
End Property

Public Function SetPosition(Position As Long)
mciSendString "stop mid", 0, 0, 0
mciSendString "play mid from " & Position, 0, 0, 0
'mciSendString "play mid", 0, 0, 0
End Function

Public Function GetPosition() As String

GetPosition = Space(100) 'Since we can't use fixed length in title
mciSendString "status mid position", GetPosition, 100, 0

If GetPosition = GetLength Then
    BlnPlaying = False
    CloseFile
    GetPosition = 0
End If

End Function

Public Function SetVolume(Volume As Integer)
'Remember Volume spans from 0 to 1000.
mciSendString "setaudio mid volume to " & Volume, 0, 0, 0
End Function

Public Function GetLength() As String
GetLength = Space(100) 'Since we can't use fixed length in title
mciSendString "status mid length", GetLength, 100, 0
End Function

Public Function CloseFile()
mciSendString "close mid", 0, 0, 0
BlnPlaying = False
End Function

Public Function StopFile()
mciSendString "stop mid", 0, 0, 0
BlnPlaying = False
End Function

Public Function PauseFile()
mciSendString "pause mid", 0, 0, 0
BlnPlaying = False
End Function

Public Function UnPauseFile()
mciSendString "play mid", 0, 0, 0
BlnPlaying = True
End Function

Private Function StripNullChars(StrStrip As String) As String
Dim I As Integer
Dim Character As Byte

    For I = 1 To Len(StrStrip) 'Reads each char
     
        Character = Asc(Mid$(StrStrip, I, 1)) 'Gets the char
        
        If (Character < 0 Or Character > 31) Then StripNullChars = StripNullChars & Chr(Character)    'adds the char if its displayable/not a null
    
    Next I 'next char


End Function


