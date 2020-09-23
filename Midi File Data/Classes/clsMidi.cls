VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMidi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''
'Class Based on Stefaan Casier's Version'
'of this code, wrote by Steven Paterson''
'''''''''''''''''''''''''''''''''''''''''
'This code was wrote when i wanted to find data
'from a midi, so i searched code guru.com for
'a midi util, and it came up with
'"Advanced Midi Programming / Util" and i liked
'(it's the same as the one on PSC), it's data
'it could get on mids (instrument changed, sections
'text notes, singular notes, etc.)
'But the code really was very hard to read and all
'the messages were in dutch; while some still are
'as i couldn't translate their meaning, most are
'now in english, also i've made it available to
'change the max length of a report while you couldn't
'before.
'
'Stefaan Casier's Original Code is in ModOriginal.
'While the output it basically the same, my code
'i think is much easier to learn from.

'Author: Steven Paterson <Ste@Tolon.co.uk>
'© Steven Paterson 2003
'This Code is to only be used in non-commerical
'applications
'Class may be freely distrubuted and used.
'Please keep this notice above when re-distrubuting
'or being used.

Option Explicit

'Report Flags for ReadMidi
Private FilterNoteMessages As Boolean
Private FilterControlMessages As Boolean
Private FilterSysExMessages As Boolean

'Max Report length default is 32k.
Private MaxReportLen As Long

Public Function ReadMidiHeader(FileLocation As String) As String
Dim B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte
Dim NumberOfTracks As Integer, Track As Integer
Dim FormatType As Integer
Dim FileNum As Integer
Dim FileData As String
Dim Division As Long
Dim MT As String * 4

    FileNum = FreeFile

    Open FileLocation For Binary As FileNum

    Get #FileNum, 1, MT
        
        'All Midi's start with "MThd" MT Header
        If MT <> "MThd" Then
            FileData = FileData & "No Header/Unknown Header" & vbCrLf
        Else
            Get #FileNum, 5, B1
            Get #FileNum, 6, B2
            Get #FileNum, 7, B3
            Get #FileNum, 8, B4
        End If
        
        'B1,B2,B3,B4 in this instance must be 00 00 00 06'
        'because these specify the header length.
        If Not (B1 = 0 And B2 = 0 And B3 = 0 And B4 = 6) Then
            FileData = FileData & "Invalid Midi Header Length" & vbCrLf
            GoTo FunctionEnd
        End If
        
        'Format Type is whether it has multiple tracks
        'or a single track, and whether the multiple
        'is asynchronous or synchronous
        'this means asynchronous is multiple tracks
        'starting at different times
        'synchronous means multiple tracks all start
        'at the same time.
        
        Get #FileNum, 9, B1
        Get #FileNum, 10, B2
        FormatType = B1 * 256 + B2
        FileData = FileData & "Format type = " & CStr(FormatType)
        
        Select Case FormatType
        
         Case 0
             FileData = FileData & " - Single track" & vbCrLf
    
         Case 1
             FileData = FileData & " - Multiple tracks, synchronous" & vbCrLf
    
         Case 2
             FileData = FileData & " - Multiple tracks, asynchronous" & vbCrLf
         
         Case Else
             FileData = FileData & " - Unknown"
             GoTo FunctionEnd
        
        End Select
        
        'Get's the number of tracks
        Get #FileNum, 11, B1
        Get #FileNum, 12, B2
        
        NumberOfTracks = B1 * 256 + B2
        FileData = FileData & "Number of tracks = " & CStr(NumberOfTracks) & vbCrLf
        
        'Format Type 0 = Single Track, if it has more than 1 track it's Invalid.
        If FormatType = 0 And NumberOfTracks > 1 Then
            FileData = FileData & "Invalid Format Type, Multiple Tracks Found on a Single Track Format"
            GoTo FunctionEnd
        End If
        
        'PPQN is Note Resolution.
        Get #FileNum, 13, B1
        Get #FileNum, 14, B2
        
        Division = B1 * 256 + B2
        FileData = FileData & "Division = " & CStr(Division) & " PPQN" & vbCrLf
        
    Close FileNum

FunctionEnd:
    ReadMidiHeader = FileData
    
End Function

Public Function ReadMidi(FileLocation As String) As String
Dim B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte
Dim NumberOfTracks As Integer, Track As Integer
Dim Position As Long, Pos As Long
Dim FormatType As Integer
Dim EndOfTrack As Boolean
Dim Delta As String * 7
Dim MT As String * 4
Dim FileNum As Integer
Dim FileData As String
Dim Division As Long
Dim nBytes As Long
Dim Status As Byte
Dim Stat As String
Dim Reg As String
Dim Lng As Long
Dim DT As Long
Dim P As Long

    FileNum = FreeFile

    Open FileLocation For Binary As FileNum

    Get #FileNum, 1, MT
        
        'All Midi's start with "MThd" MT Header
        If MT <> "MThd" Then
            FileData = FileData & "No Header/Unknown Header" & vbCrLf
        Else
            Get #FileNum, 5, B1
            Get #FileNum, 6, B2
            Get #FileNum, 7, B3
            Get #FileNum, 8, B4
        End If
        
        'B1,B2,B3,B4 in this instance must be 00 00 00 06'
        'because these specify the header length.
        If Not (B1 = 0 And B2 = 0 And B3 = 0 And B4 = 6) Then
            FileData = FileData & "Invalid Midi Header Length" & vbCrLf
            GoTo FunctionEnd
        End If
        
        'Format Type is whether it has multiple tracks
        'or a single track, and whether the multiple
        'is asynchronous or synchronous
        'this means asynchronous is multiple tracks
        'starting at different times
        'synchronous means multiple tracks all start
        'at the same time.
        
        Get #FileNum, 9, B1
        Get #FileNum, 10, B2
        FormatType = B1 * 256 + B2
        FileData = FileData & "Format type = " & CStr(FormatType)
        
        Select Case FormatType
        
         Case 0
             FileData = FileData & " - Single track" & vbCrLf
    
         Case 1
             FileData = FileData & " - Multiple tracks, synchronous" & vbCrLf
    
         Case 2
             FileData = FileData & " - Multiple tracks, asynchronous" & vbCrLf
         
         Case Else
             FileData = FileData & " - Unknown"
             GoTo FunctionEnd
        
        End Select
        
        'Get's the number of tracks
        Get #FileNum, 11, B1
        Get #FileNum, 12, B2
        
        NumberOfTracks = B1 * 256 + B2
        FileData = FileData & "Number of tracks = " & CStr(NumberOfTracks) & vbCrLf
        
        'Format Type 0 = Single Track, if it has more than 1 track it's Invalid.
        If FormatType = 0 And NumberOfTracks > 1 Then
            FileData = FileData & "Invalid Format Type, Multiple Tracks Found on a Single Track Format"
            GoTo FunctionEnd
        End If
        
        'PPQN is Note Resolution.
        Get #FileNum, 13, B1
        Get #FileNum, 14, B2
        
        Division = B1 * 256 + B2
        FileData = FileData & "Division = " & CStr(Division) & " PPQN" & vbCrLf
        
        Position = 15
        '15 Where the first track should start
   
        For Track = 1 To NumberOfTracks
            
            EndOfTrack = False
   
            Get #FileNum, Position, MT
            
            If MT <> "MTrk" Then 'MT Track
                'This error occurs on wrong # of tracks, or wrong track length.
                FileData = FileData & "No Midi track found where expected!" & vbCrLf
            Else
          
              Get #FileNum, , B1
              Get #FileNum, , B2
              Get #FileNum, , B3
              Get #FileNum, , B4
          
              'Track Length Calculation
              nBytes = CLng(CLng(B1) * 256 ^ 3 + CLng(B2) * 256 ^ 2 + CLng(B3) * 256 + CLng(B4))
                
              FileData = FileData & vbCrLf & "Track " & CStr(Track) & "     Length = " & CStr(nBytes) & vbCrLf
        
              Position = Position + 8
              Pos = Position
        
              Status = 0
              
              'This part uses ReadMidiFF, this is just
              'to read what a byte means.
              
              While Position - Pos < nBytes
              
                  Get #FileNum, Pos, B1
                  
                  If B1 = &HFF Then
                  
                      Position = Position + 1
                      Status = B1
                      Reg = readMidiFF(FileNum, Position, EndOfTrack)
                      Stat = " " & HexByte(Status) & " "
                      
                  Else
                  
                      DT = ReadVariableLength(FileNum, Position)
                      Delta = CStr(DT)
                      
                      Get #FileNum, Position, B1
                      
                      If (B1 And &H80) = &H80 Then
                      
                         Status = B1
                         Stat = " " & HexByte(Status) & " "
                         Position = Position + 1
                         
                      Else
                      
                         Stat = "r" & HexByte(Status) & " "
                         
                      End If
                         
                      Select Case Status And &HF0
                      
                      Case &H80
                      
                          Get #FileNum, Position, B2
                          Position = Position + 1
                          Get #FileNum, Position, B3
                          Position = Position + 1
                          
                          If FilterNoteMessages = False Then
                              Reg = "Note off.... " & isNote(B2) & "-" & CStr(B3)
                          End If
                          
                      Case &H90
                         
                          Get #FileNum, Position, B2
                          Position = Position + 1
                          Get #FileNum, Position, B3
                          Position = Position + 1
                          
                          If FilterNoteMessages = False Then
                              Reg = "Note on..... " & isNote(B2) & "-" & CStr(B3)
                          End If
                          
                      Case &HB0
                      
                          Get #FileNum, Position, B2
                          Position = Position + 1
                          Get #FileNum, Position, B3
                          Position = Position + 1
                         
                          If FilterControlMessages = False Then
                              Reg = "Ctl FileNumange.. " & HexByte(B2) & " " & HexByte(B3)
                          End If
                          
                      Case &HC0
                      
                          Get #FileNum, Position, B2
                          Position = Position + 1
                          Reg = "Prg FileNumange.. " & HexByte(B2)
                         
                      Case &HD0
                      
                          Get #FileNum, Position, B2
                          Position = Position + 1
                          Reg = "FileNuman Press.. " & HexByte(B2)
                         
                      Case &HE0
                      
                          Get #FileNum, Position, B2
                          Position = Position + 1
                          Get #FileNum, Position, B3
                          Position = Position + 1
                          Reg = "PitFileNum bend.. " & HexByte(B2) & " " & HexByte(B3)
                      
                      Case &HF0
                      
                          Select Case Status
                         
                              Case &HFE
                              
                              Case &HFF
                              
                                  Reg = readMidiFF(FileNum, Position, EndOfTrack)
                                 
                              Case &HF0
                              
                                  P = Pos
                                  Lng = ReadVariableLength(FileNum, Position)
                                  If FilterSysExMessages = False Then Reg = "SysEx - len: " & CStr(Lng)
                                  Position = Position + Lng
                                 
                              Case &HF7
                              
                              Case Else
                              
                          End Select
                              
                  End Select
                  
              End If
              
              If Reg <> "" Then
                  FileData = FileData & Delta & Stat & Reg & vbCrLf
                  Reg = ""
              End If
              
              If Len(FileData) > MaxReportLen Then
                  FileData = FileData & vbCrLf & "Maximum Report Length!" & vbCrLf
                  GoTo FunctionEnd
              End If
              
              Wend
              
              DoEvents
              
            End If
    
        Next Track

    Close FileNum

FunctionEnd:
    ReadMidi = FileData
    
End Function

Private Function readMidiFF(ByVal FileNum As Long, Position As Long, EndOfTrack As Boolean)
Dim Bytes As Long
Dim I As Long
Dim B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte, B5 As Byte
Dim TotalData As String, Data As String * 13
'TotalData is all the data; Data is just a helper.

Get #FileNum, Position, B2
Position = Position + 1

If B2 = 0 Then
    
    Get #FileNum, Position, B3
    Position = Position + 1
    
    If B3 = 0 Then
        TotalData = TotalData & "Seqnr/Posfile"
    Else
    
        Get #FileNum, Position, B4
        Position = Position + 1
        
        Get #FileNum, Position, B5
        Position = Position + 1
        
        TotalData = TotalData & "Seq Nr       " & CStr(B5 * 256 + B4)
        
    End If
    
ElseIf B2 >= 1 And B2 <= 7 Then
    
    'Finds out what the data is by the preceding bit tag.
    Data = Choose(B2, "Text", "Copyright", "Seq/Tr. Name", "Instrument", "Lyric", "Marker", "Cue Point") & " - "
    TotalData = TotalData & Data
    
    Bytes = ReadVariableLength(FileNum, Position)
    
    For I = 1 To Bytes
        
        Get #FileNum, Position, B1
        Position = Position + 1
        TotalData = TotalData & Chr(B1)
        
    Next I
    
ElseIf B2 = &H20 Then

    TotalData = TotalData & "Midi FileNumannel   "
    
    Get #FileNum, Position, B3
    Position = Position + 1
    
    Get #FileNum, Position, B4
    Position = Position + 1
    
    If B3 <> 0 Then TotalData = TotalData & "???Len"
    TotalData = TotalData & HexByte(B4)
    
ElseIf B2 = &H21 Then

    TotalData = TotalData & "Midi Port    "
    
    Get #FileNum, Position, B3
    Position = Position + 1
    
    Get #FileNum, Position, B4
    Position = Position + 1
    
    If B3 <> 0 Then TotalData = TotalData & "???len"
    TotalData = TotalData & HexByte(B4)
    
ElseIf B2 = &H2F Then

    TotalData = TotalData & "End of track "
    Get #FileNum, Position, B3
    Position = Position + 1
    EndOfTrack = True

ElseIf B2 = &H51 Then

    TotalData = TotalData & "Tempo        "
    Get #FileNum, Position, B3
    Position = Position + 1
    Bytes = B3
      
    If Bytes <> 3 Then TotalData = TotalData & " ???Len "
      
    Get #FileNum, Position, B3
    Position = Position + 1
      
    Get #FileNum, Position, B4
    Position = Position + 1
      
    Get #FileNum, Position, B5
    Position = Position + 1

ElseIf B2 = &H54 Then
    
    TotalData = TotalData & "SMPTE Offs   "
    
    Get #FileNum, Position, B3
    Position = Position + 1
    Bytes = B3
      
    If Bytes <> 5 Then TotalData = TotalData & " ???Len"
      
    For I = 1 To Bytes
      
        Get #FileNum, Position, B1
        Position = Position + 1
        TotalData = TotalData & HexByte(B1)
         
    Next I
    
ElseIf B2 = &H58 Then

    TotalData = TotalData & "Time sSgn    "
      
    Get #FileNum, Position, B3
    Position = Position + 1
    Bytes = B3
      
    If Bytes <> 4 Then TotalData = TotalData & " ???Len "
      
    Get #FileNum, Position, B4
    Position = Position + 1
      
    Get #FileNum, Position, B5
    Position = Position + 1
      
    TotalData = TotalData & CStr(B4) & "/" & CStr(2 ^ B5) & " - "
      
    Get #FileNum, Position, B4
    Position = Position + 1
    TotalData = TotalData & B4 & " Clocks/Metr.Click - "
      
    Get #FileNum, Position, B5
    Position = Position + 1
    TotalData = TotalData & B5 & " 32nd/Quarter "
      

ElseIf B2 = &H59 Then

    TotalData = TotalData & "Key Sign     "
    
    Get #FileNum, Position, B3
    Position = Position + 1
    Bytes = B3
    
    If Bytes <> 2 Then TotalData = TotalData & " ???Len"
    
    For I = 1 To Bytes
        
        Get #FileNum, Position, B1
        Position = Position + 1
        TotalData = TotalData & HexByte(B1) & " "
        
    Next I

ElseIf B2 = &H7F Then

    Bytes = ReadVariableLength(FileNum, Position)
    TotalData = TotalData & "Propr.- Len  " & CStr(Bytes)
    Position = Position + Bytes

End If

'We Strip it of null chars incase it picked up
'the wrong length, thus if it get's an error
'for this it can show it and have it just end the
'text field.

'NOTE: this wasn't in Stefaan's because
'he displayed to a html page that
'showed after the null/undisplayable
'character either way.
readMidiFF = StripNullChars(TotalData)

End Function

Private Function isNote(ByVal Nr As Long) As String
Dim Octaaf As Long
Dim Note As String

'"octaaf", 8/4?

Octaaf = (Nr \ 12)

Note = Nr Mod 12
isNote = Choose(Note + 1, "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B") & Format(Octaaf - 1)
End Function

' reads Big Endian - variable length variable
Private Function ReadVariableLength(ByVal FileNum As Long, Position As Long) As Long
Dim Value As Long
Dim C As Byte

Get #FileNum, Position, C
Position = Position + 1

Value = C

 If (Value And &H80) <> 0 Then
     
     Value = Value And &H7F
    
     Do
    
         Value = Value * 128
         Get #FileNum, Position, C
         Position = Position + 1
         
         C = C And &H7F
         Value = Value + C
      
     Loop While (C And &H80) <> 0
    
 End If
    
ReadVariableLength = Value
End Function

'Helper function, get's rid of any null bytes causing exceptions.
Private Function StripNullChars(StrStrip As String) As String
Dim I As Integer
Dim Character As Byte

    For I = 1 To Len(StrStrip) 'Reads each char
     
        Character = Asc(Mid$(StrStrip, I, 1)) 'Gets the char
        
        'if it's an undisplayable char but not linefeed/return then doesn't add it.
        If (Character > 31 And Not Character = 13 And Not Character = 10) Then StripNullChars = StripNullChars & Chr(Character)    'adds the char if its displayable/not a null
    
    Next I 'next char

End Function

' hex string of a byte, fixed to 2 digits
Private Function HexByte(ByVal B As Byte) As String
Dim Data As String

Data = Hex(B)

If Len(Data) = 1 Then Data = "0" & Data
HexByte = Data

End Function

'###Class Properties###'

Public Property Get OmitNoteMessages() As Boolean
OmitNoteMessages = FilterNoteMessages
End Property

Public Property Get OmitControlMessages() As Boolean
OmitControlMessages = FilterControlMessages
End Property

Public Property Let OmitNoteMessages(FilterNotes As Boolean)
FilterNoteMessages = FilterNotes
End Property

Public Property Let OmitControlMessages(FilterControls As Boolean)
FilterControlMessages = FilterControls
End Property

Public Property Get OmitSysExMessages() As Boolean
OmitSysExMessages = FilterSysExMessages
End Property

Public Property Let OmitSysExMessages(FilterSysEx As Boolean)
FilterSysExMessages = FilterSysEx
End Property

Public Property Get MaxReportLength() As Long
MaxReportLength = MaxReportLen
End Property

Public Property Let MaxReportLength(MaxReportLength As Long)
MaxReportLen = MaxReportLength
End Property

Private Sub Class_Initialize()
MaxReportLen = 32000
End Sub
