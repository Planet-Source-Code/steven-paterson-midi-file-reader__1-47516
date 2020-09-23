Attribute VB_Name = "modOriginal"
''this is Stefaan Casier's Code of the 2 main function's
''i re-wrote myself, some small functions are
''missing as in other modules i couldn't find since
''this is just to show the difference between learning
''from my code and his

'commented so it doesn't get compile errors

'
'' read midi FF type properties
'Function readMidiFF(ByVal ch As Long, Pos As Long, EndOfTrack As Boolean) As String
'   Dim I As Long, Bytes As Long
'   Dim B As Byte, B2 As Byte, B3 As Byte, B4 As Byte, B5 As Byte
'   Dim txt As String, txt2 As String * 13
'   Get #ch, Pos, B2: Pos = Pos + 1
'   If B2 = 0 Then
'      Get #ch, Pos, B3: Pos = Pos + 1
'      If B3 = 0 Then
'         txt = txt & "seqnr/posfile"
'         Else
'         Get #ch, Pos, B4: Pos = Pos + 1
'         Get #ch, Pos, B5: Pos = Pos + 1
'         txt = txt & "seq nr       " & CStr(B5 * 256 + B4)
'         End If
'   ElseIf B2 >= 1 And B2 <= 7 Then
'      txt2 = Choose(B2, "text", "copyright", "seq/tr. name", "instrument", "lyric", "marker", "cue point") & " - "
'      txt = txt & txt2
'      Bytes = ReadVariableLength(ch, Pos)
'      For I = 1 To Bytes
'         Get #ch, Pos, B: Pos = Pos + 1: txt = txt & Chr(B)
'      Next I
'   ElseIf B2 = &H20 Then
'      txt = txt & "midi chann   "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      Get #ch, Pos, B4: Pos = Pos + 1
'      If B3 <> 0 Then txt = txt & "???len"
'      txt = txt & HexByte(B4)
'   ElseIf B2 = &H21 Then
'      txt = txt & "midi port    "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      Get #ch, Pos, B4: Pos = Pos + 1
'      If B3 <> 0 Then txt = txt & "???len"
'      txt = txt & HexByte(B4)
'   ElseIf B2 = &H2F Then
'      txt = txt & "end of track "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      EndOfTrack = True
'   ElseIf B2 = &H51 Then
'      txt = txt & "tempo        "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      Bytes = B3
'      If Bytes <> 3 Then txt = txt & " ???len "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      Get #ch, Pos, B4: Pos = Pos + 1
'      Get #ch, Pos, B5: Pos = Pos + 1
'      txt = txt & CStr(CLng(60000000 / CLng(CLng(B3) * 256 * 256 + CLng(B4) * 256 + CLng(B5)))) & " BPM"
'   ElseIf B2 = &H54 Then
'      txt = txt & "SMPTE Offs   "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      Bytes = B3
'      If Bytes <> 5 Then txt = txt & " ???len"
'      For I = 1 To Bytes
'         Get #ch, Pos, B: Pos = Pos + 1: txt = txt & HexByte(B)
'      Next I
'   ElseIf B2 = &H58 Then
'      txt = txt & "time sign    "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      Bytes = B3
'      If Bytes <> 4 Then txt = txt & " ???len "
'      Get #ch, Pos, B4: Pos = Pos + 1
'      Get #ch, Pos, B5: Pos = Pos + 1
'      txt = txt & CStr(B4) & "/" & CStr(2 ^ B5) & " - "
'      Get #ch, Pos, B4: Pos = Pos + 1: txt = txt & B4 & " clocks/metr.click - "
'      Get #ch, Pos, B5: Pos = Pos + 1: txt = txt & B5 & " 32nd/quarter "
'   ElseIf B2 = &H59 Then
'      txt = txt & "key sign     "
'      Get #ch, Pos, B3: Pos = Pos + 1
'      Bytes = B3
'      If Bytes <> 2 Then txt = txt & " ???len"
'      For I = 1 To Bytes
'         Get #ch, Pos, B: Pos = Pos + 1: txt = txt & HexByte(B) & " "
'      Next I
'   ElseIf B2 = &H7F Then
'      Bytes = readVarLen(ch, Pos)
'      txt = txt & "propr.- len  " & CStr(Bytes)
'      Pos = Pos + Bytes
'   End If
'   readMidiFF = txt
'End Function
'
'' read and display a midi file
'Public Function readMidiFile(ByVal File As String) As String
'   Dim ch As Long, I As Long
'   Dim txt As String, reg As String, deltaT As String * 7, Stat As String
'   Dim MT As String * 4
'   Dim FormatType As Integer
'   Dim NumTracks As Integer, Track As Integer
'   Dim Division As Integer
'   Dim NumBytes As Long, Bte As Long
'   Dim strBytes As Long
'   Dim Status As Byte
'   Dim Pos As Long, pPos As Long, P As Long
'   Dim Lng As Long
'   Dim Intg As Integer
'   Dim B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte, B5 As Byte, B As Byte
'   Dim DT As Long
'   Dim EndOfTrack As Boolean
'
'   txt = txt & UCase(GetFileTitle(File)) & vbCrLf
'   frmReadMid.SetMax FileLen(File)
'
'   ch = FreeFile
'   Open File For Binary As ch
'   Get #ch, 1, MT
'   If MT <> "MThd" Then txt = "Geen Midi header! ": GoTo ReadMidiFileEND
'   Get #ch, 5, B1
'   Get #ch, 6, B2
'   Get #ch, 7, B3
'   Get #ch, 8, B4
'   If Not (B1 = 0 And B2 = 0 And B3 = 0 And B4 = 6) Then txt = txt & "Midi Header lengte is fout (moet 00 00 00 06 zijn)": GoTo ReadMidiFileEND
'
'   Get #ch, 9, B1
'   Get #ch, 10, B2
'   FormatType = B1 * 256 + B2
'   txt = txt & "Format type = " & CStr(FormatType)
'   Select Case FormatType
'   Case 0: txt = txt & " - single track any channel" & vbCrLf
'   Case 1: txt = txt & " - multi tracks sep channels" & vbCrLf
'   Case 2: txt = txt & " - multi patterns-songs" & vbCrLf
'   Case Else: txt = txt & " - onbekend = fout": GoTo ReadMidiFileEND
'   End Select
'
'   Get #ch, 11, B1
'   Get #ch, 12, B2
'   NumTracks = B1 * 256 + B2
'   txt = txt & "NumTracks = " & CStr(NumTracks) & vbCrLf
'   If FormatType = 0 And NumTracks > 1 Then txt = txt & "Aantal tracks klopt niet met het formaat type = fout.": GoTo ReadMidiFileEND
'
'   Get #ch, 13, B1
'   Get #ch, 14, B2
'   Division = B1 * 256 + B2
'   txt = txt & "Division = " & CStr(Division) & " PPQN" & vbCrLf
'
'   Pos = 15
'   For Track = 1 To NumTracks
'      EndOfTrack = False
'
'      Get #ch, Pos, MT
'      If MT <> "MTrk" Then txt = txt & "Geen Midi track gevonden op de verwachte plaats! ": GoTo ReadMidiFileEND
'      Get #ch, , B1
'      Get #ch, , B2
'      Get #ch, , B3
'      Get #ch, , B4
'      NumBytes = CLng(CLng(B1) * 256 ^ 3 + CLng(B2) * 256 ^ 2 + CLng(B3) * 256 + CLng(B4))
'      txt = txt & vbCrLf & "Track " & CStr(Track) & "     lengte = " & CStr(NumBytes) & vbCrLf
'      Pos = Pos + 8
'      pPos = Pos
'
'      Status = 0
'      While Pos - pPos < NumBytes
'         Get #ch, Pos, B1
'         If B1 = &HFF Then
'            Pos = Pos + 1
'            Status = B1
'            reg = readMidiFF(ch, Pos, EndOfTrack)
'            Stat = " " & HexByte(Status) & " "
'
'            Else
'
'            DT = readVarLen(ch, Pos)
'            deltaT = CStr(DT)
'            Get #ch, Pos, B1
'            If (B1 And &H80) = &H80 Then
'               Status = B1
'               Stat = " " & HexByte(Status) & " "
'               Pos = Pos + 1
'               Else
'               Stat = "r" & HexByte(Status) & " "
'               End If
'            Select Case Status And &HF0
'            Case &H80
'               Get #ch, Pos, B2: Pos = Pos + 1
'               Get #ch, Pos, B3: Pos = Pos + 1
'               If FilterNoteMsg = False Then reg = "Note off.... " & isNote(B2) & "-" & CStr(B3)
'            Case &H90
'               Get #ch, Pos, B2: Pos = Pos + 1
'               Get #ch, Pos, B3: Pos = Pos + 1
'               If FilterNoteMsg = False Then reg = "Note on..... " & isNote(B2) & "-" & CStr(B3)
'            Case &HB0
'               Get #ch, Pos, B2: Pos = Pos + 1
'               Get #ch, Pos, B3: Pos = Pos + 1
'               If FilterCtlChMsg = False Then reg = "Ctl Change.. " & HexByte(B2) & " " & HexByte(B3)
'            Case &HC0
'               Get #ch, Pos, B2: Pos = Pos + 1
'               reg = "Prg Change.. " & HexByte(B2)
'            Case &HD0
'               Get #ch, Pos, B2: Pos = Pos + 1
'               reg = "Chan Press.. " & HexByte(B2)
'            Case &HE0
'               Get #ch, Pos, B2: Pos = Pos + 1
'               Get #ch, Pos, B3: Pos = Pos + 1
'               reg = "Pitch bend.. " & HexByte(B2) & " " & HexByte(B3)
'            Case &HF0
'               Select Case Status
'               Case &HFE
'               Case &HFF
'                  reg = readMidiFF(ch, Pos, EndOfTrack)
'               Case &HF0
'                  P = Pos
'                  Lng = readVarLen(ch, Pos)
'                  If FilterSysExMsg = False Then reg = "SysEx - len: " & CStr(Lng)
'                  Pos = Pos + Lng
'               Case &HF7
'               Case Else
'               End Select
'            End Select
'            End If
'         frmReadMid.SetProgress Pos
'         DoEvents
'         If Cancel = True Then GoTo ReadMidiFileEND:
'         If reg <> "" Then txt = txt & deltaT & Stat & reg & vbCrLf: reg = ""
'         If Len(txt) > 32000 Then txt = txt & vbCrLf & "Tekst te lang...." & vbCrLf: GoTo ReadMidiFileEND:
'      Wend
'ReadMidiFileNEXTTRACK:
'   If Len(txt) > 32000 Then txt = txt & vbCrLf & "Tekst te lang...." & vbCrLf: GoTo ReadMidiFileEND:
'   frmReadMid.SetProgress Pos
'   DoEvents
'   Next Track
'
'ReadMidiFileEND:
'   Close ch
'   readMidiFile = txt
'End Function
'
