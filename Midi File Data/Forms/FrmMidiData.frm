VERSION 5.00
Begin VB.Form FrmMidiData 
   AutoRedraw      =   -1  'True
   Caption         =   "Midi File Data"
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6420
   Icon            =   "FRMMID~1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   4650
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6405
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadMid 
         Caption         =   "Load Mid"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play Mid"
      End
      Begin VB.Menu mnuStopMid 
         Caption         =   "Stop Mid"
      End
      Begin VB.Menu mnubrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "FrmMidiData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public midi As New clsMidi
Public Sound As New clsSound_MID

Public FileLocation As String

Private Sub Form_Load()
midi.MaxReportLength = 64000 'just to show you can set it

'we miss all these as we don't Need them
'they're just if you want them,
'note if you have note messages, you'll
'need to make the report length a lot
'bigger.
midi.OmitControlMessages = True
midi.OmitNoteMessages = True
midi.OmitSysExMessages = True

'txtData.Text = midi.ReadMidi(FrmMain.FileLocation)
End Sub

Private Sub Form_Resize()
'txtData.Width = Me.ScaleWidth
'txtData.Height = Me.ScaleHeight
txtData.Move txtData.Left, txtData.Top, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sound.CloseFile
End Sub

Private Sub mnuLoadMid_Click()
FileLocation = OpenDialog(Me, "*.Mid|*.Mid", "Select a Mid", "")

If FileLocation <> "" Then
    txtData.Text = midi.ReadMidi(FileLocation)
End If

End Sub

Private Sub mnuPlay_Click()
Sound.CloseFile
Sound.PlayFile FileLocation, 800
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub mnuStopMid_Click()
Sound.CloseFile
End Sub
