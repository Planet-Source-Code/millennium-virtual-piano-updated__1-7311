Attribute VB_Name = "modGlobal"
Option Explicit

Public numDevices As Long               ' number of midi output devices
Public curDevice As Long                ' current midi device
Public hmidi As Long                    ' midi output handle
Public rc As Long                       ' return code
Public midiMsg As Long                  ' midi output short message buffer
Public Channel As Integer               ' midi output channel
Public volume As Integer                ' midi velocity
Public pitchPos As Integer              ' pitch position
Public autoPlay As Integer              ' Counter for auto-playing notes
Public lastNote As Integer              ' Last note to stop in preset song
Public LastDrag As Integer              ' Last note to stop in our piano
Public baseNote As Integer              ' the first note on our "piano"
Public locSong As String                ' Location of midi file
Public sendmsg As String                ' SysEx msg
Public lastChordInDrag As Integer       ' Last chord pattern to stop in our piano
Public RetSts As String * 128           ' Buffer for MCI Midi Status
Public YmhDvc As Integer                ' Yamaha MIDI Device
Public YMH_MSG As String                ' Yamaha SysEx ID

Public Const Octave = 0
Public Const Augment = 1
Public Const Augmented7th = 2
Public Const Diminish = 3

Public Const Major = 4
Public Const Major6th = 5
Public Const Major7th = 6
Public Const Major9th = 7
Public Const MajorMin9 = 8

Public Const MajorMin9Suspended4th = 9
Public Const Dominant7th = 10
Public Const Dominant7thMin5 = 11
Public Const Dominant7thSuspended4th = 12
Public Const Suspended4th = 13

Public Const Minor = 14
Public Const Minor6th = 15
Public Const Minor7th = 16
Public Const Minor7thMin5 = 17
Public Const Minor7thMin9 = 18
