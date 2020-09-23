Attribute VB_Name = "modAPIMCI"
Option Explicit

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Const MM_MCINOTIFY = &H3B9
Public Const MCI_NOTIFY_ABORTED = &H4
Public Const MCI_NOTIFY_FAILURE = &H8
Public Const MCI_NOTIFY_SUCCESSFUL = &H1
Public Const MCI_NOTIFY_SUPERSEDED = &H2

Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public pOldProc As Long     ' Old window procedure
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rtnval As Long    ' return value
    Select Case uMsg
        Case MM_MCINOTIFY
            Select Case wParam
                Case MCI_NOTIFY_ABORTED, MCI_NOTIFY_FAILURE, MCI_NOTIFY_SUCCESSFUL, MCI_NOTIFY_SUPERSEDED
                    ' Playback of the MIDI file was somehow aborted
                    ' An error occured while playing the MIDI file
                    ' Playback of the MIDI file concluded successfully
                    ' Another command requested notification from this device
                    rtnval = mciSendString("stop midi", 0&, 0, 0)
                    rtnval = mciSendString("close midi", 0&, 0, 0)
                    Form1.midiStatus = ""
                    Form1.Timer4.Enabled = False
            End Select
            WindowProc = 0
        Case Else
        rtnval = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function

