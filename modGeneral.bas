Attribute VB_Name = "modAPIGeneral"
Option Explicit

Public Declare Function MySetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' Get valid path
Public Function validPath(ByVal sPath As String) As String
    If Right(sPath, 1) = "\" Then validPath = sPath Else validPath = sPath & "\"
End Function
' Get short path name
Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long
    Dim sShortPathName As String
    Dim iLen As Integer
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    GetShortName = Left(sShortPathName, lRetVal)
End Function
