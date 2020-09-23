Attribute VB_Name = "modResFile"
Option Explicit

Public Function dumpResFile()
    Dim iResourceNum As Integer
    Dim sDestFileName As String
    Dim iFileNum As Integer
    Dim bytResourceData() As Byte
    Dim iFileNumOut As Integer
    
    iResourceNum = 101
    sDestFileName = "INS.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    bytResourceData = LoadResData(iResourceNum, "Custom")
    iFileNum = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNum
        Put #iFileNum, , bytResourceData
    Close #iFileNum
    
    iResourceNum = 102
    sDestFileName = "DRUM.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    bytResourceData = LoadResData(iResourceNum, "Custom")
    iFileNum = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNum
        Put #iFileNum, , bytResourceData
    Close #iFileNum
    
    iResourceNum = 103
    sDestFileName = "SFX.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    bytResourceData = LoadResData(iResourceNum, "Custom")
    iFileNum = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNum
        Put #iFileNum, , bytResourceData
    Close #iFileNum
    
    iResourceNum = 104
    sDestFileName = "VL.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    bytResourceData = LoadResData(iResourceNum, "Custom")
    iFileNum = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNum
        Put #iFileNum, , bytResourceData
    Close #iFileNum
    
    iResourceNum = 105
    sDestFileName = "SONG.MID"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    bytResourceData = LoadResData(iResourceNum, "Custom")
    iFileNum = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNum
        Put #iFileNum, , bytResourceData
    Close #iFileNum
End Function
Public Function killResFile()
    Dim sDestFileName As String
    sDestFileName = "INS.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    sDestFileName = "DRUM.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    sDestFileName = "SFX.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    sDestFileName = "VL.TXT"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
    sDestFileName = "SONG.MID"
    If Dir(sDestFileName) <> "" Then Kill sDestFileName
End Function
