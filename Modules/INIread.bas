Attribute VB_Name = "INIread"
Option Explicit

Const INIFILE = "Mavica.cfg"

Private Function ReadIniValue(INIpath As String, KEY As String, Variable As String) As String
    Dim NF As Integer
    Dim Temp As String
    Dim LcaseTemp As String
    Dim ReadyToRead As Boolean
    
AssignVariables:
        NF = FreeFile
        ReadIniValue = ""
        KEY = "[" & LCase$(KEY) & "]"
        Variable = LCase$(Variable)
    
EnsureFileExists:
    INIwrite.CheckFileExists
    
LoadFile:
    Open INIpath For Input As NF
    While Not EOF(NF)
    Line Input #NF, Temp
    LcaseTemp = LCase$(Temp)
    If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
    If LcaseTemp = KEY Then ReadyToRead = True
    If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
        If InStr(LcaseTemp, Variable & "=") = 1 Then
            ReadIniValue = Mid$(Temp, 1 + Len(Variable & "="))
            Close NF: Exit Function
            End If
        End If
    Wend
    Close NF
End Function

Public Function ReadDrive()
    ReadDrive = ReadIniValue(App.path & "\" & INIFILE, "Disk", "Drive")
End Function

Public Function ReadFormat()
    ReadFormat = ReadIniValue(App.path & "\" & INIFILE, "Import", "Format")
End Function

Public Function ReadFolder()
    ReadFolder = ReadIniValue(App.path & "\" & INIFILE, "Import", "Folder")
End Function