Attribute VB_Name = "INIwrite"
Option Explicit

Const INIFILE = "Mavica.cfg"

Private Function WriteIniValue(INIpath As String, PutKey As String, PutVariable As String, PutValue As String)
    Dim Temp As String
    Dim LcaseTemp As String
    Dim ReadKey As String
    Dim ReadVariable As String
    Dim LOKEY As Integer
    Dim HIKEY As Integer
    Dim KEYLEN As Integer
    Dim VAR As Integer
    Dim VARENDOFLINE As Integer
    Dim NF As Integer
    Dim X As Integer

AssignVariables:
    NF = FreeFile
    ReadKey = vbCrLf & "[" & LCase$(PutKey) & "]" & Chr$(13)
    KEYLEN = Len(ReadKey)
    ReadVariable = Chr$(10) & LCase$(PutVariable) & "="
        
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    Temp = Input$(LOF(NF), NF)
    Temp = vbCrLf & Temp & "[]"
    Close NF
    LcaseTemp = LCase$(Temp)
    
LogicMenu:
    LOKEY = InStr(LcaseTemp, ReadKey)
    If LOKEY = 0 Then GoTo AddKey:
    HIKEY = InStr(LOKEY + KEYLEN, LcaseTemp, "[")
    VAR = InStr(LOKEY, LcaseTemp, ReadVariable)
    If VAR > HIKEY Or VAR < LOKEY Then GoTo AddVariable:
    GoTo RenewVariable:
    
AddKey:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Temp & vbCrLf & vbCrLf & "[" & PutKey & "]" & vbCrLf & PutVariable & "=" & PutValue
        GoTo TrimFinalString:
        
AddVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Left$(Temp, LOKEY + KEYLEN) & PutVariable & "=" & PutValue & vbCrLf & Mid$(Temp, LOKEY + KEYLEN + 1)
        GoTo TrimFinalString:
        
RenewVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        VARENDOFLINE = InStr(VAR, Temp, Chr$(13))
        Temp = Left$(Temp, VAR) & PutVariable & "=" & PutValue & Mid$(Temp, VARENDOFLINE)
        GoTo TrimFinalString:

TrimFinalString:
        Temp = Mid$(Temp, 2)
        Do Until InStr(Temp, vbCrLf & vbCrLf & vbCrLf) = 0
        Temp = Replace(Temp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        Loop
    
        Do Until Right$(Temp, 1) > Chr$(13)
        Temp = Left$(Temp, Len(Temp) - 1)
        Loop
    
        Do Until Left$(Temp, 1) > Chr$(13)
        Temp = Mid$(Temp, 2)
        Loop
    
OutputAmendedINIFile:
        Open INIpath For Output As NF
        Print #NF, Temp
        Close NF
    
End Function

Public Sub CheckFileExists()
    If Dir(App.path & "\" & INIFILE, vbArchive + vbHidden) = "" Then
        WriteIniValue App.path & "\" & INIFILE, "Disk", "Drive", "A:\"
        WriteIniValue App.path & "\" & INIFILE, "Import", "Format", "0"
        WriteIniValue App.path & "\" & INIFILE, "Import", "Folder", "C:\"
    End If
End Sub

Public Sub WriteDrive(drive As String)
    WriteIniValue App.path & "\" & INIFILE, "Disk", "Drive", drive
End Sub

Public Sub WriteFormat(format As String)
    WriteIniValue App.path & "\" & INIFILE, "Import", "Format", format
End Sub

Public Sub WriteFolder(folder As String)
    WriteIniValue App.path & "\" & INIFILE, "Import", "Folder", folder
End Sub

Public Sub RestoreFile()
    Kill App.path & "\" & INIFILE
    CheckFileExists
End Sub

Public Sub RestoreFolder()
    WriteIniValue App.path & "\" & INIFILE, "Import", "Folder", "C:\"
End Sub
