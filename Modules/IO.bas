Attribute VB_Name = "IO"
Option Explicit

Dim drive As String
Dim files() As String
Dim counter As Integer

Const path = "mvc-*.jpg"
Const MY_NAME = "Mafovi.IO"

'Setting the read drive
Public Sub Initialize(ByRef frm As frmMain)
    drive = INIread.ReadDrive
    
    frm.drvActive.drive = drive
    frm.drvImport.drive = UCase$(Left$(App.path, 1) & ":\")
    frm.Caption = drive & " drive - " & App.Title
End Sub

'Load the filenames into files array
Public Sub FileLoad(ByRef frm As frmMain)
On Error GoTo errorHandler
    Dim total As Integer
    
    frm.MousePointer = vbHourglass
    total = ImageCount()
    
    If total = 0 Then
        Err.Raise 513, MY_NAME, "The drive does not contain any SONY Mavica images"
    Else
        ReDim files(total - 1)
        
        Dim filename As String
        
        counter = 0
        filename = Dir(drive & path)
        
        Do While filename <> ""
            files(counter) = filename
            filename = Dir()
            counter = counter + 1
        Loop
        
        counter = -1
        
        Call IO.ChangeImage(1, frm.imgRect)
    End If

errorHandler:
    If Err.Number = 52 Then frm.imgRect.Picture = Nothing: MsgBox "Drive not ready", vbExclamation + vbOKOnly, "Drive error"
    If Err.Number = 75 Then Exit Sub
    If Err.Number = 513 Then frm.imgRect.Picture = Nothing: MsgBox Err.Description, vbExclamation + vbOKOnly
    
    frm.MousePointer = vbDefault   'Set mouse pointer back to default, even if an error occurs
End Sub

Public Sub ChangeDrive(drv As String, ByRef frm As frmMain)
    INIwrite.WriteDrive UCase$(Left$(drv, 1) & ":\")
    Call Initialize(frm)
End Sub

'Procedure which changes the current image
'Direction is 1 for advancing and -1 for going back!
Public Sub ChangeImage(direction As Integer, ByRef pic As Image)
On Error GoTo errorHandler
    counter = counter + direction
    
errorHandler:
    Select Case Err.Number
        'Error 9: Negative counter
        Case 9
            If counter < 0 Then
                counter = UBound(files)
            Else
                counter = 0
            End If
        'Error 71: File deleted / disk removed
        Case 71
            MsgBox "File has been deleted or the disk has been removed from drive", _
                    vbCritical + vbOKOnly, "File Error"
            counter = counter - 1
            Exit Sub
        'Error 75, 76: Counter out of bounds
        Case 75 To 76
            counter = 0
        'Error 53: Not found
        Case 53
            MsgBox "The file could not be found", vbCritical + vbOKOnly, "File error"
            Exit Sub
        'Error 481: Invalid format
        Case 481
            MsgBox "The file format is not compatible", vbCritical + vbOKOnly, "Format error"
            Exit Sub
        'Other errors: raise them
        Case 0
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    
    pic.Picture = LoadPicture(drive & files(counter))
End Sub

Public Function ImageCount() As Integer
    Dim filename As String
    Dim total As Integer
    
    total = 0
    
    filename = Dir(drive & path)
    
    Do While filename <> ""
        total = total + 1   'Counting every file which is the same as the path variable
        filename = Dir()
    Loop
    
    ImageCount = total
End Function

Public Function DeleteImg() As Boolean
    Dim del As Boolean
    del = False
    
    If ImageCount > 0 Then
        If MsgBox("Delete file " & drive & files(counter) & "?", vbYesNo + vbQuestion) = vbYes Then
            Call DeleteFile(counter)
            del = True
        End If
    Else
        MsgBox "No files to delete", vbExclamation + vbOKOnly
    End If
    
    DeleteImg = del
End Function

Public Function DeleteDisk() As Boolean
    Dim del As Boolean
    del = False
    
    If ImageCount > 0 Then
        If MsgBox("Delete all files on " & drive & "?", vbYesNo + vbQuestion) = vbYes Then
            Dim i As Integer
            
            For i = 0 To UBound(files)
                DeleteFile (i)
            Next i
            del = True
        End If
    Else
        MsgBox "No files to delete", vbExclamation + vbOKOnly
    End If
    
    DeleteDisk = del
End Function

Private Sub DeleteFile(fileNum As Integer)
On Error GoTo errorHandler
    Dim curFile As String
    
    curFile = UCase$(Left$(files(fileNum), 8))
    
    SetAttr drive & curFile & ".411", vbNormal      ' Change attributes for 411 files, so they can be deleted
    
    Kill drive & curFile & ".*"

    Exit Sub
errorHandler:
    Select Case Err.Number
        Case 75
            MsgBox "Error while deleting" & vbNewLine & "The disk might be write protected", vbOKOnly + vbExclamation, "Delete error"
        Case 53
            Resume Next
    End Select
End Sub

'File import, destination and folder name format
Public Sub FileImport(dest As String, folder_format As String)
On Error GoTo errorHandler
    If ImageCount = 0 Then
        Err.Raise 514, MY_NAME, "No files on disk, cannot continue"
    Else
        Dim curFile As String
        Dim fileDuplicate As Integer
        Dim i As Integer
            
        For i = 0 To UBound(files)
            If Dir(dest & format(FileDateTime(drive & files(i)), folder_format) & "\") = "" Then
                MkDir (dest & format(FileDateTime(drive & files(i)), folder_format) & "\")
            End If
            
            curFile = UCase$(Left$(files(i), 8))
            fileDuplicate = 1
            
            'We search for
            Do While Dir(dest & format(FileDateTime(drive & files(i)), folder_format) & "\" & curFile & ".JPG") <> ""
                curFile = UCase$(Left$(files(i), 8) & " (" & fileDuplicate & ")")
                fileDuplicate = fileDuplicate + 1
            Loop
            
            FileCopy drive & files(i), dest & format(FileDateTime(drive & files(i)), folder_format) & "\" & curFile & ".JPG"
        Next i
        
        If MsgBox("Image import completed" & vbNewLine & "Do you want to open the output folder?", _
                vbYesNo + vbQuestion, "Import complete") = vbYes Then
            Shell "explorer " & dest, vbNormalFocus
        End If
    End If
    
    Exit Sub
errorHandler:
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
