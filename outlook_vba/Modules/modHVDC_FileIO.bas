Attribute VB_Name = "modHVDC_FileIO"
Option Explicit

Private Const TEMP_PREFIX As String = "hvdc_tmp_"

Public Function JoinPath(ByVal parentPath As String, ByVal childPath As String) As String
    Dim trimmedParent As String
    Dim trimmedChild As String
    trimmedParent = Trim$(parentPath)
    trimmedChild = Trim$(childPath)

    If Len(trimmedParent) = 0 Then
        JoinPath = trimmedChild
        Exit Function
    End If

    If Len(trimmedChild) = 0 Then
        JoinPath = trimmedParent
        Exit Function
    End If

    If Right$(trimmedParent, 1) = "\" Or Right$(trimmedParent, 1) = "/" Then
        trimmedParent = Left$(trimmedParent, Len(trimmedParent) - 1)
    End If

    If Left$(trimmedChild, 1) = "\" Or Left$(trimmedChild, 1) = "/" Then
        trimmedChild = Mid$(trimmedChild, 2)
    End If

    JoinPath = trimmedParent & "\" & trimmedChild
End Function

Public Sub EnsureFolderExists(ByVal targetPath As String)
    On Error GoTo CleanFail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(Trim$(targetPath)) = 0 Then Exit Sub
    If Not fso.FolderExists(targetPath) Then
        Dim parentPath As String
        parentPath = fso.GetParentFolderName(targetPath)
        If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
            EnsureFolderExists parentPath
        End If
        fso.CreateFolder targetPath
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "EnsureFolderExists", Err, targetPath
End Sub

Public Function BuildDatedExportFolder(ByVal receivedTime As Date) As String
    Dim safeTime As Date
    If receivedTime = 0 Then
        safeTime = Now
    Else
        safeTime = receivedTime
    End If

    Dim baseFolder As String
    baseFolder = modHVDC_Config.HVDCExportRoot
    Dim yearFolder As String
    yearFolder = JoinPath(baseFolder, Format$(safeTime, "yyyy"))
    EnsureFolderExists yearFolder

    Dim monthFolder As String
    monthFolder = JoinPath(yearFolder, Format$(safeTime, "mm"))
    EnsureFolderExists monthFolder

    Dim dayFolder As String
    dayFolder = JoinPath(monthFolder, Format$(safeTime, "dd"))
    EnsureFolderExists dayFolder

    BuildDatedExportFolder = dayFolder
End Function

Public Function BuildMessageFileName(ByVal mail As Outlook.MailItem) As String
    Dim safeTime As Date
    If mail.ReceivedTime = 0 Then
        safeTime = Now
    Else
        safeTime = mail.ReceivedTime
    End If
    Dim entryHash As String
    entryHash = EntryIdHash(mail.EntryID)
    BuildMessageFileName = Format$(safeTime, "yyyy-mm-dd_hh-nn-ss") & "_" & entryHash & ".msg"
End Function

Public Function ComposeMessageFilePath(ByVal mail As Outlook.MailItem) As String
    Dim targetFolder As String
    targetFolder = BuildDatedExportFolder(mail.ReceivedTime)
    ComposeMessageFilePath = JoinPath(targetFolder, BuildMessageFileName(mail))
End Function

Public Function ComposeMetadataFilePath(ByVal messagePath As String) As String
    ComposeMetadataFilePath = Left$(messagePath, Len(messagePath) - 4) & ".json"
End Function

Public Function ComposeAttachmentsFolder(ByVal entryId As String, ByVal datedFolder As String) As String
    Dim attachmentsRoot As String
    attachmentsRoot = JoinPath(datedFolder, "attachments")
    EnsureFolderExists attachmentsRoot
    Dim folderName As String
    folderName = SanitizeFolderName(entryId)
    Dim finalFolder As String
    finalFolder = JoinPath(attachmentsRoot, folderName)
    EnsureFolderExists finalFolder
    ComposeAttachmentsFolder = finalFolder
End Function

Public Function SaveMailItem(ByVal mail As Outlook.MailItem, ByVal targetPath As String) As Boolean
    On Error GoTo CleanFail
    Dim tempPath As String
    tempPath = BuildTempFilePath(targetPath)
    mail.SaveAs tempPath, olMSGUnicode
    MoveWithRetry tempPath, targetPath
    SaveMailItem = True
    Exit Function
CleanFail:
    SaveMailItem = False
    modHVDC_Logging.LogException "SaveMailItem", Err, targetPath
    On Error Resume Next
    If Len(Dir$(tempPath)) > 0 Then
        Kill tempPath
    End If
End Function

Public Function SaveAttachment(ByVal attachment As Outlook.Attachment, ByVal targetFolder As String) As String
    On Error GoTo CleanFail
    Dim sanitizedName As String
    sanitizedName = SanitizeFileName(attachment.FileName)
    Dim finalPath As String
    finalPath = JoinPath(targetFolder, sanitizedName)
    Dim tempPath As String
    tempPath = BuildTempFilePath(finalPath)
    attachment.SaveAsFile tempPath
    MoveWithRetry tempPath, finalPath
    SaveAttachment = finalPath
    Exit Function
CleanFail:
    modHVDC_Logging.LogException "SaveAttachment", Err, targetFolder
    SaveAttachment = ""
    On Error Resume Next
    If Len(tempPath) > 0 And Dir$(tempPath) <> "" Then
        Kill tempPath
    End If
End Function

Public Sub WriteJsonMetadata(ByVal jsonContent As String, ByVal targetPath As String)
    On Error GoTo CleanFail
    WriteTextAtomic targetPath, jsonContent
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "WriteJsonMetadata", Err, targetPath
End Sub

Public Sub WriteTextAtomic(ByVal targetPath As String, ByVal content As String)
    On Error GoTo CleanFail
    Dim tempPath As String
    tempPath = BuildTempFilePath(targetPath)

    Dim handle As Integer
    handle = FreeFile
    Open tempPath For Output As #handle
    Print #handle, content
    Close #handle

    MoveWithRetry tempPath, targetPath
    Exit Sub
CleanFail:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    If Len(tempPath) > 0 And Dir$(tempPath) <> "" Then Kill tempPath
    modHVDC_Logging.LogException "WriteTextAtomic", Err, targetPath
End Sub

Public Function ReadAllText(ByVal sourcePath As String) As String
    On Error GoTo CleanFail
    If Dir$(sourcePath) = "" Then
        Exit Function
    End If
    Dim handle As Integer
    handle = FreeFile
    Open sourcePath For Input As #handle
    ReadAllText = Input$(LOF(handle), handle)
    Close #handle
    Exit Function
CleanFail:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    modHVDC_Logging.LogException "ReadAllText", Err, sourcePath
End Function

Public Function ReadCheckpointTimestamp() As Date
    On Error GoTo CleanFail
    Dim rawText As String
    rawText = ReadAllText(modHVDC_Config.HVDCCheckpointFile)
    If Len(Trim$(rawText)) = 0 Then
        ReadCheckpointTimestamp = CDate("1970-01-01 00:00:00")
        Exit Function
    End If
    Dim marker As String
    marker = "last_processed"
    Dim position As Long
    position = InStr(1, rawText, marker, vbTextCompare)
    If position = 0 Then
        ReadCheckpointTimestamp = CDate("1970-01-01 00:00:00")
        Exit Function
    End If
    Dim startQuote As Long
    startQuote = InStr(position, rawText, ":", vbTextCompare)
    If startQuote = 0 Then
        ReadCheckpointTimestamp = CDate("1970-01-01 00:00:00")
        Exit Function
    End If
    startQuote = InStr(startQuote, rawText, """", vbTextCompare)
    If startQuote = 0 Then
        ReadCheckpointTimestamp = CDate("1970-01-01 00:00:00")
        Exit Function
    End If
    Dim endQuote As Long
    endQuote = InStr(startQuote + 1, rawText, """", vbTextCompare)
    If endQuote = 0 Then
        ReadCheckpointTimestamp = CDate("1970-01-01 00:00:00")
        Exit Function
    End If
    Dim isoValue As String
    isoValue = Mid$(rawText, startQuote + 1, endQuote - startQuote - 1)
    ReadCheckpointTimestamp = CDate(Replace(Mid$(isoValue, 1, 19), "T", " "))
    Exit Function
CleanFail:
    modHVDC_Logging.LogException "ReadCheckpointTimestamp", Err, ""
    ReadCheckpointTimestamp = CDate("1970-01-01 00:00:00")
End Function

Public Sub WriteCheckpointTimestamp(ByVal timestampValue As Date)
    On Error GoTo CleanFail
    Dim isoValue As String
    isoValue = Format$(timestampValue, "yyyy-mm-dd\THH:nn:ss")
    Dim jsonContent As String
    jsonContent = "{""last_processed"":""" & isoValue & """}"
    WriteTextAtomic modHVDC_Config.HVDCCheckpointFile, jsonContent
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "WriteCheckpointTimestamp", Err, ""
End Sub

Public Sub EnsureCheckpointFile(ByVal checkpointPath As String)
    On Error GoTo CleanFail
    If Dir$(checkpointPath) = "" Then
        WriteCheckpointTimestamp CDate("1970-01-01 00:00:00")
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "EnsureCheckpointFile", Err, checkpointPath
End Sub

Public Sub LoadProcessedCache(ByVal cache As Object)
    On Error GoTo CleanFail
    Dim cachePath As String
    cachePath = modHVDC_Config.HVDCProcessedCacheFile
    If Dir$(cachePath) = "" Then
        Exit Sub
    End If
    Dim raw As String
    raw = ReadAllText(cachePath)
    If Len(raw) = 0 Then Exit Sub
    Dim lines As Variant
    lines = Split(raw, vbCrLf)
    Dim line As Variant
    For Each line In lines
        Dim trimmed As String
        trimmed = Trim$(CStr(line))
        If Len(trimmed) > 0 Then
            If Not cache.Exists(trimmed) Then
                cache.Add trimmed, True
            End If
        End If
    Next line
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "LoadProcessedCache", Err, cachePath
End Sub

Public Sub PersistProcessedCache(ByVal cache As Object)
    On Error GoTo CleanFail
    Dim cachePath As String
    cachePath = modHVDC_Config.HVDCProcessedCacheFile
    Dim builder As String
    Dim key As Variant
    For Each key In cache.Keys
        builder = builder & CStr(key) & vbCrLf
    Next key
    WriteTextAtomic cachePath, builder
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "PersistProcessedCache", Err, cachePath
End Sub

Private Sub MoveWithRetry(ByVal sourcePath As String, ByVal targetPath As String)
    On Error GoTo CleanFail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim attempt As Long
    Dim delayMs As Long
    delayMs = 50
    For attempt = 1 To 5
        On Error Resume Next
        If fso.FileExists(targetPath) Then
            fso.DeleteFile targetPath, True
        End If
        On Error GoTo RetryFailure
        fso.MoveFile sourcePath, targetPath
        Exit For
RetryFailure:
        On Error GoTo CleanFail
        If attempt = 5 Then
            Err.Raise vbObjectError + 513, "MoveWithRetry", "Unable to move file after retries"
        End If
        SleepMilliseconds delayMs
        delayMs = delayMs * 2
    Next attempt
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "MoveWithRetry", Err, targetPath
    On Error Resume Next
    If fso.FileExists(sourcePath) Then
        fso.DeleteFile sourcePath, True
    End If
End Sub

Private Function BuildTempFilePath(ByVal targetPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempFolder As String
    tempFolder = Environ$("TEMP")
    If Len(tempFolder) = 0 Then
        tempFolder = modHVDC_Config.HVDCStatePath
    End If
    EnsureFolderExists tempFolder
    Dim tempName As String
    tempName = TEMP_PREFIX & CStr(Timer * 1000#) & "_" & EntryIdHash(targetPath) & ".tmp"
    BuildTempFilePath = JoinPath(tempFolder, tempName)
End Function

Private Sub SleepMilliseconds(ByVal durationMs As Long)
    Dim endTick As Single
    endTick = Timer + (durationMs / 1000#)
    Do While Timer < endTick
        DoEvents
    Loop
End Sub

Public Function SanitizeFileName(ByVal candidate As String) As String
    Dim sanitized As String
    sanitized = candidate
    sanitized = Replace(sanitized, "\", "_")
    sanitized = Replace(sanitized, "/", "_")
    sanitized = Replace(sanitized, ":", "_")
    sanitized = Replace(sanitized, "*", "_")
    sanitized = Replace(sanitized, "?", "_")
    sanitized = Replace(sanitized, """", "_")
    sanitized = Replace(sanitized, "<", "_")
    sanitized = Replace(sanitized, ">", "_")
    sanitized = Replace(sanitized, "|", "_")
    If Len(sanitized) = 0 Then
        sanitized = "attachment.bin"
    End If
    SanitizeFileName = sanitized
End Function

Private Function SanitizeFolderName(ByVal candidate As String) As String
    Dim sanitized As String
    sanitized = SanitizeFileName(candidate)
    If Len(sanitized) > 60 Then
        sanitized = Left$(sanitized, 60)
    End If
    SanitizeFolderName = sanitized
End Function

Public Function EntryIdHash(ByVal value As String) As String
    Dim table As Variant
    table = Crc32Table()
    Dim crc As Long
    crc = &HFFFFFFFF
    Dim idx As Long
    For idx = 1 To Len(value)
        Dim byteValue As Long
        byteValue = Asc(Mid$(value, idx, 1)) And &HFF&
        crc = ((crc And &HFFFFFF00) \ &H100) Xor table((crc Xor byteValue) And &HFF&)
    Next idx
    crc = Not crc
    EntryIdHash = Right$("00000000" & Hex$(crc And &HFFFFFFFF), 8)
End Function

Private Function Crc32Table() As Variant
    Static table As Variant
    If IsEmpty(table) Then
        Dim poly As Long
        poly = &HEDB88320
        ReDim table(0 To 255)

        Dim i As Long
        Dim j As Long
        Dim crc As Long

        For i = 0 To 255
            crc = i
            For j = 0 To 7
                If (crc And 1) <> 0 Then
                    crc = (crc \ 2) Xor poly
                Else
                    crc = crc \ 2
                End If
            Next j
            table(i) = crc And &HFFFFFFFF
        Next i
    End If
    Crc32Table = table
End Function

