Attribute VB_Name = "modHVDC_Logging"
Option Explicit

Private Const LOG_FILE_BASENAME As String = "hvdc_pipeline.log"

Public Sub LogEvent(ByVal level As String, ByVal message As String, Optional ByVal context As String = "")
    On Error GoTo CleanFail
    If Not modHVDC_Config.LoggingEnabled Then
        Exit Sub
    End If

    Dim logFile As String
    logFile = modHVDC_FileIO.JoinPath(modHVDC_Config.HVDCLogPath, LOG_FILE_BASENAME)

    modHVDC_FileIO.EnsureFolderExists modHVDC_Config.HVDCLogPath
    RotateLogsIfNeeded logFile

    Dim handle As Integer
    handle = FreeFile
    Open logFile For Append As #handle
    Print #handle, BuildLogLine(level, message, context)
    Close #handle
    Exit Sub
CleanFail:
    On Error Resume Next
    If handle <> 0 Then Close #handle
End Sub

Public Sub LogInfo(ByVal message As String, Optional ByVal context As String = "")
    LogEvent "INFO", message, context
End Sub

Public Sub LogWarn(ByVal message As String, Optional ByVal context As String = "")
    LogEvent "WARN", message, context
End Sub

Public Sub LogError(ByVal message As String, Optional ByVal context As String = "")
    LogEvent "ERROR", message, context
End Sub

Public Sub LogException(ByVal procName As String, ByVal ex As ErrObject, Optional ByVal context As String = "")
    On Error Resume Next
    Dim payload As String
    payload = "{""procedure"":""" & JsonEscape(procName) & """,""number"":""" & Format$(ex.Number, "0") & """,""description"":""" & JsonEscape(ex.Description) & """"}
    LogEvent "ERROR", "Unhandled exception", MergeContext(context, payload)
End Sub

Public Sub LogBatchStats(ByRef stats As HVDCBatchStats, ByVal context As String)
    On Error GoTo CleanFail
    Dim avgLatency As Double
    If stats.ProcessedCount > 0 Then
        avgLatency = stats.LatencyTotalMs / stats.ProcessedCount
    Else
        avgLatency = 0#
    End If
    Dim payload As String
    payload = "{""processed"":""" & Format$(stats.ProcessedCount, "0.00") & """,""exported"":""" & Format$(stats.ExportedCount, "0.00") & """,""skipped"":""" & Format$(stats.SkippedCount, "0.00") & """,""errors"":""" & Format$(stats.ErrorCount, "0.00") & """,""avg_latency_ms"":""" & Format$(avgLatency, "0.00") & """"}
    LogEvent "METRIC", context, payload
    Exit Sub
CleanFail:
    LogException "LogBatchStats", Err, context
End Sub

Private Function BuildLogLine(ByVal level As String, ByVal message As String, ByVal context As String) As String
    Dim ts As String
    ts = Format$(Now, "yyyy-mm-dd\THH:nn:ss")
    Dim line As String
    line = "{""timestamp"":""" & ts & """,""level"":""" & JsonEscape(UCase$(level)) & """,""message"":""" & JsonEscape(message) & """"
    If Len(Trim$(context)) > 0 Then
        line = line & ",""context"":" & context
    End If
    line = line & "}"
    BuildLogLine = line
End Function

Private Sub RotateLogsIfNeeded(ByVal activeFile As String)
    On Error GoTo CleanFail
    If Dir$(activeFile) = "" Then
        Exit Sub
    End If
    If FileLen(activeFile) < modHVDC_Config.LogMaxSizeBytes Then
        Exit Sub
    End If

    Dim limit As Long
    limit = modHVDC_Config.LogFileLimit

    Dim index As Long
    For index = limit - 1 To 1 Step -1
        Dim sourceFile As String
        sourceFile = activeFile & "." & CStr(index)
        If Dir$(sourceFile) <> "" Then
            Name sourceFile As activeFile & "." & CStr(index + 1)
        End If
    Next index

    Dim rotated As String
    rotated = activeFile & ".1"
    Name activeFile As rotated

    Exit Sub
CleanFail:
    ' swallow rotation errors to avoid blocking pipeline
End Sub

Private Function JsonEscape(ByVal value As String) As String
    Dim result As String
    result = Replace(value, "\", "\\")
    result = Replace(result, """", "\"")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    JsonEscape = result
End Function

Private Function MergeContext(ByVal context As String, ByVal addition As String) As String
    Dim trimmed As String
    trimmed = Trim$(context)
    If Len(trimmed) = 0 Then
        MergeContext = addition
    ElseIf Left$(trimmed, 1) = "{" And Right$(trimmed, 1) = "}" Then
        MergeContext = "{" & Trim$(Mid$(trimmed, 2, Len(trimmed) - 2)) & "," & Trim$(Mid$(addition, 2, Len(addition) - 2)) & "}"
    Else
        Dim wrapped As String
        wrapped = "{\"context\":\"" & JsonEscape(trimmed) & "\"}"
        MergeContext = MergeContext(wrapped, addition)
    End If
End Function

