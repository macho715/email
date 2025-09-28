Attribute VB_Name = "modHVDC_Config"
Option Explicit

Private Const LOG_MAX_SIZE_BYTES As Long = 10485760 '10 MB
Private Const LOG_FILE_COUNT As Long = 5
Private Const DEFAULT_BATCH_SIZE As Long = 25
Private Const DEFAULT_QUEUE_INTERVAL_SECONDS As Long = 30
Private Const HOURLY_TASK_NAME As String = "HVDC Hourly Tick"
Private Const QUEUE_TASK_NAME As String = "HVDC Queue Pulse"
Private Const CONFIG_FOLDER_NAME As String = "HVDC"
Private Const EXPORT_FOLDER_NAME As String = "exports"
Private Const LOG_FOLDER_NAME As String = "logs"
Private Const STATE_FOLDER_NAME As String = "state"
Private Const CHECKPOINT_FILE_NAME As String = "catchup_checkpoint.json"
Private Const PROCESSED_CACHE_NAME As String = "processed_index.json"
Private Const ADDITIONAL_FOLDERS As String = "" 'Example: "Inbox\\HVDC Alerts|Inbox\\Operations"

Private m_nextQueuePulse As Date
Private m_processedCache As Object

Public Function LogMaxSizeBytes() As Long
    LogMaxSizeBytes = LOG_MAX_SIZE_BYTES
End Function

Public Function LogFileLimit() As Long
    LogFileLimit = LOG_FILE_COUNT
End Function

Public Function BatchSizeLimit() As Long
    BatchSizeLimit = DEFAULT_BATCH_SIZE
End Function

Public Function QueueIntervalSeconds() As Long
    QueueIntervalSeconds = DEFAULT_QUEUE_INTERVAL_SECONDS
End Function

Public Function HourlyTaskName() As String
    HourlyTaskName = HOURLY_TASK_NAME
End Function

Public Function QueueTaskName() As String
    QueueTaskName = QUEUE_TASK_NAME
End Function

Public Function HVDCBasePath() As String
    Dim basePath As String
    basePath = Trim$(Environ$("LOCALAPPDATA"))
    If Len(basePath) = 0 Then
        basePath = Trim$(Environ$("USERPROFILE"))
    End If
    If Len(basePath) = 0 Then
        basePath = "C:\\HVDC"
    End If
    HVDCBasePath = modHVDC_FileIO.JoinPath(basePath, CONFIG_FOLDER_NAME)
End Function

Public Function HVDCExportRoot() As String
    HVDCExportRoot = modHVDC_FileIO.JoinPath(HVDCBasePath(), EXPORT_FOLDER_NAME)
End Function

Public Function HVDCLogPath() As String
    HVDCLogPath = modHVDC_FileIO.JoinPath(HVDCBasePath(), LOG_FOLDER_NAME)
End Function

Public Function HVDCStatePath() As String
    HVDCStatePath = modHVDC_FileIO.JoinPath(HVDCBasePath(), STATE_FOLDER_NAME)
End Function

Public Function HVDCCheckpointFile() As String
    HVDCCheckpointFile = modHVDC_FileIO.JoinPath(HVDCStatePath(), CHECKPOINT_FILE_NAME)
End Function

Public Function HVDCProcessedCacheFile() As String
    HVDCProcessedCacheFile = modHVDC_FileIO.JoinPath(HVDCStatePath(), PROCESSED_CACHE_NAME)
End Function

Public Function HVDCProcessedCache() As Object
    If m_processedCache Is Nothing Then
        Set m_processedCache = CreateObject("Scripting.Dictionary")
        m_processedCache.CompareMode = vbTextCompare
    End If
    Set HVDCProcessedCache = m_processedCache
End Function

Public Sub ResetProcessedCache()
    If Not m_processedCache Is Nothing Then
        m_processedCache.RemoveAll
    End If
End Sub

Public Function LoggingEnabled() As Boolean
    LoggingEnabled = True
End Function

Public Sub InitializeHVDC()
    On Error GoTo CleanFail
    EnsureHVDCInfrastructure
    EnsureReminderTasks
    modHVDC_FileIO.LoadProcessedCache HVDCProcessedCache
    modHVDC_FileIO.EnsureCheckpointFile HVDCCheckpointFile
    RequestQueuePulse
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "InitializeHVDC", Err, ""
End Sub

Public Sub EnsureHVDCInfrastructure()
    On Error GoTo CleanFail
    modHVDC_FileIO.EnsureFolderExists HVDCBasePath
    modHVDC_FileIO.EnsureFolderExists HVDCExportRoot
    modHVDC_FileIO.EnsureFolderExists HVDCLogPath
    modHVDC_FileIO.EnsureFolderExists HVDCStatePath
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "EnsureHVDCInfrastructure", Err, ""
End Sub

Public Sub EnsureReminderTasks()
    On Error GoTo CleanFail
    EnsureHiddenTask HourlyTaskName, True
    EnsureHiddenTask QueueTaskName, False
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "EnsureReminderTasks", Err, ""
End Sub

Private Sub EnsureHiddenTask(ByVal taskName As String, ByVal isRecurringHourly As Boolean)
    On Error GoTo CleanFail
    Dim session As Outlook.NameSpace
    Set session = Application.Session
    Dim tasksFolder As Outlook.MAPIFolder
    Set tasksFolder = session.GetDefaultFolder(olFolderTasks)

    Dim items As Outlook.Items
    Set items = tasksFolder.Items

    Dim task As Outlook.TaskItem
    Dim index As Long
    For index = 1 To items.Count
        If TypeOf items(index) Is Outlook.TaskItem Then
            Set task = items(index)
            If StrComp(task.Subject, taskName, vbTextCompare) = 0 Then
                Exit For
            Else
                Set task = Nothing
            End If
        End If
    Next index

    If task Is Nothing Then
        Set task = items.Add(olTaskItem)
        With task
            .Subject = taskName
            .Categories = "HVDC Hidden"
            .ReminderSet = True
            .ReminderOverrideDefault = True
            .ReminderPlaySound = False
            .ReminderSoundFile = ""
            .BusyStatus = olFree
            .Importance = olImportanceLow
            .Sensitivity = olConfidential
            .Body = "HVDC automation support item. Do not delete."
            .ReminderSet = True
            If isRecurringHourly Then
                .ReminderTime = DateAdd("h", 1, Now)
            Else
                .ReminderTime = Now + (QueueIntervalSeconds / 86400#)
            End If
            .Save
        End With
    End If

    If isRecurringHourly Then
        Dim pattern As Outlook.RecurrencePattern
        Set pattern = task.GetRecurrencePattern
        With pattern
            .RecurrenceType = olRecursHourly
            .Interval = 1
            .PatternStartDate = Date
            .Occurrences = 0
        End With
        If task.ReminderTime < Now Then
            task.ReminderTime = DateAdd("h", 1, Now)
        End If
    Else
        If task.ReminderTime < Now Then
            task.ReminderTime = Now + (QueueIntervalSeconds / 86400#)
        End If
    End If
    task.Save
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "EnsureHiddenTask", Err, "task=" & taskName
End Sub

Public Sub RequestQueuePulse()
    On Error GoTo CleanFail
    Dim nextDue As Date
    nextDue = Now + (QueueIntervalSeconds / 86400#)
    If m_nextQueuePulse = 0# Or nextDue < m_nextQueuePulse - (QueueIntervalSeconds / 86400#) Then
        UpdateQueuePulseReminder nextDue
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "RequestQueuePulse", Err, ""
End Sub

Public Sub UpdateHourlyReminder()
    On Error GoTo CleanFail
    Dim session As Outlook.NameSpace
    Set session = Application.Session
    Dim task As Outlook.TaskItem
    Set task = FindHiddenTask(HourlyTaskName)
    If task Is Nothing Then
        EnsureHiddenTask HourlyTaskName, True
        Set task = FindHiddenTask(HourlyTaskName)
    End If
    If Not task Is Nothing Then
        task.ReminderTime = DateAdd("h", 1, Now)
        task.Save
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "UpdateHourlyReminder", Err, ""
End Sub

Private Sub UpdateQueuePulseReminder(ByVal reminderTime As Date)
    On Error GoTo CleanFail
    Dim task As Outlook.TaskItem
    Set task = FindHiddenTask(QueueTaskName)
    If task Is Nothing Then
        EnsureHiddenTask QueueTaskName, False
        Set task = FindHiddenTask(QueueTaskName)
    End If
    If Not task Is Nothing Then
        task.ReminderTime = reminderTime
        task.Save
        m_nextQueuePulse = reminderTime
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "UpdateQueuePulseReminder", Err, ""
End Sub

Private Function FindHiddenTask(ByVal taskName As String) As Outlook.TaskItem
    On Error GoTo CleanFail
    Dim tasks As Outlook.MAPIFolder
    Set tasks = Application.Session.GetDefaultFolder(olFolderTasks)
    Dim items As Outlook.Items
    Set items = tasks.Items
    Dim index As Long
    For index = 1 To items.Count
        If TypeOf items(index) Is Outlook.TaskItem Then
            Dim candidate As Outlook.TaskItem
            Set candidate = items(index)
            If StrComp(candidate.Subject, taskName, vbTextCompare) = 0 Then
                Set FindHiddenTask = candidate
                Exit Function
            End If
        End If
    Next index
    Exit Function
CleanFail:
    modHVDC_Logging.LogException "FindHiddenTask", Err, "task=" & taskName
End Function

Public Function ResolveWatchedFolders(ByVal session As Outlook.NameSpace) As Collection
    On Error GoTo CleanFail
    Dim result As Collection
    Set result = New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim inbox As Outlook.MAPIFolder
    Set inbox = session.GetDefaultFolder(olFolderInbox)
    If Not inbox Is Nothing Then
        result.Add inbox
        seen.Add inbox.FolderPath, True
    End If

    Dim specs As Variant
    Dim spec As Variant
    If Len(Trim$(ADDITIONAL_FOLDERS)) > 0 Then
        specs = Split(ADDITIONAL_FOLDERS, "|")
        For Each spec In specs
            Dim folder As Outlook.MAPIFolder
            Set folder = ResolveFolderPath(session, CStr(spec))
            If Not folder Is Nothing Then
                If Not seen.Exists(folder.FolderPath) Then
                    result.Add folder
                    seen.Add folder.FolderPath, True
                End If
            End If
        Next spec
    End If

    Set ResolveWatchedFolders = result
    Exit Function
CleanFail:
    modHVDC_Logging.LogException "ResolveWatchedFolders", Err, ""
End Function

Public Function ResolveFolderPath(ByVal session As Outlook.NameSpace, ByVal relativePath As String) As Outlook.MAPIFolder
    On Error GoTo CleanFail
    Dim cleanPath As String
    cleanPath = Trim$(relativePath)
    If Len(cleanPath) = 0 Then
        Exit Function
    End If

    Dim segments As Variant
    segments = Split(cleanPath, "\\")

    Dim currentFolder As Outlook.MAPIFolder
    Set currentFolder = session.GetDefaultFolder(olFolderInbox)
    Dim idx As Long
    For idx = LBound(segments) To UBound(segments)
        Dim nameSegment As String
        nameSegment = Trim$(segments(idx))
        If Len(nameSegment) = 0 Then
            GoTo CleanFail
        End If
        If idx = LBound(segments) Then
            If StrComp(nameSegment, currentFolder.Name, vbTextCompare) = 0 Then
                GoTo ContinueLoop
            End If
        End If
        If currentFolder Is Nothing Then Exit Function
        If currentFolder.Folders.Count = 0 Then Exit Function
        On Error Resume Next
        Set currentFolder = currentFolder.Folders(nameSegment)
        If Err.Number <> 0 Then
            On Error GoTo CleanFail
            Exit Function
        End If
        On Error GoTo CleanFail
        If currentFolder Is Nothing Then
            Exit Function
        End If
ContinueLoop:
    Next idx
    Set ResolveFolderPath = currentFolder
    Exit Function
CleanFail:
    On Error Resume Next
    modHVDC_Logging.LogException "ResolveFolderPath", Err, "path=" & relativePath
End Function

Public Sub RefreshQueuePulse()
    RequestQueuePulse
End Sub

