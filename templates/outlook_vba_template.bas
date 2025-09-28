Attribute VB_Name = "HVDC_EmailCapture"
Option Explicit

' =============================================================================
' HVDC 이메일 캡처 VBA 모듈
' Outlook VBA를 사용한 실시간 이메일 캡처 시스템
' =============================================================================

' 전역 변수
Private WithEvents olApp As Outlook.Application
Private Const CAPTURE_FOLDER As String = "C:\HVDC\EmailCapture\"
Private Const ATTACHMENT_FOLDER As String = "C:\HVDC\EmailCapture\Attachments\"
Private Const PYTHON_SCRIPT As String = "C:\HVDC\email_processing_pipeline\run_cli.py"

' =============================================================================
' Outlook 시작 시 이벤트 핸들러 등록
' =============================================================================
Private Sub Application_Startup()
    Set olApp = Application
    CreateCaptureFolders
    LogMessage "HVDC 이메일 캡처 시스템이 시작되었습니다."
End Sub

' =============================================================================
' 새 이메일 수신 시 이벤트 처리 (실시간 캡처)
' =============================================================================
Private Sub olApp_NewMailEx(EntryIDCollection As String)
    Dim arrEntryIDs As Variant
    Dim i As Integer
    Dim olItem As Object
    Dim olMail As MailItem
    
    On Error GoTo ErrorHandler
    
    ' EntryID를 배열로 분할
    arrEntryIDs = Split(EntryIDCollection, ",")
    
    ' 각 이메일 처리
    For i = 0 To UBound(arrEntryIDs)
        Set olItem = Application.Session.GetItemFromID(Trim(arrEntryIDs(i)))
        
        If olItem.Class = olMail Then
            Set olMail = olItem
            If ShouldCapture(olMail) Then
                ProcessEmail olMail
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    LogError "NewMailEx", Err.Number, Err.Description
    Resume Next
End Sub

' =============================================================================
' 이메일 처리 함수
' =============================================================================
Private Sub ProcessEmail(olMail As MailItem)
    Dim fileName As String
    Dim filePath As String
    Dim entryID As String
    Dim captureTime As String
    
    On Error GoTo ErrorHandler
    
    ' 파일명 생성 (EntryID + 타임스탬프)
    entryID = olMail.EntryID
    captureTime = Format(Now, "yyyymmdd_hhmmss")
    fileName = "email_" & captureTime & "_" & Left(entryID, 8) & ".msg"
    filePath = CAPTURE_FOLDER & fileName
    
    ' 이메일을 .msg 파일로 저장
    olMail.SaveAs filePath, olMSG
    
    ' 첨부 파일 처리
    If olMail.Attachments.Count > 0 Then
        ProcessAttachments olMail, entryID
    End If
    
    ' 로그 기록
    LogEmailCapture olMail, filePath
    
    ' Python 파이프라인으로 전송
    SendToPipeline filePath, entryID
    
    Exit Sub
    
ErrorHandler:
    LogError "ProcessEmail", Err.Number, Err.Description
    Resume Next
End Sub

' =============================================================================
' 첨부 파일 처리 함수
' =============================================================================
Private Sub ProcessAttachments(olMail As MailItem, entryID As String)
    Dim att As Attachment
    Dim attPath As String
    Dim attFolder As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' 첨부 파일 폴더 생성
    attFolder = ATTACHMENT_FOLDER & Left(entryID, 8) & "\"
    CreateFolderIfNotExists attFolder
    
    ' 각 첨부 파일 저장
    For i = 1 To olMail.Attachments.Count
        Set att = olMail.Attachments(i)
        attPath = attFolder & CleanFileName(att.FileName)
        att.SaveAsFile attPath
    Next i
    
    Exit Sub
    
ErrorHandler:
    LogError "ProcessAttachments", Err.Number, Err.Description
    Resume Next
End Sub

' =============================================================================
' 이메일 캡처 여부 판단 함수
' =============================================================================
Private Function ShouldCapture(olMail As MailItem) As Boolean
    ' 기본적으로 모든 이메일 캡처
    ShouldCapture = True
    
    ' 특정 조건으로 필터링 가능
    ' 예: 특정 발신자만 캡처
    ' If InStr(olMail.SenderEmailAddress, "@hvdc.local") = 0 Then
    '     ShouldCapture = False
    '     Exit Function
    ' End If
    
    ' 예: 특정 키워드가 포함된 제목만 캡처
    ' If InStr(LCase(olMail.Subject), "logistics") = 0 And _
    '    InStr(LCase(olMail.Subject), "hvdc") = 0 Then
    '     ShouldCapture = False
    '     Exit Function
    ' End If
End Function

' =============================================================================
' 폴더 생성 함수들
' =============================================================================
Private Sub CreateCaptureFolders()
    CreateFolderIfNotExists CAPTURE_FOLDER
    CreateFolderIfNotExists ATTACHMENT_FOLDER
End Sub

Private Sub CreateFolderIfNotExists(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

' =============================================================================
' 로깅 함수들
' =============================================================================
Private Sub LogEmailCapture(olMail As MailItem, filePath As String)
    Dim logFile As String
    Dim logText As String
    Dim fso As Object
    Dim logStream As Object
    
    On Error GoTo ErrorHandler
    
    logFile = CAPTURE_FOLDER & "capture_log.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logText = Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & _
              "CAPTURED" & " | " & _
              olMail.Subject & " | " & _
              olMail.SenderName & " | " & _
              filePath & vbCrLf
    
    If fso.FileExists(logFile) Then
        Set logStream = fso.OpenTextFile(logFile, 8, True) ' 8 = ForAppending
    Else
        Set logStream = fso.CreateTextFile(logFile, True)
    End If
    
    logStream.Write logText
    logStream.Close
    
    Exit Sub
    
ErrorHandler:
    LogError "LogEmailCapture", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub LogMessage(message As String)
    Dim logFile As String
    Dim logText As String
    Dim fso As Object
    Dim logStream As Object
    
    On Error GoTo ErrorHandler
    
    logFile = CAPTURE_FOLDER & "system_log.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logText = Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & message & vbCrLf
    
    If fso.FileExists(logFile) Then
        Set logStream = fso.OpenTextFile(logFile, 8, True)
    Else
        Set logStream = fso.CreateTextFile(logFile, True)
    End If
    
    logStream.Write logText
    logStream.Close
    
    Exit Sub
    
ErrorHandler:
    ' 에러 로깅 실패 시 무시
    Resume Next
End Sub

Private Sub LogError(functionName As String, errorNumber As Long, errorDescription As String)
    Dim logFile As String
    Dim logText As String
    Dim fso As Object
    Dim logStream As Object
    
    On Error Resume Next
    
    logFile = CAPTURE_FOLDER & "error_log.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logText = Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & _
              "ERROR" & " | " & _
              functionName & " | " & _
              "Error " & errorNumber & ": " & errorDescription & vbCrLf
    
    If fso.FileExists(logFile) Then
        Set logStream = fso.OpenTextFile(logFile, 8, True)
    Else
        Set logStream = fso.CreateTextFile(logFile, True)
    End If
    
    logStream.Write logText
    logStream.Close
End Sub

' =============================================================================
' Python 파이프라인 연동 함수
' =============================================================================
Private Sub SendToPipeline(filePath As String, entryID As String)
    Dim command As String
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    ' Python 스크립트 존재 확인
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(PYTHON_SCRIPT) Then
        LogMessage "Python 스크립트를 찾을 수 없습니다: " & PYTHON_SCRIPT
        Exit Sub
    End If
    
    ' 명령어 구성
    command = "python """ & PYTHON_SCRIPT & """ ingest """ & filePath & """ --entry-id """ & entryID & """"
    
    ' 백그라운드에서 실행
    Call Shell(command, vbHide)
    
    LogMessage "Python 파이프라인으로 전송: " & filePath
    
    Exit Sub
    
ErrorHandler:
    LogError "SendToPipeline", Err.Number, Err.Description
    Resume Next
End Sub

' =============================================================================
' 유틸리티 함수들
' =============================================================================
Private Function CleanFileName(fileName As String) As String
    Dim cleanName As String
    Dim i As Integer
    Dim char As String
    
    cleanName = ""
    For i = 1 To Len(fileName)
        char = Mid(fileName, i, 1)
        If char >= " " And char <= "~" And char <> "<" And char <> ">" And char <> ":" And char <> """ And char <> "/" And char <> "\" And char <> "|" And char <> "?" And char <> "*" Then
            cleanName = cleanName & char
        Else
            cleanName = cleanName & "_"
        End If
    Next i
    
    CleanFileName = cleanName
End Function

' =============================================================================
' 수동 실행 함수들 (디버깅 및 테스트용)
' =============================================================================
Public Sub TestEmailCapture()
    ' 테스트용 이메일 캡처 함수
    Dim olMail As MailItem
    Set olMail = Application.ActiveExplorer.Selection(1)
    
    If Not olMail Is Nothing Then
        ProcessEmail olMail
        MsgBox "테스트 이메일이 캡처되었습니다.", vbInformation
    Else
        MsgBox "이메일을 선택해주세요.", vbExclamation
    End If
End Sub

Public Sub ShowCaptureStatus()
    ' 캡처 상태 확인 함수
    Dim logFile As String
    Dim fso As Object
    Dim logStream As Object
    Dim logContent As String
    
    On Error GoTo ErrorHandler
    
    logFile = CAPTURE_FOLDER & "capture_log.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(logFile) Then
        Set logStream = fso.OpenTextFile(logFile, 1, False)
        logContent = logStream.ReadAll
        logStream.Close
        
        ' 최근 10줄만 표시
        Dim lines As Variant
        lines = Split(logContent, vbCrLf)
        Dim recentLines As String
        Dim i As Integer
        Dim startLine As Integer
        
        startLine = Application.Max(0, UBound(lines) - 9)
        For i = startLine To UBound(lines)
            If Trim(lines(i)) <> "" Then
                recentLines = recentLines & lines(i) & vbCrLf
            End If
        Next i
        
        MsgBox "최근 캡처 이력:" & vbCrLf & vbCrLf & recentLines, vbInformation, "HVDC 이메일 캡처 상태"
    Else
        MsgBox "캡처 로그 파일을 찾을 수 없습니다.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "상태 확인 중 오류가 발생했습니다: " & Err.Description, vbCritical
End Sub

Public Sub ClearLogs()
    ' 로그 파일 정리 함수
    Dim fso As Object
    Dim logFiles As Variant
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFiles = Array("capture_log.txt", "error_log.txt", "system_log.txt")
    
    For i = 0 To UBound(logFiles)
        If fso.FileExists(CAPTURE_FOLDER & logFiles(i)) Then
            fso.DeleteFile CAPTURE_FOLDER & logFiles(i)
        End If
    Next i
    
    MsgBox "로그 파일이 정리되었습니다.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "로그 정리 중 오류가 발생했습니다: " & Err.Description, vbCritical
End Sub
