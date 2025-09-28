# Outlook VBA 이메일 캡처 가이드 | Outlook VBA Email Capture Guide

## 📧 개요 | Overview

Outlook VBA를 사용하여 `Application_NewMailEx` 이벤트를 통해 실시간으로 이메일을 캡처하고 HVDC 이메일 파이프라인으로 전송하는 방법을 설명합니다.

## 🔧 설정 방법 | Setup Instructions

### 1. Outlook VBA 개발자 모드 활성화

1. **Outlook 열기** → **파일** → **옵션**
2. **보안 센터** → **보안 센터 설정**
3. **매크로 설정** → **모든 매크로 사용** 선택
4. **개발자 탭 표시** 체크

### 2. VBA 편집기 열기

1. **개발자** 탭 → **Visual Basic** 클릭
2. 또는 **Alt + F11** 단축키 사용

### 3. ThisOutlookSession 모듈에 코드 추가

```vba
Option Explicit

' 전역 변수
Private WithEvents olApp As Outlook.Application
Private Const CAPTURE_FOLDER As String = "C:\HVDC\EmailCapture\"
Private Const ATTACHMENT_FOLDER As String = "C:\HVDC\EmailCapture\Attachments\"

' Outlook 시작 시 이벤트 핸들러 등록
Private Sub Application_Startup()
    Set olApp = Application
    CreateCaptureFolders
End Sub

' 새 이메일 수신 시 이벤트 처리
Private Sub olApp_NewMailEx(EntryIDCollection As String)
    Dim arrEntryIDs As Variant
    Dim i As Integer
    Dim olItem As Object
    Dim olMail As MailItem
    
    ' EntryID를 배열로 분할
    arrEntryIDs = Split(EntryIDCollection, ",")
    
    ' 각 이메일 처리
    For i = 0 To UBound(arrEntryIDs)
        Set olItem = Application.Session.GetItemFromID(Trim(arrEntryIDs(i)))
        
        If olItem.Class = olMail Then
            Set olMail = olItem
            ProcessEmail olMail
        End If
    Next i
End Sub

' 이메일 처리 함수
Private Sub ProcessEmail(olMail As MailItem)
    Dim fileName As String
    Dim filePath As String
    Dim entryID As String
    Dim captureTime As String
    
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
    
    ' Python 파이프라인으로 전송 (선택사항)
    SendToPipeline filePath, entryID
End Sub

' 첨부 파일 처리 함수
Private Sub ProcessAttachments(olMail As MailItem, entryID As String)
    Dim att As Attachment
    Dim attPath As String
    Dim attFolder As String
    
    ' 첨부 파일 폴더 생성
    attFolder = ATTACHMENT_FOLDER & Left(entryID, 8) & "\"
    CreateFolderIfNotExists attFolder
    
    ' 각 첨부 파일 저장
    For Each att In olMail.Attachments
        attPath = attFolder & att.FileName
        att.SaveAsFile attPath
    Next att
End Sub

' 폴더 생성 함수
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

' 로그 기록 함수
Private Sub LogEmailCapture(olMail As MailItem, filePath As String)
    Dim logFile As String
    Dim logText As String
    Dim fso As Object
    Dim logStream As Object
    
    logFile = CAPTURE_FOLDER & "capture_log.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logText = Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & _
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
End Sub

' Python 파이프라인으로 전송 (선택사항)
Private Sub SendToPipeline(filePath As String, entryID As String)
    Dim pythonScript As String
    Dim command As String
    
    ' Python 스크립트 경로
    pythonScript = "C:\HVDC\email_processing_pipeline\run_cli.py"
    
    ' 명령어 구성
    command = "python """ & pythonScript & """ ingest """ & filePath & """ --entry-id """ & entryID & """"
    
    ' 백그라운드에서 실행
    Call Shell(command, vbHide)
End Sub
```

## 🔄 대안: Items.ItemAdd 이벤트

Outlook 규칙으로 이메일이 이동된 후에도 캡처하려면:

```vba
' 특정 폴더 모니터링
Private WithEvents olInbox As Outlook.Folder
Private WithEvents olInboxItems As Outlook.Items

Private Sub Application_Startup()
    Set olApp = Application
    Set olInbox = olApp.Session.GetDefaultFolder(olFolderInbox)
    Set olInboxItems = olInbox.Items
End Sub

' 폴더에 새 아이템 추가 시 이벤트
Private Sub olInboxItems_ItemAdd(ByVal Item As Object)
    If Item.Class = olMail Then
        ProcessEmail Item
    End If
End Sub
```

## ⚙️ 고급 설정 | Advanced Configuration

### 1. 특정 폴더만 모니터링

```vba
' 특정 폴더 설정
Private Sub SetMonitorFolder(folderName As String)
    Dim olFolder As Outlook.Folder
    Set olFolder = GetFolderByName(folderName)
    Set olInboxItems = olFolder.Items
End Sub

Private Function GetFolderByName(folderName As String) As Outlook.Folder
    Dim olFolder As Outlook.Folder
    Dim i As Integer
    
    Set olFolder = olApp.Session.GetDefaultFolder(olFolderInbox)
    
    For i = 1 To olFolder.Folders.Count
        If olFolder.Folders(i).Name = folderName Then
            Set GetFolderByName = olFolder.Folders(i)
            Exit Function
        End If
    Next i
End Function
```

### 2. 필터링 조건 추가

```vba
Private Function ShouldCapture(olMail As MailItem) As Boolean
    ' 특정 발신자만 캡처
    If InStr(olMail.SenderEmailAddress, "@hvdc.local") = 0 Then
        ShouldCapture = False
        Exit Function
    End If
    
    ' 특정 키워드가 포함된 제목만 캡처
    If InStr(LCase(olMail.Subject), "logistics") = 0 And _
       InStr(LCase(olMail.Subject), "hvdc") = 0 Then
        ShouldCapture = False
        Exit Function
    End If
    
    ShouldCapture = True
End Function
```

### 3. 에러 처리 강화

```vba
Private Sub ProcessEmail(olMail As MailItem)
    On Error GoTo ErrorHandler
    
    ' 이메일 처리 로직
    ' ...
    
    Exit Sub
    
ErrorHandler:
    LogError "ProcessEmail", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub LogError(functionName As String, errorNumber As Long, errorDescription As String)
    Dim logFile As String
    Dim logText As String
    Dim fso As Object
    Dim logStream As Object
    
    logFile = CAPTURE_FOLDER & "error_log.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logText = Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & _
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
```

## 📁 폴더 구조 | Folder Structure

```
C:\HVDC\EmailCapture\
├── 20250126_143022_ABC12345.msg
├── 20250126_143045_DEF67890.msg
├── capture_log.txt
├── error_log.txt
└── Attachments\
    ├── ABC12345\
    │   ├── report.pdf
    │   └── data.xlsx
    └── DEF67890\
        └── invoice.pdf
```

## 🔧 설정 파일 | Configuration File

`config.vba` 파일 생성:

```vba
' HVDC 이메일 캡처 설정
Public Const CAPTURE_FOLDER As String = "C:\HVDC\EmailCapture\"
Public Const ATTACHMENT_FOLDER As String = "C:\HVDC\EmailCapture\Attachments\"
Public Const PYTHON_SCRIPT As String = "C:\HVDC\email_processing_pipeline\run_cli.py"
Public Const LOG_LEVEL As String = "INFO" ' DEBUG, INFO, WARN, ERROR

' 필터링 설정
Public Const SENDER_FILTER As String = "@hvdc.local"
Public Const SUBJECT_KEYWORDS As String = "logistics,hvdc,status,report"
Public Const MAX_FILE_SIZE As Long = 10485760 ' 10MB
```

## 🚀 자동화 설정 | Automation Setup

### 1. Outlook 시작 시 자동 실행

1. **파일** → **옵션** → **고급**
2. **개발자용** → **COM 추가 기능** 체크
3. VBA 매크로를 COM 추가 기능으로 등록

### 2. Windows 작업 스케줄러 연동

```batch
@echo off
REM HVDC 이메일 캡처 모니터링
cd /d "C:\HVDC\email_processing_pipeline"
python run_cli.py status
```

## 📊 모니터링 | Monitoring

### 1. 실시간 로그 확인

```vba
Private Sub ShowCaptureStatus()
    Dim logFile As String
    Dim fso As Object
    Dim logStream As Object
    Dim logContent As String
    
    logFile = CAPTURE_FOLDER & "capture_log.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(logFile) Then
        Set logStream = fso.OpenTextFile(logFile, 1, False)
        logContent = logStream.ReadAll
        logStream.Close
        
        ' 로그 내용을 메시지박스로 표시
        MsgBox logContent, vbInformation, "HVDC 이메일 캡처 상태"
    End If
End Sub
```

### 2. 성능 지표 수집

```vba
Private Sub CollectMetrics()
    Dim metricsFile As String
    Dim fso As Object
    Dim metricsStream As Object
    Dim metrics As String
    
    metricsFile = CAPTURE_FOLDER & "metrics.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    metrics = "Capture Rate: " & GetCaptureRate() & "%" & vbCrLf & _
              "Total Emails: " & GetTotalEmails() & vbCrLf & _
              "Last Capture: " & GetLastCaptureTime() & vbCrLf
    
    Set metricsStream = fso.CreateTextFile(metricsFile, True)
    metricsStream.Write metrics
    metricsStream.Close
End Sub
```

## ⚠️ 주의사항 | Important Notes

1. **보안**: VBA 매크로는 보안 위험이 있을 수 있으므로 신뢰할 수 있는 소스에서만 실행
2. **성능**: 대량의 이메일 처리 시 Outlook 성능에 영향을 줄 수 있음
3. **에러 처리**: 네트워크 오류나 파일 시스템 오류에 대한 적절한 처리 필요
4. **백업**: 중요한 이메일은 별도 백업 시스템 구축 권장

## 🔧 문제 해결 | Troubleshooting

### 1. 이벤트가 발생하지 않는 경우
- Outlook 재시작
- VBA 코드 재컴파일 (F5)
- 매크로 보안 설정 확인

### 2. 파일 저장 실패
- 폴더 권한 확인
- 디스크 공간 확인
- 파일명에 특수문자 포함 여부 확인

### 3. Python 연동 실패
- Python 경로 확인
- 가상환경 활성화 확인
- 의존성 설치 확인

---

이 VBA 코드를 사용하면 Outlook에서 실시간으로 이메일을 캡처하고 HVDC 이메일 파이프라인으로 자동 전송할 수 있습니다.
