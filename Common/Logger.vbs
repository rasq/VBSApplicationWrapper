'******************************************************************************
' Logger.vbs        Script for logging
'
' Description     : Logging feature used by both common script and package script
'                   
'******************************************************************************
Option Explicit

'===================================================================
' (Global variable) Default logging level
'===================================================================
Dim g_defaultLogLevel
g_defaultLogLevel = LOG_LEVEL_INFO

'===================================================================
' (Global variable) Logger object of currenly processing script
'===================================================================
Dim g_currentLogger

'===================================================================
' (Global variable) Logger object managemnet table
'===================================================================
Dim g_loggers
Set g_loggers = CreateObject("Scripting.Dictionary")

Class Logger

    Private m_logLevel
    Private m_appendLog
    Private m_logFilePath
    Private m_logFile

    Private Sub Class_Initialize
        m_logLevel = LOG_LEVEL_INFO
        m_appendLog = False
        m_logFilePath = ""
        Set m_logFile = Nothing
    End Sub

    Private Sub Class_Terminate
        CloseLogFile
    End Sub

    Property Get LogLevel
        LogLevel = m_logLevel
    End Property

    Property Let LogLevel(newLogLevel)
        m_logLevel = newLogLevel
    End Property

    Property Get AppendLog
        AppendLog = m_appendLog
    End Property

    Property Let AppendLog(newAppendLog)
        m_appendLog = newAppendLog
    End Property

    Property Get LogFilePath
        LogFilePath = m_logFilePath
    End Property

    Property Let LogFilePath(newLogFilePath)
        m_logFilePath = newLogFilePath
        Update
    End Property

    Private Sub Update
        CloseLogFile
        OpenLogFile
    End Sub

    Sub CloseLogFile

        If m_logFile Is Nothing Or IsNull(m_logFile) Then
            Exit Sub
        End If

        m_logFile.Close

        Set m_logFile = Nothing

    End Sub

    Sub OpenLogFile

        Dim strParentFolder
        strParentFolder = objFS.GetParentFolderName(objFS.GetAbsolutePathName(m_logFilePath))

        If Not objFS.FolderExists(strParentFolder) Then

            CreateFolder2(strParentFolder)

        End If

        If objFS.FileExists(m_logFilePath) And (m_appendLog = False) Then
            objFS.DeleteFile m_logFilePath, True
        End If

        Set m_logFile = objFS.OpenTextFile(m_logFilePath, 8, True)

    End Sub

    Sub WriteLog(message)
'       m_logFile.WriteLine Now & " " & message
        Dim tmNow
        tmNow = Now()
        m_logFile.WriteLine Year(tmNow) & "/" & _
                            Right( "0" & Month (tmNow), 2 ) & "/" & _
                            Right( "0" & Day   (tmNow), 2 ) & " " & _
                            Right( "0" & Hour  (tmNow), 2 ) & ":" & _
                            Right( "0" & Minute(tmNow), 2 ) & ":" & _
                            Right( "0" & Second(tmNow), 2 ) & " " & _
                            message
    End Sub

    Sub Debug(message)
        If IsLogLevelDebug(m_logLevel) Then
            WriteLog LOG_LEVEL_DEBUG_PREFIX & message
        End If
    End Sub

    Sub Info(message)
        If IsLogLevelInfo(m_logLevel) Then
            WriteLog LOG_LEVEL_INFO_PREFIX & message
        End If
    End Sub

    Sub Warn(message)
        If IsLogLevelWarn(m_logLevel) Then
            WriteLog LOG_LEVEL_WARN_PREFIX & message
        End If
    End Sub

    Sub Error(message)
        If IsLogLevelError(m_logLevel) Then
            WriteLog LOG_LEVEL_ERROR_PREFIX & message
        End If
    End Sub

    Sub Fatal(message)
        If IsLogLevelFatal(m_logLevel) Then
            WriteLog LOG_LEVEL_FATAL_PREFIX & message
        End If
    End Sub

    Sub PopupMessage(message, second)
        Dim mes
        If second <> 0 Then
            mes = message & vbCr & vbCr & "(This message will be closed in " & second & " seconds automatically.)"
        Else
            mes = message
        End If
        objWshShell.Popup mes, second, "Message", vbOkOnly

    End Sub

    Sub PopupLog(second)
        Dim buffer, mes
        m_logFile.Close
        Set m_logFile = objFS.OpenTextFile(m_logFilePath, ForReading, False)
        buffer = m_logFile.ReadAll
        If second <> 0 Then
            mes = buffer & vbCr & vbCr & "(This message will be closed in " & second & " seconds automatically.)"
        Else
            mes = buffer
        End If
        objWshShell.Popup mes, second, "Log", vbOkOnly
        m_logFile.Close
        Set m_logFile = objFS.OpenTextFile(m_logFilePath, ForAppending, True)
    End sub

    Function IsLogLevelDebug(logLevel)
        If (GetLogLevelPriority(logLevel) <= LOG_LEVEL_DEBUG_PRIORITY) Then
            IsLogLevelDebug = True
        Else
            IsLogLevelDebug = False
        End If
    End Function

    Function IsLogLevelInfo(logLevel)
        If (GetLogLevelPriority(logLevel) <= LOG_LEVEL_INFO_PRIORITY) Then
            IsLogLevelInfo = True
        Else
            IsLogLevelInfo = False
        End If
    End Function

    Function IsLogLevelWarn(logLevel)
        If (GetLogLevelPriority(logLevel) <= LOG_LEVEL_WARN_PRIORITY) Then
            IsLogLevelWarn = True
        Else
            IsLogLevelWarn = False
        End If
    End Function

    Function IsLogLevelError(logLevel)
        If (GetLogLevelPriority(logLevel) <= LOG_LEVEL_ERROR_PRIORITY) Then
            IsLogLevelError = True
        Else
            IsLogLevelError = False
        End If
    End Function

    Function IsLogLevelFatal(logLevel)
        If (GetLogLevelPriority(logLevel) <= LOG_LEVEL_FATAL_PRIORITY) Then
            IsLogLevelFatal = True
        Else
            IsLogLevelFatal = False
        End If
    End Function

    Function GetLogLevelPriority(logLevelName)
        Dim intLogLevel
        Select Case UCase(logLevelName)
        Case LOG_LEVEL_DEBUG            
            intLogLevel = LOG_LEVEL_DEBUG_PRIORITY
        Case LOG_LEVEL_INFO
            intLogLevel = LOG_LEVEL_INFO_PRIORITY
        Case LOG_LEVEL_WARN
            intLogLevel = LOG_LEVEL_WARN_PRIORITY
        Case LOG_LEVEL_ERROR
            intLogLevel = LOG_LEVEL_ERROR_PRIORITY
        Case LOG_LEVEL_FATAL
            intLogLevel = LOG_LEVEL_FATAL_PRIORITY
        Case Else
            intLogLevel = -1
        End Select
        GetLogLevelPriority = intLogLevel
    End Function

End Class

'*************************************************************************
' Function Id  : GetLogger
'
' Parameter    : logFilePath : Log file path
'
' Return Value : Logger object corresponding to Log file path
'
' Description  : Return Logger object corresponding to Log file path
'
'                If you call this function with same log file path twice or more,
'                same object as first time will be returned
'                as Logger object is cached in Global variable with key of log file path
'                
'
'*************************************************************************
Function GetLogger(logFilePath)
    Dim objTargetLogger
    If g_loggers.Exists(logFilePath) Then
        Set objTargetLogger = g_loggers(logFilePath)
    Else
        Set objTargetLogger = New Logger
        objTargetLogger.LogLevel = g_defaultLogLevel
        objTargetLogger.LogFilePath = logFilePath
        Set g_loggers(logFilePath) = objTargetLogger
    End If
    Set GetLogger = objTargetLogger
    Set objTargetLogger = Nothing
End Function

'*************************************************************************
' Function Id  : GetCurrentLogger
'
' Parameter    : |
'
' Return Value : Logger object of currenly processing script
'
' Description  : Return Logger object of currenly processing script
'*************************************************************************
Function GetCurrentLogger
    Set GetCurrentLogger = g_currentLogger
End Function

'*************************************************************************
' Function Id  : SetCurrentLogger
'
' Parameter    : logger Logger object you want to set as current Logger
'
' Return Value : N/A
'
' Description  : Set Logger object specified in parameter to current Logger object
'                
'*************************************************************************
Function SetCurrentLogger(logger)
    Set g_currentLogger = logger
End Function
