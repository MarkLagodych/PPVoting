Attribute VB_Name = "Logging"
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'                          ОТЧЁТЫ ДЛЯ PPVOTING                        '
' ------------------------------------------------------------------- '
'                                                                     '
' Задачи этого модуля:                                                '
'   - записывать различные ошибки и результаты опросов в файл. Путь к '
'     файлу указан также в настройках PPVoting                        '
' Формат:                                                             '
'   Ошибка:     [дата,время] ERROR (процедура) #номерОшибки: описание '
'   Информация: [дата,время] INFO (процедура|голоса): ...             '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

' Флаг, нужный для безошибочного завершения
' И выборочной работы (в зависимости от успешности начала)
' По умолчанию false,
' После успешного начала true,
' В конце false
Private started As Boolean

Private logFile As TextStream

Public Sub start()
    On Error GoTo Error_label
    
    If Not Base.settings("logging").Exists("file") Then
        Err.Raise 1, Description:="logging.file does not exist"
    End If
    
    Set logFile = Base.FS.OpenTextFile(Base.settings("logging")("file"), ForAppending, True, TristateTrue)
    
    started = True
    
    logLine
    logInfo "Logging.start", "Logging turned on"
    
    Exit Sub
Error_label:
    Base.ShowRawError "Logging.start"
    
End Sub

' Если Logging не используется (в настройках), ничего не делает
Public Sub logError(proc As String, number As Integer, msg As String)
    On Error GoTo Error_label
    
    If Not started Then Exit Sub
    
    logFile.WriteLine "[" & Now() & "] ERROR (" & proc & ") #" & number & ": " & msg
    
    Exit Sub
Error_label:
    Base.ShowRawError "Logging.logError"
End Sub

Public Sub logVotes(ByRef votes As Dictionary)
    On Error GoTo Error_label
    
    If Not started Then Exit Sub
    
    logFile.Write "[" & Now() & "] INFO (votes): "
    
    Dim i As Integer
    For i = 1 To 7
        logFile.Write CStr(votes("total")(i))
        If i <> 7 Then logFile.Write ","
        logFile.Write " "
    Next i
    
    logFile.WriteLine
    
    If Not votes.Exists("votes") Then
        logInfo "Logging.logVotes", "Voting information is not detailed (no ID--vote info)"
        Exit Sub
    End If
    
    Dim vote As Variant
    For Each vote In votes("votes")
        ' Полученные голоса в промежутке [0;6], а нужно [1;7]
        logFile.WriteLine "ID " & vote(1) & " voted for " & vote(2) + 1
    Next vote
    
    Exit Sub
Error_label:
    Base.ShowRawError "Logging.logVotes"
End Sub

Public Sub logInfo(proc As String, msg As String)
    On Error GoTo Error_label
    
    If Not started Then Exit Sub
    
    logFile.WriteLine "[" & Now() & "] INFO (" & proc & "): " & msg
    
    Exit Sub
Error_label:
    Base.ShowRawError "Logging.logInfo"
End Sub

Public Sub logLine()
    On Error GoTo Error_label
    
    If Not started Then Exit Sub
    
    logFile.WriteLine String(16, "-")
    
    Exit Sub
Error_label:
    Base.ShowRawError "Logging.logLine"
End Sub

Public Sub finish()
    On Error GoTo Error_label
    
    logInfo "Logging.finish", "Logging shut down"
    
    If started Then
        logFile.Close
    End If
    started = False
    
    Exit Sub
Error_label:
    started = False
    Base.ShowRawError "Logging.finish"
End Sub
