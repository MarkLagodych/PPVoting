Attribute VB_Name = "Voting"
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'                      ГОЛОСОВАНИЕ ДЛЯ PPVOTING                       '
' ------------------------------------------------------------------- '
'                                                                     '
' Задачи этого модуля:                                                '
'   - подключиться к серверу(ESP8266) через COM-порт;                 '
'   - проверять кнопки показа диаграммы                               '
'   - по команде "показать", отослать на сервер 'g', принять ответ и  '
'     отобоазить его на диаграмме;                                    '
'   - по команде "стереть", отослать на сервер 'с'.                   '
'   - вне показа слайдов предоставлять проверку подключения к серверу '
'     (см. ui.xml)                                                    '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit


' Флаг, нужный для безошибочного завершения
' По умолчанию false,
' После успешного начала true,
' В конце false
Private started As Boolean

Dim COMPort As Integer          ' Номер порта, к которому подключится ESP8266

Dim timerID                     ' Системный ID таймера

Dim lastSlide As Integer        ' Номер последнего показаного слайда

Public votes As Dictionary      ' Сюда запишутся результаты голосования

Dim lngStatus As Long           ' Код результата, нужный для библиотеки modComm (COMPortAPI в проекте)
Dim strError As String          ' Для библиотеки modComm (COMPortAPI в проекте)
Dim strData As String


Public Sub start()
    On Error GoTo Error_label
    
    COMPort = Base.settings("voting")("port")
    
    lngStatus = CommOpen(COMPort, "COM" & COMPort, "baud=115200 parity=N data=8 stop=1")
    If lngStatus <> 0 Then
        lngStatus = CommGetError(strError)
        Err.Raise lngStatus, Description:=strError
    End If
    
    CommSetLine COMPort, LINE_RTS, True
    CommSetLine COMPort, LINE_DTR, True

    timerID = Base.SetTimer(0, 0, 100, AddressOf Run) ' Повторять каждые 100 мс
    
    started = True
    
    logging.logInfo "Voting.start", "Communication started on COM port " & COMPort & ", frequency 115200 bauds, 8N1 configuration"
    
    Exit Sub
Error_label:
    Base.ShowError "Voting.start"
    
End Sub


Sub Run()
    On Error GoTo Error_label


    ' Проверить управляющие кнопки и выполнить действия при нажатии
    
    Dim key As Variant
    
    ' Показать/скрыть диаграмму
    For Each key In Base.diagramShowKeys
    
        If Base.CheckKeyPressed(key) Then
            With diagram
                If .Visible Then
                    .Hide
                    
                    ' Нужно чтобы убрать белый/чёрный экран
                    ActivePresentation.SlideShowWindow.View.GotoSlide lastSlide
                Else
                    lastSlide = ActivePresentation.SlideShowWindow.View.slide.slideIndex
                    loadVotes
                    .Show
                End If
            End With
        End If
        
    Next key
    
    
    If diagram.Visible Then
        
        For Each key In Base.backwardKeys
            If Base.CheckKeyPressed(key) Then
                clearVotes
                diagram.Hide
                ActivePresentation.SlideShowWindow.View.GotoSlide lastSlide
            End If
        Next key
        
        For Each key In Base.forwardKeys
            If Base.CheckKeyPressed(key) Then loadVotes
        Next key
        
    End If
    
    
    Exit Sub
Error_label:
    Base.ShowError "Voting.run"
    
End Sub


Public Sub finish()
    On Error GoTo Error_label
    
    ' Безопасное завершение
    If started Then
        Base.KillTimer 0, timerID
        
        CommSetLine COMPort, LINE_RTS, False
        CommSetLine COMPort, LINE_DTR, False
        CommClose COMPort
        
        logging.logInfo "Voting.finish", "COM communication shut down"
    End If
    
    started = False
    
    
    Exit Sub
Error_label:
    started = False
    Base.ShowError "Voting.finish"
    
End Sub


Sub SendMessage(msg As String)
    On Error GoTo Error_label
    
    Dim lngSize As Long
    lngSize = Len(msg)
    lngStatus = CommWrite(COMPort, msg)
    
    If lngStatus <> lngSize Then
        lngStatus = CommGetError(strError)
        Err.Raise lngStatus, Description:=strError
    End If
    
    Exit Sub
Error_label:
    Base.ShowError "Voting.SendMessage"
End Sub


Public Sub clearVotes()
    On Error GoTo Error_label
    
    diagram.clearVotes
    SendMessage "c"
    logging.logInfo "Voting.clearVotes", "All votes cleared from the server"
    
    Exit Sub
Error_label:
    Base.ShowError "Voting.clearVotes"
End Sub


Public Sub loadVotes()
    On Error GoTo Error_label
    
    SendMessage "g"
    Base.Sleep 100
    lngStatus = CommRead(COMPort, strData, 1)
    
    If lngStatus < 0 Then
        lngStatus = CommGetError(strError)
        Err.Raise lngStatus, Description:=strError
    End If
    
    Set votes = Json.ParseJson(strData)
    
    If Not votes.Exists("total") Then
        Err.Raise 0, Description:="Invalid JSON respond: no 'total' key"
    End If
    
    logging.logVotes votes
    
    Dim totalVotes(7) As Integer ' Без этого >>Diagram.DrawValues votes("total")<< выбросит ошибку
    ' Копируем массив (на самом деле Variant) во временный (настоящий массив)
    Dim i As Integer
    For i = 1 To 7
        totalVotes(i) = votes("total")(i)
    Next i
    
    diagram.DrawValues totalVotes
    
    Exit Sub
Error_label:
    Base.ShowError "Voting.loadVotes"
End Sub


Public Sub checkConnection()
    On Error GoTo Error_label
    
    COMPort = Base.settings("voting")("port")
    
    lngStatus = CommOpen(COMPort, "COM" & COMPort, "baud=115200 parity=N data=8 stop=1")
    If lngStatus <> 0 Then
        MsgBox "Failed to connect (code: 1)", vbCritical, "PPVoting error"
        Exit Sub
    End If
    
    CommSetLine COMPort, LINE_RTS, True
    CommSetLine COMPort, LINE_DTR, True
    
    lngStatus = CommWrite(COMPort, "k")
    If lngStatus <> 1 Then
        MsgBox "Failed to send validation request (code: 2)", vbCritical, "PPVoting error"
        GoTo End_label
    End If
    
    Base.Sleep 100
    
    lngStatus = CommRead(COMPort, strData, 2)
    If lngStatus < 0 Then
        MsgBox "Failed to get validation respond (code: 3)", vbCritical, "PPVoting error"
        GoTo End_label
    End If
    
    If strData <> "OK" Then
        MsgBox "Validation failed (code: 4)", vbCritical, "PPVoting error"
        GoTo End_label
    End If
    
    MsgBox "Success", vbInformation, "PPVoting info"
    
End_label:
    CommSetLine COMPort, LINE_RTS, False
    CommSetLine COMPort, LINE_DTR, False
    
    CommClose COMPort
    
    Exit Sub
Error_label:
    Base.ShowError "Voting.checkConnection"
End Sub





