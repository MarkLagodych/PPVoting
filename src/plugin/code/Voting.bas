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
    
    If Not Base.settings("voting").Exists("port") Then
        Err.Raise 0, Description:="voting.port does not exist"
    ElseIf notNumber(Base.settings("voting")("port")) Then
        Err.Raise 0, Description:="voting.port is not a number"
    End If
    
    COMPort = Base.settings("voting")("port")
    
    lngStatus = CommOpen(COMPort, "COM" & COMPort, "baud=115200 parity=N data=8 stop=1")
    If lngStatus <> 0 Then
        lngStatus = CommGetError(strError)
        Err.Raise lngStatus, Description:=strError
    End If
    
    lngStatus = CommSetLine(COMPort, LINE_RTS, True)
    lngStatus = CommSetLine(COMPort, LINE_DTR, True)

    timerID = Base.SetTimer(0, 0, 100, AddressOf Run) ' Повторять каждые 100 мс
    
    started = True
    
    Logging.logInfo "Voting.start", "Communication started on COM port " & COMPort & ", frequency 115200 bauds, 8N1 configuration"
    
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
            With Diagram
                If .Visible Then
                    .Hide
                    
                    ' Нужно чтобы убрать белый/чёрный экран
                    ActivePresentation.SlideShowWindow.View.GotoSlide lastSlide
                Else
                    lastSlide = ActivePresentation.SlideShowWindow.View.slide.slideIndex
                    .Show
                End If
            End With
        End If
        
    Next key
    
    
    If Diagram.Visible Then
        
        For Each key In Base.backwardKeys
            If Base.CheckKeyPressed(key) Then
                clearVotes
                Diagram.Hide
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
        
        Logging.logInfo "Voting.finish", "COM communication shut down"
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
    
    Diagram.clearVotes
    SendMessage "c"
    Logging.logInfo "Voting.clearVotes", "All votes cleared from the server"
    
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
    
    Logging.logVotes votes
    
    Dim totalVotes(7) As Integer ' Без этого >>Diagram.DrawValues votes("total")<< выбросит ошибку
    Dim i As Integer
    For i = 1 To 7
        totalVotes(i) = votes("total")(i)
    Next i
    
    Diagram.DrawValues totalVotes
    
    Exit Sub
Error_label:
    Base.ShowError "Voting.loadVotes"
End Sub





