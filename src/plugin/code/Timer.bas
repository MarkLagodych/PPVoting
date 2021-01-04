Attribute VB_Name = "Timer"
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'                        ТАЙМЕР ДЛЯ POWERPOINT                        '
' ------------------------------------------------------------------- '
'                                                                     '
' Задачи этого модуля:                                                '
'   - рисовать в окне MS PowerPoint в верхнем левом углу              '
'     прямоугольник с текущим временем в формате ЧЧ:ММ:СС и время     '
'     до конца презентации в формате ММ:СС;                           '
'   - за указанное кол-во минут до конца, сменить фон на красноватый; '
'   - если в данный момент идёт анимированная смена слайдов, то       '
'     остановиться на заданноое кол-во секунд во избежание мигания.   '
'     Дело в том, что слайд логически ещё не сменился на следующий, а '
'     отрисовка таймера осуществляется на текущем. То есть, анимация  '
'     прерывается на возврат к текущему слайду, происходит отрисовка  '
'     прямоугольника и потом анимация снова продолжается).            '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '


Option Explicit

' Флаг, нужный для безошибочного завершения
' По умолчанию false,
' После успешного начала true,
' В конце false
Private started As Boolean

' Если на данном слайде уже есть прямоугольник со временем, его нужно
' обновить, а не создавать новый. Поиск нужной фигуры из присутствующих
' возможен по имени фигуры.
Const TimerShapeName = "SPECIAL_TIMER_SHAPE_NAME"

' Время в секундах
Public currentTime As Long ' Текущее время

Public endTime As Long     ' Время конца презентации

Public blushTime As Long   ' Кол-во минут до конца чтобы покраснеть

Public idleTime As Integer ' Время простоя - если нажата хоть одна кнопка
                           ' смены слайдов, остановить обновление таймера
                           ' на это время

Public timerID             ' Системный ID таймера

' Оставшееся время простоя
Private remainingIdleTime As Integer


Public Sub start()
    On Error GoTo Error_label

    remainingIdleTime = 0
    
    currentTime = getTime()
    
    If Not Base.settings("timer").Exists("totalTime_min") _
    Or Not Base.settings("timer").Exists("blushTime_min") Then
        Err.Raise 0, Description:="timer.totalTime_min or timer.blushTime_min do not exist"
    End If
    
    If notNumber(Base.settings("timer")("totalTime_min")) _
    Or notNumber(Base.settings("timer")("blushTime_min")) Then
        Err.Raise 0, Description:="timer.totalTime_min or timer.blushTime_min are invalid"
    End If
    
    ' Чтение настроек из общево файла конфигурации PPVoting
    If Not Base.settings("timer").Exists("totalTime_min") Then
        endTime = currentTime + 15 * 60 ' 15 минут
    Else
        endTime = currentTime + Base.settings("timer")("totalTime_min") * 60
    End If
    
    If Not Base.settings("timer").Exists("blushTime_min") Then
        blushTime = 3 * 60 ' 3 минуты
    Else
        blushTime = Base.settings("timer")("blushTime_min") * 60
    End If
    
    timerID = Base.SetTimer(0, 0, 1000, AddressOf Run) ' Повторять каждые 1000 мс
    
    started = True
    
    Exit Sub
Error_label:
    Base.ShowError "Timer.Start"
    
End Sub


Public Sub finish()
    On Error GoTo Error_label
    
    ' Безопасное завершение
    If started Then
        Base.KillTimer 0, timerID
    End If
    
    started = False
        
    Exit Sub
Error_label:
    started = False
    Base.ShowError "Timer.Finish"
End Sub


Function getTime() As Long

    Dim currentMoment
    currentMoment = Now()
    getTime = Hour(currentMoment) * 3600 + Minute(currentMoment) * 60 + Second(currentMoment)
    
End Function


Sub Run()
    On Error GoTo Error_label
    
    If Diagram.Visible Then Exit Sub ' Тоже избежание мигания
    
    Dim slideIndex As Integer
    slideIndex = ActivePresentation.SlideShowWindow.View.slide.slideIndex
    
    Dim myslide As slide
    Set myslide = ActivePresentation.Slides(slideIndex)
    
    Dim key As Variant
    For Each key In Base.forwardKeys
        If Base.CheckKeyPressed(key) Then remainingIdleTime = myslide.SlideShowTransition.Duration
    Next
    For Each key In Base.backwardKeys
        If Base.CheckKeyPressed(key) Then remainingIdleTime = myslide.SlideShowTransition.Duration
    Next
    
    If remainingIdleTime <> 0 Then
    
        remainingIdleTime = remainingIdleTime - 1
        
    Else ' remainingIdleTime = 0

        currentTime = getTime()
        
        ' Первый слайд обязан иметь ТОЛЬКО надпись с путём к файлу настроек PPVoting!
        ' Поэтому, на нём нельзя рисовать
        If slideIndex = 1 Then Exit Sub
        
        Dim myshape As Shape
        Set myshape = getShapeByName(TimerShapeName, slideIndex)
    
        ' Создать/изменить вид и содержание прямоугольника
        If myshape Is Nothing Then
            With myslide.Shapes.AddShape(msoShapeRectangle, 0, 0, 110, 20)
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Fill.BackColor.RGB = RGB(255, 255, 255)
                .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
                .name = TimerShapeName
                .TextFrame.TextRange.Text = FormatTime() ' 12:34:56 | 01:23
                .TextFrame.TextRange.Font.name = "Arial"
                .TextFrame.TextRange.Font.Bold = True
                .TextFrame.TextRange.Font.Size = 12
            End With
        Else
            If endTime - currentTime <= blushTime Then
                myshape.Fill.ForeColor.RGB = RGB(255, 109, 109)
                myshape.Fill.BackColor.RGB = RGB(255, 109, 109)
            Else
                myshape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                myshape.Fill.BackColor.RGB = RGB(255, 255, 255)
            End If
            
            myshape.TextFrame.TextRange.Text = FormatTime() ' 12:34:56 | 01:23
        End If
        
    End If ' remainingIdleTime = 0
    
    
    Exit Sub
Error_label:
    Base.ShowError "Timer.Run"
    
End Sub


Function getShapeByName(name As String, slideIndex As Integer) As Shape
    On Error GoTo Error_label
    
    Dim currentSlide As slide
    Set currentSlide = ActivePresentation.Slides(slideIndex)
    
    Dim slideShape As Variant
    
    For Each slideShape In currentSlide.Shapes
        If slideShape.name = name Then
            Set getShapeByName = slideShape
            Exit Function
        End If
    Next
    
    Set getShapeByName = Nothing
    
    
    Exit Function
Error_label:
    Base.ShowError "Timer.getShapeByName"
    
End Function


' Получить 2 символа
' "1" -> "01"
' "37" -> "37"
Function get2(x As Integer) As String
    If Len("" & x) = 1 Then
        get2 = "0" & x
    Else
        get2 = "" & x
    End If
End Function


' Возвращает "HH:MM:SS | MM:SS" - текущее | оставшееся время
Function FormatTime() As String

    Dim remainingTime As Long
    remainingTime = endTime - currentTime

    FormatTime = "" _
        & get2(currentTime \ 3600) _
        & ":" _
        & get2((currentTime Mod 3600) \ 60) _
        & ":" _
        & get2(currentTime Mod 60) _
        & " | " _
        & get2(remainingTime \ 60) _
        & ":" _
        & get2(remainingTime Mod 60)
        
End Function



