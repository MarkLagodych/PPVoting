Attribute VB_Name = "Base"
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'                ПЛАГИН PPVOTING ДЛЯ MICROSOFT POWERPOINT             '
' ------------------------------------------------------------------- '
' Эта программа является плагином для Microsoft PowerPoint. Она       '
' добавляет таймер, указывающий на текущее и оставшееся время,        '
' а также занимается приёмом информации от голосовальной системы и ёё '
' отображением в виде диаграммы.                                      '
' ------------------------------------------------------------------- '
'                                                                     '
'                 * Настройка сообщений об ошибках *                  '
'   1. Перейдите в Tools > VBAProject Properties.                     '
'   2. В поле Conditional Compilation Arguments введите               '
'      "SHOW_ERRORS = 1" без кавычек чтобы отображать                 '
'      или "SHOW_ERRORS = 0" чтобы скрывать сообщения об ошибках.     '
'                                                                     '
'                        * Экспорт плагина *                          '
'   3. В основном окне PowerPoint, зайдите в Файл > Сохранить как...  '
'   4. Выберите формат "PowerPoint Add-In (*.ppam)"                   '
'      ("Надстройка PowerPoint (*.ppam)")                             '
'   5. Нажмите "Сохранить"                                            '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit

' Файл настроек (конфигурации) - файл в формате JSON,
' имя которого указывается на первом слайде презентации.
' Это должна быть ЕДИНСТВЕННАЯ надпись. Для удобства
' слайд можно скрыть (контекстное меню > скрыть слайд).
' Структура файла:
' {
'    "timer": {                        -- Настройки таймера
'        "use": <boolean>,             -- Использовать таймер?
'        "totalTime_min": <integer>,   -- Время на презентация (мин)
'        "blushTime_min": <integer>,   -- За сколько минут до конца изменить фон на красноватый?
'        "idleTime_sec": <integer>     -- При анимированном переключении слайдов, на сколько секунд
'                                                   остановить обновление таймера?
'    },
'
'    "voting": {                       -- Настройки голосования
'        "use": <boolean>,             -- Использовать голосование?
'        "port": <integer>             -- Номер COM-порта, к которому подключен сервер (ESP8266)
'                                                   Можно узнать в диспетчере устройств
'                                                   (Win+R > devmgmt.msc > OK)
'    },
'
'    "logging": {                       -- Настройки очёта голосований
'        "use": <boolean>,              -- Записывать ли результаты голосований в файл?
'        "file": <string>               -- Путь к файлу (или просто имя фалйа)
'    }
' }
Public FS As New FileSystemObject  ' Для чтения файлов
Public settings As Dictionary      ' Сюда прочитаются настройки

' Нужно внутри OnSlideShowPageChange для срабатывания
' начального кода только один раз
' После OnSlideShowTerminate устанавливается в False
Private started As Boolean

' Коды кнопок управления плагином
Public forwardKeys(6), backwardKeys(6), diagramShowKeys(2) As Integer


#If VBA7 Then

    Public Declare PtrSafe Function SetTimer _
        Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) _
        As LongPtr
                             
    Public Declare PtrSafe Function KillTimer _
        Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) _
        As LongPtr
                            
    Public Declare PtrSafe Function CheckKeyPressed _
        Lib "user32" Alias "GetAsyncKeyState" _
        (ByVal vKey As Long) _
        As Integer
        
    Public Declare PtrSafe Sub Sleep _
        Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
        
#Else

    Public Declare Function SetTimer _
        Lib "user32" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) _
        As Long
                             
    Public Declare Function KillTimer _
        Lib "user32" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long) _
        As Long
        
    Public Declare Function CheckKeyPressed _
        Lib "user32" Alias "GetAsyncKeyState" _
        (ByVal vKey As Long) _
        As Integer
        
    Public Declare Sub Sleep _
        Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#End If

' Проверка на числовой тип
Public Function notNumber(x As Variant) As Boolean
    Select Case VarType(x)
        Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            notNumber = False
        Case Else
            notNumber = True
    End Select
End Function


' Вызывается один раз при загрузке плагина
Sub Auto_Open()
    
    forwardKeys(1) = 39 ' Стрелка вправо
    forwardKeys(2) = 40 ' Стрелка вниз
    forwardKeys(3) = 34 ' Page DOWN
    forwardKeys(4) = 78 ' N
    forwardKeys(5) = 13 ' Enter
    forwardKeys(6) = 32 ' Пробел
    
    backwardKeys(1) = 37 ' Стрелка влево
    backwardKeys(2) = 38 ' Стрелка вверх
    backwardKeys(3) = 33 ' Page UP
    backwardKeys(4) = 80 ' P
    
    diagramShowKeys(1) = 87  ' W
    diagramShowKeys(2) = 188 ' ,
    
End Sub


' Эту процедуру вызывает любая другая при ошибке
Sub ShowError(functionName As String)
    Dim num As Integer
    Dim desc As String
    num = Err.number
    desc = Err.Description
    
    Logging.logError functionName, num, desc

    #If SHOW_ERRORS = 1 Then
        MsgBox Title:="PPVoting Error", Prompt:="Error " & num & " in " & functionName & ": " & vbNewLine & desc
    #End If
    
End Sub

' Для вызова в Loging -- без рекурсии
Sub ShowRawError(functionName)

    #If SHOW_ERRORS = 1 Then
        MsgBox Title:="PPVoting Error", Prompt:="Error " & Err.number & " in " & functionName & ": " & vbNewLine & Err.Description
    #End If
    
End Sub


' ======================== ГЛАВНАЯ ФУНКЦИЯ ==========================
Sub OnSlideShowPageChange()
    On Error GoTo Error_label

    If Not started Then
    
        started = True
        LoadSettings
        
        If settings.Exists("logging") Then
            If settings("logging").Exists("use") Then
                If settings("logging")("use") Then Logging.start
            End If
        End If
        
        If settings.Exists("timer") Then
            If settings("timer").Exists("use") Then
                If settings("timer")("use") Then Timer.start
            End If
        End If
        
        If settings.Exists("voting") Then
            If settings("voting").Exists("use") Then
                If settings("voting")("use") Then Voting.start
            End If
        End If
        
        
        
    End If
    
    Exit Sub
Error_label:
    ShowError "Base.OnSlideShowPageChange"
    
End Sub


Sub OnSlideShowTerminate()
    On Error GoTo Error_label
    
    started = False
    If settings.Exists("timer") Then If settings("timer")("use") Then Timer.finish
    If settings.Exists("voting") Then If settings("voting")("use") Then Voting.finish
    If settings.Exists("logging") Then If settings("logging")("use") Then Logging.finish
    
    Exit Sub
Error_label:
    ShowError "Base.OnSlideShowTerminate"
    
End Sub


Public Sub LoadSettings()
    On Error GoTo Error_label
        
    Dim fileName As String
    
    ' Массивы начинаются с 1
    fileName = ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text
    
    Dim file As TextStream
    Set file = FS.OpenTextFile(fileName, ForReading)
    Set settings = Json.ParseJson(file.ReadAll())
    file.Close
    
    Exit Sub
Error_label:
    ShowRawError "Base.LoadSettings"

End Sub




