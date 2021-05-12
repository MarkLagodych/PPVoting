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
'   В Tools > VBAProject Properties в поле                            '
'   "Conditional Compilation Arguments" введите "SHOW_ERRORS = 1" без '
'   кавычек чтобы отображать или "SHOW_ERRORS = 0" чтобы скрывать     '
'   сообщения об ошибках.                                             '
'                                                                     '
'                        * Экспорт плагина *                          '
'   Файл > Сохранить как... > PowerPoint Add-In (*.ppam)              '
'                          (Надстройка PowerPoint (*.ppam))           '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit


Public FS As New FileSystemObject  ' Для чтения файлов
Public settings As New Dictionary  ' Сюда прочитаются настройки

' Нужно внутри OnSlideShowPageChange для срабатывания
' начального кода только один раз
' После OnSlideShowTerminate устанавливается в False
Private started As Boolean

' Коды кнопок управления плагином
Public forwardKeys(6), backwardKeys(6), diagramShowKeys(2) As Integer

' Для запуска различных команд (как в командной строке)
Public CmdShell As Object


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
Public Sub Auto_Open()
    
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
    
    diagramShowKeys(1) = 66  ' B
    diagramShowKeys(2) = 190 ' .
    
    Set CmdShell = CreateObject("WScript.Shell")
    
End Sub


' Эту процедуру вызывает любая другая при ошибке
Sub ShowError(functionName As String)
    Dim num As Long
    Dim desc As String
    num = Err.number
    desc = Err.Description
    
    logging.logError functionName, num, desc

    #If SHOW_ERRORS = 1 Then
        MsgBox _
            Title:="PPVoting error", _
            Prompt:="Error " & num & " in " & functionName & ": " & vbNewLine & desc, _
            Buttons:=vbCritical
    #End If
    
End Sub

' Для вызова в Loging -- без рекурсии
Sub ShowRawError(functionName)

    #If SHOW_ERRORS = 1 Then
        MsgBox _
            Title:="PPVoting error", _
            Prompt:="Error " & Err.number & " in " & functionName & ": " & vbNewLine & Err.Description, _
            Buttons:=vbCritical
    #End If
    
End Sub


' ======================== ГЛАВНАЯ ФУНКЦИЯ ==========================
Sub OnSlideShowPageChange()
    On Error GoTo Error_label

    If Not started Then
    
        started = True
        
        ' Logging включаем первым
        If settings.Exists("logging") Then
            If settings("logging").Exists("use") Then
                If settings("logging")("use") Then logging.start
            End If
        End If
        
        If settings.Exists("timer") Then
            If settings("timer").Exists("use") Then
                If settings("timer")("use") Then timer.start
            End If
        End If
        
        If settings.Exists("voting") Then
            If settings("voting").Exists("use") Then
                If settings("voting")("use") Then voting.start
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
    If settings.Exists("timer") Then If settings("timer")("use") Then timer.finish
    If settings.Exists("voting") Then If settings("voting")("use") Then voting.finish
    If settings.Exists("logging") Then If settings("logging")("use") Then logging.finish
    ' Logging выключаем последним
    
    Exit Sub
Error_label:
    ShowError "Base.OnSlideShowTerminate"
    
End Sub





