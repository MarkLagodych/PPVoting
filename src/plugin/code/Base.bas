Attribute VB_Name = "Base"
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'                ������ PPVOTING ��� MICROSOFT POWERPOINT             '
' ------------------------------------------------------------------- '
' ��� ��������� �������� �������� ��� Microsoft PowerPoint. ���       '
' ��������� ������, ����������� �� ������� � ���������� �����,        '
' � ����� ���������� ������ ���������� �� ������������� ������� � �� '
' ������������ � ���� ���������.                                      '
' ------------------------------------------------------------------- '
'                                                                     '
'                 * ��������� ��������� �� ������� *                  '
'   � Tools > VBAProject Properties � ����                            '
'   "Conditional Compilation Arguments" ������� "SHOW_ERRORS = 1" ��� '
'   ������� ����� ���������� ��� "SHOW_ERRORS = 0" ����� ��������     '
'   ��������� �� �������.                                             '
'                                                                     '
'                        * ������� ������� *                          '
'   ���� > ��������� ���... > PowerPoint Add-In (*.ppam)              '
'                          (���������� PowerPoint (*.ppam))           '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit


Public FS As New FileSystemObject  ' ��� ������ ������
Public settings As New Dictionary  ' ���� ����������� ���������

' ����� ������ OnSlideShowPageChange ��� ������������
' ���������� ���� ������ ���� ���
' ����� OnSlideShowTerminate ��������������� � False
Private started As Boolean

' ���� ������ ���������� ��������
Public forwardKeys(6), backwardKeys(6), diagramShowKeys(2) As Integer

' ��� ������� ��������� ������ (��� � ��������� ������)
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

' �������� �� �������� ���
Public Function notNumber(x As Variant) As Boolean
    Select Case VarType(x)
        Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            notNumber = False
        Case Else
            notNumber = True
    End Select
End Function


' ���������� ���� ��� ��� �������� �������
Public Sub Auto_Open()
    
    forwardKeys(1) = 39 ' ������� ������
    forwardKeys(2) = 40 ' ������� ����
    forwardKeys(3) = 34 ' Page DOWN
    forwardKeys(4) = 78 ' N
    forwardKeys(5) = 13 ' Enter
    forwardKeys(6) = 32 ' ������
    
    backwardKeys(1) = 37 ' ������� �����
    backwardKeys(2) = 38 ' ������� �����
    backwardKeys(3) = 33 ' Page UP
    backwardKeys(4) = 80 ' P
    
    diagramShowKeys(1) = 66  ' B
    diagramShowKeys(2) = 190 ' .
    
    Set CmdShell = CreateObject("WScript.Shell")
    
End Sub


' ��� ��������� �������� ����� ������ ��� ������
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

' ��� ������ � Loging -- ��� ��������
Sub ShowRawError(functionName)

    #If SHOW_ERRORS = 1 Then
        MsgBox _
            Title:="PPVoting error", _
            Prompt:="Error " & Err.number & " in " & functionName & ": " & vbNewLine & Err.Description, _
            Buttons:=vbCritical
    #End If
    
End Sub


' ======================== ������� ������� ==========================
Sub OnSlideShowPageChange()
    On Error GoTo Error_label

    If Not started Then
    
        started = True
        
        ' Logging �������� ������
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
    ' Logging ��������� ���������
    
    Exit Sub
Error_label:
    ShowError "Base.OnSlideShowTerminate"
    
End Sub





