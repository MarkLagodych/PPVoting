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
'   1. ��������� � Tools > VBAProject Properties.                     '
'   2. � ���� Conditional Compilation Arguments �������               '
'      "SHOW_ERRORS = 1" ��� ������� ����� ����������                 '
'      ��� "SHOW_ERRORS = 0" ����� �������� ��������� �� �������.     '
'                                                                     '
'                        * ������� ������� *                          '
'   3. � �������� ���� PowerPoint, ������� � ���� > ��������� ���...  '
'   4. �������� ������ "PowerPoint Add-In (*.ppam)"                   '
'      ("���������� PowerPoint (*.ppam)")                             '
'   5. ������� "���������"                                            '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit

' ���� �������� (������������) - ���� � ������� JSON,
' ��� �������� ����������� �� ������ ������ �����������.
' ��� ������ ���� ������������ �������. ��� ��������
' ����� ����� ������ (����������� ���� > ������ �����).
' ��������� �����:
' {
'    "timer": {                        -- ��������� �������
'        "use": <boolean>,             -- ������������ ������?
'        "totalTime_min": <integer>,   -- ����� �� ����������� (���)
'        "blushTime_min": <integer>,   -- �� ������� ����� �� ����� �������� ��� �� �����������?
'        "idleTime_sec": <integer>     -- ��� ������������� ������������ �������, �� ������� ������
'                                                   ���������� ���������� �������?
'    },
'
'    "voting": {                       -- ��������� �����������
'        "use": <boolean>,             -- ������������ �����������?
'        "port": <integer>             -- ����� COM-�����, � �������� ��������� ������ (ESP8266)
'                                                   ����� ������ � ���������� ���������
'                                                   (Win+R > devmgmt.msc > OK)
'    },
'
'    "logging": {                       -- ��������� ����� �����������
'        "use": <boolean>,              -- ���������� �� ���������� ����������� � ����?
'        "file": <string>               -- ���� � ����� (��� ������ ��� �����)
'    }
' }
Public FS As New FileSystemObject  ' ��� ������ ������
Public settings As Dictionary      ' ���� ����������� ���������

' ����� ������ OnSlideShowPageChange ��� ������������
' ���������� ���� ������ ���� ���
' ����� OnSlideShowTerminate ��������������� � False
Private started As Boolean

' ���� ������ ���������� ��������
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
Sub Auto_Open()
    
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
    
    diagramShowKeys(1) = 87  ' W
    diagramShowKeys(2) = 188 ' ,
    
End Sub


' ��� ��������� �������� ����� ������ ��� ������
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

' ��� ������ � Loging -- ��� ��������
Sub ShowRawError(functionName)

    #If SHOW_ERRORS = 1 Then
        MsgBox Title:="PPVoting Error", Prompt:="Error " & Err.number & " in " & functionName & ": " & vbNewLine & Err.Description
    #End If
    
End Sub


' ======================== ������� ������� ==========================
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
    
    ' ������� ���������� � 1
    fileName = ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text
    
    Dim file As TextStream
    Set file = FS.OpenTextFile(fileName, ForReading)
    Set settings = Json.ParseJson(file.ReadAll())
    file.Close
    
    Exit Sub
Error_label:
    ShowRawError "Base.LoadSettings"

End Sub




