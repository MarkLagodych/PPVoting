Attribute VB_Name = "PPVotingUI"
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'                   ВКЛАДКА "PPVOTING" ДЛЯ POWERPOINT                 '
' ------------------------------------------------------------------- '
'                                                                     '
' Этот модуль обрабатывает все виджеты во вкладке "PPVoting" в ленте  '
' PowerPoint. Все настройки сохраняются с помощью SaveSetting.        '
'                                                                     '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit

Private ribbon As Office.IRibbonUI

Const appName = "PPVoting"

' Вызывается один раз при загрузке вкладки меню
Public Sub OnUILoaded(rbn As IRibbonUI)
    Set ribbon = rbn
    LoadSettings
End Sub

Public Sub LoadSettings()
    On Error GoTo Error_label
    
    If ribbon Is Nothing Then
        Err.Raise 1, Description:="UI is not initialized"
    End If
    
    ' Создание словаря настроек и установка значений по умолчанию
    settings.Add "logging", New Dictionary
    settings("logging").Add "use", False
    settings("logging").Add "file", ""
    
    settings.Add "timer", New Dictionary
    settings("timer").Add "use", False
    settings("timer").Add "total_time", 2
    settings("timer").Add "blush_time", 1
    
    settings.Add "voting", New Dictionary
    settings("voting").Add "use", False
    settings("voting").Add "port", 1
    settings("voting").Add "diagram_width", 640
    settings("voting").Add "diagram_height", 480
    settings("voting").Add "diagram_gap", 10
    
    ' Загрузка настройки. Если какая-либо настройка не сохранена, оставить ту, что была установлена
    Dim section, key As Variant
    For Each section In Base.settings.Keys()
        For Each key In Base.settings(section).Keys()
            If key = "use" Then
                settings(section)(key) = CBool(GetSetting(appName, section, key, settings(section)(key)))
            ElseIf key = "file" Then
                settings(section)(key) = GetSetting(appName, section, key, settings(section)(key))
            Else
                settings(section)(key) = CInt(GetSetting(appName, section, key, settings(section)(key)))
            End If
        Next key
    Next section
    
    ' Обновить все виджеты
    ribbon.Invalidate
    
    Exit Sub
Error_label:
    Base.ShowError "PPVotingUI.LoadSettings"
    
End Sub

Public Sub RemoveSettings()
    
    On Error GoTo Error_label
    
    DeleteSetting appName
    MsgBox "Successfully removed all PPVoting settings", vbInformation, "PPVoting info"
    
    Exit Sub
Error_label:
    MsgBox "No PPVoting settings present", vbExclamation, "PPVoting info"
    
End Sub

Public Sub ProcessText(control As IRibbonControl, text As String)

    If Not IsNumeric(text) Then
        MsgBox "Positive integer number required!", vbCritical, "PPVoting error"
        ribbon.InvalidateControl control.id
        Exit Sub
    End If
    
    If CInt(text) <= 0 Then
        MsgBox "Positive integer number required!", vbCritical, "PPVoting error"
        ribbon.InvalidateControl control.id
        Exit Sub
    End If
    
    Dim parts() As String
    parts = Split(control.Tag, ".")
    
    Base.settings(parts(0))(parts(1)) = CInt(text)
    SaveSetting appName, parts(0), parts(1), text
    
End Sub

Public Sub GetText(control As IRibbonControl, ByRef text)
    
    Dim parts() As String
    parts = Split(control.Tag, ".")
    
    text = Base.settings(parts(0))(parts(1))
    
End Sub

Public Sub ProcessBool(control As IRibbonControl, state As Boolean)
    
    Dim parts() As String
    parts = Split(control.Tag, ".")
    
    Base.settings(parts(0))(parts(1)) = state
    SaveSetting appName, parts(0), parts(1), CStr(state)
    
End Sub

Public Sub GetBool(control As IRibbonControl, ByRef state)
    
    Dim parts() As String
    parts = Split(control.Tag, ".")
    
    state = Base.settings(parts(0))(parts(1))
    
End Sub

Public Sub checkConnection()
    voting.checkConnection
End Sub


Public Sub ChooseLogFile()
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .InitialFileName = "*.txt"
        If .Show <> -1 Then Exit Sub
        Dim path As String
        path = .SelectedItems(1)
        Base.settings("logging")("file") = path
        SaveSetting appName, "logging", "file", path
    End With
End Sub

Public Sub ViewLogFile()
    On Error GoTo Error_label
    
    Dim q As String
    q = """" ' На самом деле, одна двойная кавычка
    
    Dim path As String
    path = Base.settings("logging")("file")
    
    If path = "" Then
        MsgBox "No log file selected", vbInformation, "PPVoting info"
        Exit Sub
    End If
    
    If Dir(path) = "" Then
        MsgBox "Can not find " & q & path & q, vbExclamation, "PPVoting info"
        Exit Sub
    End If
    
    Base.CmdShell.Run "notepad " & q & path & q
    
    Exit Sub
Error_label:
    Base.ShowError "PPVotingUI.ViewLogFile"
End Sub

Public Sub OpenDeviceManager()
    On Error GoTo Error_label
    
    Base.CmdShell.Run "devmgmt.msc"
    
    Exit Sub
Error_label:
    Base.ShowError "PPVotingUI.OpenDeviceManager"
End Sub
