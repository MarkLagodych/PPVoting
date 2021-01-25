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

Public Sub start()
    LoadSettings
End Sub

' Вызывается один раз при загрузке вкладки меню
Public Sub OnUILoaded(rbn As IRibbonUI)
    Set ribbon = rbn
End Sub

Public Sub LoadSettings()
    On Error GoTo Error_label
    
    If ribbon Is Nothing Then
        Err.Raise 1, Description:="UI is not initialized"
    End If
    
    ' Загружаем настройки. Если каких-либо настроек не существует, оставить те, что были установлены в Base.Auto_Open()
    
    Base.settings("logging")("use") = CBool(GetSetting(appName, "logging", "use", Base.settings("logging")("use")))
    Base.settings("logging")("file") = GetSetting(appName, "logging", "file", Base.settings("logging")("file"))
    
    Base.settings("timer")("use") = CBool(GetSetting(appName, "timer", "use", Base.settings("timer")("use")))
    Base.settings("timer")("total_time") = CInt(GetSetting(appName, "timer", "total_time", Base.settings("timer")("total_time")))
    Base.settings("timer")("total_time") = CInt(GetSetting(appName, "timer", "total_time", Base.settings("timer")("total_time")))
    
    Base.settings("voting")("use") = CBool(GetSetting(appName, "voting", "use", Base.settings("voting")("use")))
    Base.settings("voting")("port") = CInt(GetSetting(appName, "voting", "port", Base.settings("voting")("port")))
    Base.settings("voting")("diagram_width") = CInt(GetSetting(appName, "voting", "diagram_width", Base.settings("voting")("diagram_width")))
    Base.settings("voting")("diagram_height") = CInt(GetSetting(appName, "voting", "diagram_height", Base.settings("voting")("diagram_height")))
    Base.settings("voting")("diagram_gap") = CInt(GetSetting(appName, "voting", "diagram_gap", Base.settings("voting")("diagram_gap")))
    
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
        MsgBox "Integer positive number required!", vbCritical, "PPVoting error"
        ribbon.InvalidateControl control.id
        Exit Sub
    End If
    
    If CInt(text) <= 0 Then
        MsgBox "Integer positive number required!", vbCritical, "PPVoting error"
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
