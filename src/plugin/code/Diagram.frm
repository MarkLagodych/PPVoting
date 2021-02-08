VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Diagram 
   Caption         =   "Диаграмма"
   ClientHeight    =   11205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7905
   OleObjectBlob   =   "Diagram.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Diagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''
'                  ДИАГРАММА ОТВЕТОВ              '
'             С НАСТРАИВАЕМЫМИ РАЗМЕРАМИ          '
' ----------------------------------------------- '
' В редакторе форм нужно установить:              '
'   1. Х-координату всех надписей (это будет      '
'      константа)                                 '
'   2. Ширину всех надписей (тоже константа)      '
'   3. Х-координату всех линий (синие)            '
'   4. Ширину первой линии (это будет минимум)    '
' а всё остальное диаграмма возьмёт из            '
' настроек и выщитает сама                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private minLineWidth, maxLineWidth, lineHeight As Integer

Private winWidth, winHeight, gap As Integer

Private lines(7) As Label
Private labels(7) As Label
Private nullValues(7) As Integer

Public Sub DrawValues(values() As Integer)

    Dim minValue, maxValue As Integer
    minValue = values(1)
    maxValue = values(1)
    Dim value
    For Each value In values
        If value > maxValue Then maxValue = value
        If value < minValue Then minValue = value
    Next value
    
    Dim k As Integer
    
    If maxValue = minValue Then ' Если все данные значения одинаковы
        k = maxLineWidth - minLineWidth
    Else
        k = (maxLineWidth - minLineWidth) / (maxValue - minValue)
    End If
    
    Dim i As Integer
    For i = 1 To 7
        lines(i).Width = minLineWidth + values(i) * k
        lines(i).Caption = values(i)
    Next i
    
End Sub

Private Sub UserForm_Initialize()
    minLineWidth = Line1.Width

    winWidth = Base.settings("voting")("diagram_width")
    winHeight = Base.settings("voting")("diagram_height")
    gap = Base.settings("voting")("diagram_gap")
    
    maxLineWidth = winWidth - Line1.Left - gap
    lineHeight = (winHeight - 8 * gap) / 7
    
    Me.Width = Me.Width - Me.InsideWidth + winWidth
    Me.Height = Me.Height - Me.InsideHeight + winHeight
    
    Set lines(1) = Line1
    Set lines(2) = Line2
    Set lines(3) = Line3
    Set lines(4) = Line4
    Set lines(5) = Line5
    Set lines(6) = Line6
    Set lines(7) = Line7
    
    Set labels(1) = Label1
    Set labels(2) = Label2
    Set labels(3) = Label3
    Set labels(4) = Label4
    Set labels(5) = Label5
    Set labels(6) = Label6
    Set labels(7) = Label7
    
    Dim i, y As Integer
    
    y = gap
    
    For i = 1 To 7
        lines(i).Top = y
        labels(i).Top = y
        lines(i).Height = lineHeight
        y = y + lineHeight + gap
    Next i
    
    For i = 1 To 7
        nullValues(i) = 0
    Next i
    
    clearVotes
End Sub


Public Sub clearVotes()
    DrawValues nullValues
End Sub

