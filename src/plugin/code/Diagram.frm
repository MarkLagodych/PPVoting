VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Diagram 
   Caption         =   "Äטאדנאללא"
   ClientHeight    =   13680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15765
   OleObjectBlob   =   "Diagram.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Diagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private minLineLength, maxLineLength As Integer
Private lines(7) As Label
Private nullValues(7) As Integer

Public Sub DrawValues(values() As Integer)
    On Error GoTo Error_label
    
    Dim minValue, maxValue As Integer
    minValue = values(1)
    maxValue = values(1)
    Dim value
    For Each value In values
        If value > maxValue Then maxValue = value
        If value < minValue Then minValue = value
    Next value
    
    Dim k As Integer
    
    If maxValue = minValue Then ' If aguments are the same
        k = maxLineLength - minLineLength
    Else
        k = (maxLineLength - minLineLength) / (maxValue - minValue)
    End If
    
    Dim i As Integer
    For i = 1 To 7
        lines(i).Width = minLineLength + values(i) * k
        lines(i).Caption = values(i)
    Next i
    
    Exit Sub
Error_label:
    Base.ShowError "Diagram.DrawValues"
End Sub

Private Sub UserForm_Initialize()
    minLineLength = Line1.Width ' Set in form editor
    maxLineLength = Line7.Width
    
    Set lines(1) = Line1
    Set lines(2) = Line2
    Set lines(3) = Line3
    Set lines(4) = Line4
    Set lines(5) = Line5
    Set lines(6) = Line6
    Set lines(7) = Line7
    
    Dim i As Integer
    For i = 1 To 7
        nullValues(i) = 0
    Next i
    
    clearVotes
End Sub


Public Sub clearVotes()
    DrawValues nullValues
End Sub

