VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandInvoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Класс: CommandInvoker.cls
' Описание: Инвокер команд - класс, запускающий выполнение команд
Option Explicit

' История выполненных команд для отмены
Private m_CommandHistory As Collection

' Инициализация
Private Sub Class_Initialize()
    Set m_CommandHistory = New Collection
End Sub

' Очистка
Private Sub Class_Terminate()
    Set m_CommandHistory = Nothing
End Sub

' Выполнение команды
Public Function ExecuteCommand(ByVal command As ICommand) As Boolean
    On Error GoTo ErrorHandler
    
    ' Выполнение команды
    command.Execute
    
    ' Если команда поддерживает отмену, добавляем в историю
    If command.SupportsUndo Then
        m_CommandHistory.Add command
    End If
    
    ExecuteCommand = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in ExecuteCommand: " & Err.Description
    ExecuteCommand = False
End Function

' Отмена последней команды
Public Function UndoLastCommand() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем наличие команд в истории
    If m_CommandHistory.Count = 0 Then
        UndoLastCommand = False
        Exit Function
    End If
    
    ' Получаем последнюю команду
    Dim command As ICommand
    Set command = m_CommandHistory(m_CommandHistory.Count)
    
    ' Отменяем выполнение
    command.Undo
    
    ' Удаляем из истории
    m_CommandHistory.Remove m_CommandHistory.Count
    
    UndoLastCommand = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in UndoLastCommand: " & Err.Description
    UndoLastCommand = False
End Function

' Проверка возможности отмены
Public Property Get CanUndo() As Boolean
    CanUndo = (m_CommandHistory.Count > 0)
End Property

' Количество команд в истории
Public Property Get CommandHistoryCount() As Long
    CommandHistoryCount = m_CommandHistory.Count
End Property

' Очистка истории команд
Public Sub ClearHistory()
    Set m_CommandHistory = New Collection
End Sub
