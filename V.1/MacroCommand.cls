VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Класс: MacroCommand.cls
' Описание: Реализация макрокоманды (композитный паттерн Command)
Option Explicit

Implements ICommand
Implements PoolableObject

' Коллекция команд в макрокоманде
Private m_Commands As Collection
Private m_Name As String
Private m_ID As String
Private m_IsActive As Boolean

' Инициализация
Private Sub Class_Initialize()
    Set m_Commands = New Collection
    m_Name = "MacroCommand"
    m_ID = "MACRO_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
End Sub

' Очистка
Private Sub Class_Terminate()
    Set m_Commands = Nothing
End Sub

' Установка имени макрокоманды
Public Property Let Name(ByVal value As String)
    m_Name = value
End Property

' Получение имени макрокоманды
Public Property Get Name() As String
    Name = m_Name
End Property

' Добавление команды в макрокоманду
Public Sub AddCommand(ByVal command As ICommand)
    m_Commands.Add command
End Sub

' Удаление всех команд
Public Sub ClearCommands()
    Set m_Commands = New Collection
End Sub

' Получение количества команд
Public Property Get CommandCount() As Long
    CommandCount = m_Commands.Count
End Property

' Реализация ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' Выполнение всех команд последовательно
    Dim cmd As ICommand
    Dim i As Long
    
    For i = 1 To m_Commands.Count
        Set cmd = m_Commands(i)
        cmd.Execute
    Next i
End Sub

' Реализация ICommand.Undo
Private Sub ICommand_Undo()
    If Not m_IsActive Then Exit Sub
    
    ' Отмена всех команд в обратном порядке
    Dim cmd As ICommand
    Dim i As Long
    
    For i = m_Commands.Count To 1 Step -1
        Set cmd = m_Commands(i)
        If cmd.SupportsUndo Then
            cmd.Undo
        End If
    Next i
End Sub

' Реализация ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = m_Name
End Property

' Реализация ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ' Макрокоманда поддерживает отмену, если хотя бы одна из команд поддерживает
    Dim cmd As ICommand
    Dim i As Long
    
    For i = 1 To m_Commands.Count
        Set cmd = m_Commands(i)
        If cmd.SupportsUndo Then
            ICommand_SupportsUndo = True
            Exit Property
        End If
    Next i
    
    ICommand_SupportsUndo = False
End Property

' Реализация PoolableObject.Reset
Private Sub PoolableObject_Reset()
    Set m_Commands = New Collection
    m_Name = "MacroCommand"
    m_IsActive = True
End Sub

' Реализация PoolableObject.Activate
Private Sub PoolableObject_Activate()
    m_IsActive = True
End Sub

' Реализация PoolableObject.Deactivate
Private Sub PoolableObject_Deactivate()
    m_IsActive = False
End Sub

' Реализация PoolableObject.ObjectID
Private Property Get PoolableObject_ObjectID() As String
    PoolableObject_ObjectID = m_ID
End Property
