VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggingDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Класс: LoggingDecorator.cls
' Описание: Декоратор, добавляющий логирование выполнения команд
Option Explicit

Implements ICommand
Implements PoolableObject

' Декорируемая команда
Private m_DecoratedCommand As ICommand
' Идентификатор декоратора
Private m_ID As String
' Дополнительные параметры логирования
Private m_LogLevel As String
Private m_LogPrefix As String

' Инициализация
Private Sub Class_Initialize()
    m_ID = "LOGDEC_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_LogLevel = "INFO"
    m_LogPrefix = "[LOGGING] "
End Sub

' Настройка декоратора
Public Sub Initialize(ByVal command As ICommand, Optional ByVal logLevel As String = "INFO", _
                      Optional ByVal logPrefix As String = "[LOGGING] ")
    Set m_DecoratedCommand = command
    m_LogLevel = logLevel
    m_LogPrefix = logPrefix
End Sub

' Реализация ICommand.Execute
Private Sub ICommand_Execute()
    ' Логирование до выполнения
    LogMessage "Начало выполнения команды " & m_DecoratedCommand.Name
    
    ' Выполнение декорируемой команды
    If Not m_DecoratedCommand Is Nothing Then
        m_DecoratedCommand.Execute
    End If
    
    ' Логирование после выполнения
    LogMessage "Завершение выполнения команды " & m_DecoratedCommand.Name
End Sub

' Реализация ICommand.Undo
Private Sub ICommand_Undo()
    ' Логирование отмены
    LogMessage "Начало отмены команды " & m_DecoratedCommand.Name
    
    ' Отмена декорируемой команды
    If Not m_DecoratedCommand Is Nothing Then
        m_DecoratedCommand.Undo
    End If
    
    ' Логирование после отмены
    LogMessage "Завершение отмены команды " & m_DecoratedCommand.Name
End Sub

' Логирование сообщения
Private Sub LogMessage(ByVal message As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " [" & m_LogLevel & "] " & _
               m_LogPrefix & message
End Sub

' Реализация ICommand.Name
Private Property Get ICommand_Name() As String
    If Not m_DecoratedCommand Is Nothing Then
        ICommand_Name = "LoggingDecorator(" & m_DecoratedCommand.Name & ")"
    Else
        ICommand_Name = "LoggingDecorator(Empty)"
    End If
End Property

' Реализация ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    If Not m_DecoratedCommand Is Nothing Then
        ICommand_SupportsUndo = m_DecoratedCommand.SupportsUndo
    Else
        ICommand_SupportsUndo = False
    End If
End Property

' Реализация PoolableObject.Reset
Private Sub PoolableObject_Reset()
    Set m_DecoratedCommand = Nothing
    m_LogLevel = "INFO"
    m_LogPrefix = "[LOGGING] "
End Sub

' Реализация PoolableObject.Activate
Private Sub PoolableObject_Activate()
    ' Если декорируемая команда поддерживает PoolableObject
    If Not m_DecoratedCommand Is Nothing Then
        If TypeOf m_DecoratedCommand Is PoolableObject Then
            Dim poolable As PoolableObject
            Set poolable = m_DecoratedCommand
            poolable.Activate
        End If
    End If
End Sub

' Реализация PoolableObject.Deactivate
Private Sub PoolableObject_Deactivate()
    ' Если декорируемая команда поддерживает PoolableObject
    If Not m_DecoratedCommand Is Nothing Then
        If TypeOf m_DecoratedCommand Is PoolableObject Then
            Dim poolable As PoolableObject
            Set poolable = m_DecoratedCommand
            poolable.Deactivate
        End If
    End If
End Sub

' Реализация PoolableObject.ObjectID
Private Property Get PoolableObject_ObjectID() As String
    PoolableObject_ObjectID = m_ID
End Property

