VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Класс: CommandDecorator.cls
' Описание: Базовый класс для всех декораторов команд
Option Explicit

Implements ICommand
Implements PoolableObject

' Декорируемая команда
Private m_DecoratedCommand As ICommand
' Идентификатор декоратора
Private m_ID As String

' Инициализация
Private Sub Class_Initialize()
    m_ID = "DECORATOR_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
End Sub

' Установка декорируемой команды
Public Sub SetCommand(ByVal command As ICommand)
    Set m_DecoratedCommand = command
End Sub

' Получение декорируемой команды
Public Property Get DecoratedCommand() As ICommand
    Set DecoratedCommand = m_DecoratedCommand
End Property

' Реализация ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_DecoratedCommand Is Nothing Then
        m_DecoratedCommand.Execute
    End If
End Sub

' Реализация ICommand.Undo
Private Sub ICommand_Undo()
    If Not m_DecoratedCommand Is Nothing Then
        m_DecoratedCommand.Undo
    End If
End Sub

' Реализация ICommand.Name
Private Property Get ICommand_Name() As String
    If Not m_DecoratedCommand Is Nothing Then
        ICommand_Name = "Decorated_" & m_DecoratedCommand.Name
    Else
        ICommand_Name = "EmptyDecorator"
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
