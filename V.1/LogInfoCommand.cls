VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogInfoCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Класс: LogInfoCommand.cls
' Описание: Команда для логирования информационных сообщений
Option Explicit

Implements ICommand
Implements PoolableObject

' Параметры команды
Private m_Message As String
Private m_Source As String
Private m_ID As String
Private m_IsActive As Boolean

' Инициализация
Private Sub Class_Initialize()
    m_ID = "LOGINFO_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
End Sub

' Настройка команды
Public Sub Initialize(ByVal message As String, Optional ByVal source As String = "")
    m_Message = message
    m_Source = source
End Sub

' ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' Запись в лог
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " [INFO] " & _
               IIf(Len(m_Source) > 0, "[" & m_Source & "] ", "") & _
               m_Message
End Sub

' ICommand.Undo - логирование не поддерживает отмену
Private Sub ICommand_Undo()
    Debug.Print "Cannot undo log command"
End Sub

' ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = "LogInfoCommand"
End Property

' ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ICommand_SupportsUndo = False ' Логирование не поддерживает отмену
End Property

' PoolableObject.Reset
Private Sub PoolableObject_Reset()
    m_Message = ""
    m_Source = ""
    m_IsActive = True
End Sub

' PoolableObject.Activate
Private Sub PoolableObject_Activate()
    m_IsActive = True
End Sub

' PoolableObject.Deactivate
Private Sub PoolableObject_Deactivate()
    m_IsActive = False
End Sub

' PoolableObject.ObjectID
Private Property Get PoolableObject_ObjectID() As String
    PoolableObject_ObjectID = m_ID
End Property

