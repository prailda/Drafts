VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogErrorCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Класс: LogErrorCommand.cls
' Описание: Команда для логирования сообщений об ошибках
Option Explicit

Implements ICommand
Implements PoolableObject

' Параметры команды
Private m_Message As String
Private m_Source As String
Private m_ErrorCode As Long
Private m_Details As String
Private m_ID As String
Private m_IsActive As Boolean

' Инициализация
Private Sub Class_Initialize()
    m_ID = "LOGERROR_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
    m_ErrorCode = 0
End Sub

' Настройка команды
Public Sub Initialize(ByVal message As String, Optional ByVal source As String = "", _
                      Optional ByVal errorCode As Long = 0, Optional ByVal Details As String = "")
    m_Message = message
    m_Source = source
    m_ErrorCode = errorCode
    m_Details = Details
End Sub

' ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' Формирование сообщения об ошибке
    Dim LogMessage As String
    LogMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & " [ERROR] "
    
    If Len(m_Source) > 0 Then
        LogMessage = LogMessage & "[" & m_Source & "] "
    End If
    
    LogMessage = LogMessage & m_Message
    
    If m_ErrorCode <> 0 Then
        LogMessage = LogMessage & " (Code: " & m_ErrorCode & ")"
    End If
    
    If Len(m_Details) > 0 Then
        LogMessage = LogMessage & vbCrLf & "Details: " & m_Details
    End If
    
    ' Запись в лог
    Debug.Print LogMessage
    
    ' В реальном приложении можно также записать в файл
    ' или отправить в систему мониторинга ошибок
End Sub

' ICommand.Undo - логирование не поддерживает отмену
Private Sub ICommand_Undo()
    Debug.Print "Cannot undo error log command"
End Sub

' ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = "LogErrorCommand"
End Property

' ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ICommand_SupportsUndo = False ' Логирование не поддерживает отмену
End Property

' PoolableObject.Reset
Private Sub PoolableObject_Reset()
    m_Message = ""
    m_Source = ""
    m_ErrorCode = 0
    m_Details = ""
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

