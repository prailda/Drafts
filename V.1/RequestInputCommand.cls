VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RequestInputCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Класс: RequestInputCommand.cls
' Описание: Команда для запроса ввода от пользователя
Option Explicit

Implements ICommand
Implements PoolableObject

' Параметры команды
Private m_Prompt As String
Private m_Title As String
Private m_DefaultValue As String
Private m_Result As String
Private m_ID As String
Private m_IsActive As Boolean

' Инициализация
Private Sub Class_Initialize()
    m_ID = "REQINPUT_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
    m_Result = ""
End Sub

' Настройка команды
Public Sub Initialize(ByVal prompt As String, Optional ByVal title As String = "Ввод данных", _
                      Optional ByVal defaultValue As String = "")
    m_Prompt = prompt
    m_Title = title
    m_DefaultValue = defaultValue
    m_Result = ""
End Sub

' ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' Запрос ввода от пользователя
    m_Result = InputBox(m_Prompt, m_Title, m_DefaultValue)
    
    ' Запись в лог
    Debug.Print "User input requested: " & m_Title & " - " & m_Prompt
    Debug.Print "User input result: " & IIf(Len(m_Result) > 0, m_Result, "(empty or cancelled)")
End Sub

' ICommand.Undo - запрос ввода не поддерживает отмену
Private Sub ICommand_Undo()
    Debug.Print "Cannot undo user input request"
End Sub

' ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = "RequestInputCommand"
End Property

' ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ICommand_SupportsUndo = False ' Не поддерживает отмену
End Property

' PoolableObject.Reset
Private Sub PoolableObject_Reset()
    m_Prompt = ""
    m_Title = "Ввод данных"
    m_DefaultValue = ""
    m_Result = ""
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

' Получение результата ввода
Public Property Get result() As String
    result = m_Result
End Property

