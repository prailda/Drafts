VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Класс: CommandPool.cls
' Описание: Пул для управления объектами-командами
Option Explicit

' Пул доступных объектов
Private m_AvailableCommands As Collection
' Пул используемых объектов
Private m_InUseCommands As Collection
Private m_Initialized As Boolean ' Добавить в раздел private переменных

' Параметры пула
Private m_PoolName As String
Private m_MaxPoolSize As Long
Private m_CommandLifetime As Long
Private m_CommandFactory As CommandFactory

' Инициализация
Private Sub Class_Initialize()
    Set m_AvailableCommands = New Collection
    Set m_InUseCommands = New Collection
    m_MaxPoolSize = 10 ' По умолчанию
    m_CommandLifetime = 30 ' 30 минут по умолчанию
    m_Initialized = False
End Sub

' Очистка ресурсов
Private Sub Class_Terminate()
    ClearAllObjects
    Set m_AvailableCommands = Nothing
    Set m_InUseCommands = Nothing
End Sub

' Получение экземпляра фабрики команд
Public Function GetCommandFactory() As CommandFactory
    ' Убедимся, что все компоненты инициализированы
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    If m_CommandFactory Is Nothing Then
        Set m_CommandFactory = New CommandFactory
    End If
    
    Set GetCommandFactory = m_CommandFactory
End Function

' Обновить метод InitializeAllPools
Public Sub InitializeAllPools()
    If m_ModuleInitialized Then Exit Sub
    
    ' Инициализируем пулы команд
    Set m_LogCommandPool = New CommandPool
    m_LogCommandPool.Initialize "LogCommands", 10, 30
    
    Set m_DataCommandPool = New CommandPool
    m_DataCommandPool.Initialize "DataCommands", 15, 30
    
    Set m_UICommandPool = New CommandPool
    m_UICommandPool.Initialize "UICommands", 8, 30
    
    ' Инициализируем инвокер и фабрику
    Set m_CommandInvoker = New CommandInvoker
    Set m_CommandFactory = New CommandFactory
    
    m_ModuleInitialized = True
    
    ' Выводим информацию об инициализации
    Debug.Print "=== Все компоненты системы инициализированы ==="
End Sub

' В методе Initialize добавить:
Public Sub Initialize(ByVal PoolName As String, Optional ByVal MaxPoolSize As Long = 10, _
                     Optional ByVal CommandLifetime As Long = 30)
    m_PoolName = PoolName
    m_MaxPoolSize = MaxPoolSize
    m_CommandLifetime = CommandLifetime
    m_Initialized = True ' Устанавливаем флаг инициализации
End Sub

' Добавляем новое свойство
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_Initialized
End Property

' Получение команды из пула по имени команды
Public Function GetCommand(ByVal commandName As String) As ICommand
    If Not m_Initialized Then
        Debug.Print "ERROR: Command pool not initialized"
        Set GetCommand = Nothing
        Exit Function
    End If
    
    ' Проверяем, есть ли доступные команды нужного типа в пуле
    Dim i As Long
    For i = 1 To m_AvailableCommands.Count
        Dim cmd As ICommand
        Set cmd = m_AvailableCommands(i)
        
        If cmd.Name = commandName Then
            ' Нашли команду нужного типа
            m_AvailableCommands.Remove i
            
            ' Активируем команду
            If TypeOf cmd Is PoolableObject Then
                Dim poolableCmd As PoolableObject
                Set poolableCmd = cmd
                poolableCmd.Activate
            End If
            
            ' Добавляем в список используемых
            Dim key As String
            key = "CMD_" & commandName & "_" & Format(Now, "yyyymmddhhnnss") & "_" & m_InUseCommands.Count + 1
            m_InUseCommands.Add cmd, key
            
            Set GetCommand = cmd
            Exit Function
        End If
    Next i
    
    ' Создание новой команды - исправленная версия

    Set cmd = CreateCommand(commandName)
    
    ' Если успешно создали, добавляем в используемые
    If Not cmd Is Nothing Then
        Dim newKey As String
        newKey = "CMD_" & commandName & "_" & Format(Now, "yyyymmddhhnnss") & "_" & m_InUseCommands.Count + 1
        m_InUseCommands.Add cmd, newKey
    End If
    
    Set GetCommand = cmd
End Function

' Создание новой команды по имени
' Исправленная версия метода для создания команд
' В классе CommandPool.cls - исправление метода создания команд

Private Function CreateCommand(ByVal commandName As String) As ICommand
    On Error Resume Next ' Добавляем обработку ошибок
    
    ' Фабричный метод для создания команд разных типов
    Select Case commandName
        ' Команды логирования
        Case "LogInfoCommand"
            Dim logInfoCmd As New LogInfoCommand
            Set CreateCommand = logInfoCmd
            
        Case "LogErrorCommand"
            Dim logErrorCmd As New LogErrorCommand
            Set CreateCommand = logErrorCmd
            
        ' Команды для работы с данными
        Case "AddRecordCommand"
            Dim addCmd As New AddRecordCommand
            Set CreateCommand = addCmd
            
        Case "DeleteRecordCommand"
            Dim delCmd As New DeleteRecordCommand
            Set CreateCommand = delCmd
            
        ' Команды для интерфейса
        Case "ShowMessageCommand"
            Dim msgCmd As New ShowMessageCommand
            Set CreateCommand = msgCmd
            
        Case "RequestInputCommand"
            ' Добавляем более надежное создание этой команды
            Dim inputCmd As New RequestInputCommand
            ' Проверяем, что объект создан
            If Err.Number <> 0 Then
                Debug.Print "Error creating RequestInputCommand: " & Err.Description
                Err.Clear
                Set CreateCommand = Nothing
                Exit Function
            End If
            Set CreateCommand = inputCmd
            
        Case Else
            Debug.Print "ERROR: Unknown command type: " & commandName
            Set CreateCommand = Nothing
    End Select
    
    ' Проверяем, что команда создана успешно
    If Err.Number <> 0 Then
        Debug.Print "Error creating command " & commandName & ": " & Err.Description
        Err.Clear
        Set CreateCommand = Nothing
    End If
    
    On Error GoTo 0
End Function

' Возврат команды в пул
Public Sub ReturnCommand(ByRef command As ICommand)
    If command Is Nothing Then Exit Sub
    
    ' Ищем команду в списке используемых
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 1 To m_InUseCommands.Count
        If m_InUseCommands(i) Is command Then
            found = True
            Exit For
        End If
    Next i
    
    ' Если команда найдена в списке используемых
    If found Then
        ' Удаляем из списка используемых
        m_InUseCommands.Remove i
        
        ' Деактивируем и сбрасываем команду
        If TypeOf command Is PoolableObject Then
            Dim poolableCmd As PoolableObject
            Set poolableCmd = command
            poolableCmd.Deactivate
            poolableCmd.Reset
        End If
        
        ' Проверяем, не превышен ли максимальный размер пула
        If m_AvailableCommands.Count < m_MaxPoolSize Then
            ' Добавляем команду в пул доступных
            m_AvailableCommands.Add command
        End If
    End If
    
    ' Очищаем ссылку на команду
    Set command = Nothing
End Sub

' Очистка всех команд в пулах
' Класс: CommandPool.cls
' Метод очистки, который нужно исправить

Public Sub ClearAllObjects()
    On Error Resume Next ' Добавляем обработку ошибок
    
    Dim i As Long
    
    ' Очистка доступных команд
    For i = m_AvailableCommands.Count To 1 Step -1
        Set m_AvailableCommands(i) = Nothing
        m_AvailableCommands.Remove i
    Next i
    
    ' Очистка используемых команд
    For i = m_InUseCommands.Count To 1 Step -1
        Set m_InUseCommands(i) = Nothing
        m_InUseCommands.Remove i
    Next i
    
    On Error GoTo 0
End Sub

' Получение статистики пула
Public Function GetStatistics() As String
    Dim stats As String
    stats = "===== " & m_PoolName & " Pool Statistics =====" & vbCrLf
    stats = stats & "Available commands: " & m_AvailableCommands.Count & vbCrLf
    stats = stats & "In-use commands: " & m_InUseCommands.Count & vbCrLf
    stats = stats & "Max pool size: " & m_MaxPoolSize & vbCrLf
    stats = stats & "Command lifetime: " & m_CommandLifetime & " minutes" & vbCrLf
    stats = stats & "=================================="
    
    GetStatistics = stats
End Function

' Класс: CommandPool.cls
' Добавление публичных свойств для получения статистики

' Количество доступных объектов в пуле
Public Property Get AvailableObjectCount() As Long
    AvailableObjectCount = m_AvailableCommands.Count
End Property

' Количество используемых объектов
Public Property Get InUseObjectCount() As Long
    InUseObjectCount = m_InUseCommands.Count
End Property

' Максимальный размер пула
Public Property Get MaxPoolSize() As Long
    MaxPoolSize = m_MaxPoolSize
End Property

' Время жизни объекта
Public Property Get CommandLifetime() As Long
    CommandLifetime = m_CommandLifetime
End Property

' Имя пула
Public Property Get PoolName() As String
    PoolName = m_PoolName
End Property
